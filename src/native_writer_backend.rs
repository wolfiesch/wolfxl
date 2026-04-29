//! Native xlsx writer pyclass — sole write-mode backend (W5+).
//!
//! Exposes the 22 pymethods that [`python::wolfxl::_backend.make_writer`]
//! constructs for every `Workbook()`. The 13 cell/format/structure
//! pymethods + `save()` drive [`wolfxl_writer::Workbook`] and emit a
//! complete xlsx via [`wolfxl_writer::emit_xlsx`]. The legacy
//! `rust_xlsxwriter`-backed sibling pyclass was removed in W5; the
//! payload-shape contract documented below is preserved verbatim for
//! Python-side compatibility.
//!
//! # Why mirror oracle exactly
//!
//! The Python flush path (`_flush_to_writer`) calls these methods with
//! payloads built by `python_value_to_payload`, `font_to_format_dict`,
//! etc. Those builders are oracle-shaped — keys, types, and value
//! coercions match what the oracle consumes. This file consumes the
//! same dicts so no Python-side change is required to drive native.
//!
//! # Style-merge limitation (4B follow-up)
//!
//! Calling `write_cell_format` after `write_cell_value` on the same
//! cell **replaces** the style_id with one freshly interned from the
//! format dict alone. Calling `write_cell_border` after that replaces
//! again. Oracle merges format + border at save time because it stores
//! them in separate HashMaps; native interns eagerly per call. In
//! practice the Python flush path always calls `write_cell_format`
//! and `write_cell_border` together (or only one), so the smoke test
//! does not hit the merge path. Documented gap; fix in 4B by exposing
//! a `StylesBuilder::lookup_format(style_id)` reverse query.

use std::fs;

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use wolfxl_writer::model::{FormatSpec, Worksheet, WriteCellValue};
use wolfxl_writer::refs;
use wolfxl_writer::Workbook;

use crate::native_writer_autofilter::install_autofilter;
use crate::native_writer_cells::{
    array_formula_payload_to_write_cell_value, payload_to_write_cell_value,
    raw_python_to_write_cell_value,
};
use crate::native_writer_charts::parse_chart_dict;
use crate::native_writer_formats::{intern_border_only, intern_format_from_dict};
use crate::native_writer_images::dict_to_sheet_image;
use crate::native_writer_rich_text::py_runs_to_rust_writer;
use crate::native_writer_sheet_features::{
    dict_to_comment, dict_to_conditional_format, dict_to_data_validation, dict_to_hyperlink,
    dict_to_table, unwrap_optional_wrapper,
};
use crate::native_writer_sheet_state::apply_freeze_panes;
use crate::native_writer_workbook_metadata::{
    dict_to_defined_name, dict_to_doc_properties, dict_to_workbook_security,
};

// ---------------------------------------------------------------------------
// PyClass
// ---------------------------------------------------------------------------

#[pyclass(unsendable)]
pub struct NativeWorkbook {
    inner: Workbook,
    saved: bool,
}

fn parse_a1_to_row_col(a1: &str) -> PyResult<(u32, u32)> {
    let cleaned = a1.replace('$', "");
    refs::parse_a1(&cleaned)
        .ok_or_else(|| PyValueError::new_err(format!("Invalid A1 reference: {a1}")))
}

fn require_sheet<'wb>(wb: &'wb mut Workbook, name: &str) -> PyResult<&'wb mut Worksheet> {
    wb.sheet_mut_by_name(name)
        .ok_or_else(|| PyValueError::new_err(format!("Unknown sheet: {name}")))
}

// ---------------------------------------------------------------------------
// PyMethods
// ---------------------------------------------------------------------------

#[pymethods]
impl NativeWorkbook {
    #[new]
    pub fn new() -> Self {
        Self {
            inner: Workbook::new(),
            saved: false,
        }
    }

    pub fn add_sheet(&mut self, name: &str) -> PyResult<()> {
        // Mirror oracle's idempotent semantic: re-adding an existing sheet is a no-op.
        if self.inner.sheet_by_name(name).is_some() {
            return Ok(());
        }
        self.inner.add_sheet(Worksheet::new(name));
        Ok(())
    }

    pub fn rename_sheet(&mut self, old_name: &str, new_name: &str) -> PyResult<()> {
        self.inner
            .rename_sheet(old_name, new_name.to_string())
            .map_err(PyValueError::new_err)
    }

    pub fn move_sheet(&mut self, name: &str, offset: isize) -> PyResult<()> {
        self.inner
            .move_sheet(name, offset)
            .map_err(PyValueError::new_err)
    }

    pub fn write_cell_value(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let (row, col) = parse_a1_to_row_col(a1)?;
        let value = payload_to_write_cell_value(payload)?;

        // If the value is a date/datetime and no number_format has been
        // attached yet, apply the oracle's defaults on the cell's style.
        let default_nf = match (
            payload
                .cast::<PyDict>()
                .ok()
                .and_then(|d| d.get_item("type").ok().flatten())
                .and_then(|v| v.extract::<String>().ok())
                .as_deref(),
            &value,
        ) {
            (Some("date"), WriteCellValue::DateSerial(_)) => Some("yyyy-mm-dd"),
            (Some("datetime"), WriteCellValue::DateSerial(_)) => Some("yyyy-mm-dd hh:mm:ss"),
            _ => None,
        };

        let style_id = if let Some(nf) = default_nf {
            let spec = FormatSpec {
                number_format: Some(nf.to_string()),
                ..Default::default()
            };
            Some(self.inner.styles.intern_format(&spec))
        } else {
            None
        };

        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.write_cell(row, col, value, style_id);
        Ok(())
    }

    /// Sprint Ι Pod-α: write a rich-text inline-string cell.
    ///
    /// `runs` is a Python list of ``(text, font_dict_or_None)``
    /// tuples — same shape as the patcher's
    /// ``queue_rich_text_value``. The native writer emits an
    /// inline-string `<c t="inlineStr">` cell so the SST stays
    /// untouched.
    pub fn write_cell_rich_text(
        &mut self,
        sheet: &str,
        a1: &str,
        runs: &Bound<'_, pyo3::types::PyList>,
    ) -> PyResult<()> {
        use wolfxl_writer::model::cell::WriteCellValue;
        let (row, col) = parse_a1_to_row_col(a1)?;
        let parsed = py_runs_to_rust_writer(runs)?;
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.write_cell(row, col, WriteCellValue::InlineRichText(parsed), None);
        Ok(())
    }

    /// RFC-057 (Sprint Ο Pod 1C): write an array-formula / data-table
    /// formula / spill-child cell.
    ///
    /// `payload` mirrors the patcher's ``queue_array_formula`` shape:
    ///   - ``{"kind": "array", "ref": "A1:A10", "text": "B1:B10*2"}``
    ///   - ``{"kind": "data_table", "ref": "B2:F11", "ca": false,
    ///        "dt2D": true, "dtr": false, "r1": "A1", "r2": "A2"}``
    ///   - ``{"kind": "spill_child"}``
    pub fn write_cell_array_formula(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let (row, col) = parse_a1_to_row_col(a1)?;
        let value = array_formula_payload_to_write_cell_value(payload)?;
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.write_cell(row, col, value, None);
        Ok(())
    }

    /// Bulk-write a rectangular grid of values starting at `start_a1`.
    pub fn write_sheet_values(
        &mut self,
        sheet: &str,
        start_a1: &str,
        values: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let (base_row, base_col) = parse_a1_to_row_col(start_a1)?;

        let ws = require_sheet(&mut self.inner, sheet)?;
        let rows: Vec<Bound<'_, PyAny>> = values.extract()?;
        for (ri, row_obj) in rows.iter().enumerate() {
            let cols: Vec<Bound<'_, PyAny>> = row_obj.extract()?;
            for (ci, val) in cols.iter().enumerate() {
                if val.is_none() {
                    continue;
                }
                let row = base_row + ri as u32;
                let col = base_col + ci as u32;
                if let Some(value) = raw_python_to_write_cell_value(val)? {
                    ws.write_cell(row, col, value, None);
                }
                // else: skip silently like the oracle does.
            }
        }
        Ok(())
    }

    pub fn write_cell_format(
        &mut self,
        sheet: &str,
        a1: &str,
        format_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let (row, col) = parse_a1_to_row_col(a1)?;
        let dict = format_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("format_dict must be a dict"))?;

        let style_id = intern_format_from_dict(&mut self.inner, dict)?;

        let ws = require_sheet(&mut self.inner, sheet)?;
        let cell = ws
            .rows
            .entry(row)
            .or_default()
            .cells
            .entry(col)
            .or_insert_with(|| wolfxl_writer::model::WriteCell {
                value: WriteCellValue::Blank,
                style_id: None,
            });
        cell.style_id = Some(style_id);
        Ok(())
    }

    pub fn write_cell_border(
        &mut self,
        sheet: &str,
        a1: &str,
        border_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let (row, col) = parse_a1_to_row_col(a1)?;
        let dict = border_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("border_dict must be a dict"))?;

        let style_id = intern_border_only(&mut self.inner, dict)?;

        let ws = require_sheet(&mut self.inner, sheet)?;
        let cell = ws
            .rows
            .entry(row)
            .or_default()
            .cells
            .entry(col)
            .or_insert_with(|| wolfxl_writer::model::WriteCell {
                value: WriteCellValue::Blank,
                style_id: None,
            });
        cell.style_id = Some(style_id);
        Ok(())
    }

    pub fn write_sheet_formats(
        &mut self,
        sheet: &str,
        start_a1: &str,
        formats: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let (base_row, base_col) = parse_a1_to_row_col(start_a1)?;
        let rows: Vec<Bound<'_, PyAny>> = formats.extract()?;

        // Intern style ids first (need &mut self.inner), then write to sheet.
        let mut to_apply: Vec<(u32, u32, u32)> = Vec::new();
        for (ri, row_obj) in rows.iter().enumerate() {
            let cols: Vec<Bound<'_, PyAny>> = row_obj.extract()?;
            for (ci, val) in cols.iter().enumerate() {
                if val.is_none() {
                    continue;
                }
                let dict = val
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("format element must be dict or None"))?;
                if dict.is_empty() {
                    continue;
                }
                let row = base_row + ri as u32;
                let col = base_col + ci as u32;
                let style_id = intern_format_from_dict(&mut self.inner, dict)?;
                to_apply.push((row, col, style_id));
            }
        }

        let ws = require_sheet(&mut self.inner, sheet)?;
        for (row, col, style_id) in to_apply {
            let cell = ws
                .rows
                .entry(row)
                .or_default()
                .cells
                .entry(col)
                .or_insert_with(|| wolfxl_writer::model::WriteCell {
                    value: WriteCellValue::Blank,
                    style_id: None,
                });
            cell.style_id = Some(style_id);
        }
        Ok(())
    }

    pub fn write_sheet_borders(
        &mut self,
        sheet: &str,
        start_a1: &str,
        borders: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let (base_row, base_col) = parse_a1_to_row_col(start_a1)?;
        let rows: Vec<Bound<'_, PyAny>> = borders.extract()?;

        let mut to_apply: Vec<(u32, u32, u32)> = Vec::new();
        for (ri, row_obj) in rows.iter().enumerate() {
            let cols: Vec<Bound<'_, PyAny>> = row_obj.extract()?;
            for (ci, val) in cols.iter().enumerate() {
                if val.is_none() {
                    continue;
                }
                let dict = val
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("border element must be dict or None"))?;
                if dict.is_empty() {
                    continue;
                }
                let row = base_row + ri as u32;
                let col = base_col + ci as u32;
                let style_id = intern_border_only(&mut self.inner, dict)?;
                to_apply.push((row, col, style_id));
            }
        }

        let ws = require_sheet(&mut self.inner, sheet)?;
        for (row, col, style_id) in to_apply {
            let cell = ws
                .rows
                .entry(row)
                .or_default()
                .cells
                .entry(col)
                .or_insert_with(|| wolfxl_writer::model::WriteCell {
                    value: WriteCellValue::Blank,
                    style_id: None,
                });
            cell.style_id = Some(style_id);
        }
        Ok(())
    }

    pub fn set_row_height(&mut self, sheet: &str, row: u32, height: f64) -> PyResult<()> {
        // Python passes 1-based rows (openpyxl convention). The native
        // model is also 1-based, so no `saturating_sub(1)` like oracle.
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.set_row_height(row, height);
        Ok(())
    }

    pub fn set_column_width(&mut self, sheet: &str, col_str: &str, width: f64) -> PyResult<()> {
        let col = refs::letters_to_col(col_str)
            .ok_or_else(|| PyValueError::new_err(format!("Invalid column letter: {col_str}")))?;
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.set_column_width(col, width);
        Ok(())
    }

    pub fn merge_cells(&mut self, sheet: &str, range_str: &str) -> PyResult<()> {
        let cleaned = range_str.replace('$', "");
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.merge_cells(&cleaned).map_err(PyValueError::new_err)
    }

    /// Mirrors oracle: dict with keys `mode` ("freeze" | "split") and
    /// either `top_left_cell` (freeze) or `x_split` / `y_split` (split).
    /// Optional wrapper key `freeze` is also accepted.
    pub fn set_freeze_panes(&mut self, sheet: &str, settings: &Bound<'_, PyAny>) -> PyResult<()> {
        let ws = require_sheet(&mut self.inner, sheet)?;
        apply_freeze_panes(ws, settings)
    }

    pub fn save(&mut self, path: &str) -> PyResult<()> {
        if self.saved {
            return Err(PyValueError::new_err(
                "Workbook already saved (NativeWorkbook is consumed-on-save)",
            ));
        }
        // Mark consumed BEFORE emit/write so a panic in emit_xlsx or fs::write
        // leaves the workbook un-retryable on partially-mutated state.
        self.saved = true;
        let bytes = wolfxl_writer::emit_xlsx(&mut self.inner);
        fs::write(path, bytes)
            .map_err(|e| PyIOError::new_err(format!("failed to write {path}: {e}")))?;
        Ok(())
    }

    // =========================================================================
    // Wave 4B rich-feature pymethods — real implementations.
    // =========================================================================

    pub fn add_hyperlink(&mut self, sheet: &str, link_dict: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = link_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("link must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "hyperlink")?;
        let Some((a1, hyperlink)) = dict_to_hyperlink(&cfg)? else {
            return Ok(()); // silent no-op — match oracle
        };
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.hyperlinks.insert(a1, hyperlink);
        Ok(())
    }

    pub fn add_comment(&mut self, sheet: &str, comment_dict: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = comment_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("comment_dict must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "comment")?;
        // Borrow authors first (resolution must happen before re-borrowing inner for sheet).
        let Some((a1, comment)) = dict_to_comment(&cfg, &mut self.inner.comment_authors)? else {
            return Ok(()); // silent no-op — match oracle
        };
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.comments.insert(a1, comment);
        Ok(())
    }

    pub fn set_print_area(&mut self, sheet: &str, range_str: &str) -> PyResult<()> {
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.print_area = Some(range_str.to_string());
        Ok(())
    }

    /// Sprint Ο Pod 1A.5 (RFC-055) — install sheet-setup blocks on a
    /// write-mode sheet. Consumes the §10 dict shape produced by
    /// `Worksheet.to_rust_setup_dict()` and copies the parsed specs
    /// onto the worksheet model. The native writer's emit pass then
    /// renders them at the correct CT_Worksheet positions.
    pub fn set_sheet_setup_native(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let specs = crate::wolfxl::sheet_setup::parse_sheet_setup_payload(payload)?;
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.views = specs.sheet_view;
        ws.protection = specs.sheet_protection;
        ws.page_margins = specs.page_margins;
        ws.page_setup = specs.page_setup;
        ws.header_footer = specs.header_footer;
        // print_titles is workbook-scope; routed via definedNames
        // queue on the Python side, not here.
        Ok(())
    }

    /// Sprint Π Pod Π-α (RFC-062) — install page-breaks +
    /// sheetFormatPr blocks on a write-mode sheet. Consumes the
    /// merged §10 dict shape produced by
    /// ``Worksheet.to_rust_page_breaks_dict()`` +
    /// ``Worksheet.to_rust_sheet_format_dict()``.
    pub fn set_page_breaks_native(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let queued = crate::wolfxl::page_breaks::parse_page_breaks_payload(payload)?;
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.row_breaks = queued.row_breaks;
        ws.col_breaks = queued.col_breaks;
        ws.sheet_format = queued.sheet_format;
        Ok(())
    }

    /// Sprint Ο Pod 1B (RFC-056) — install an autoFilter on a write-
    /// mode sheet. Takes the §10 dict, pre-emits the `<autoFilter>`
    /// block via `wolfxl_autofilter::emit::emit`, and evaluates the
    /// filter against the worksheet's in-memory cells to stamp
    /// `row.hidden = true` on rows excluded by the filter.
    pub fn set_autofilter_native(&mut self, sheet: &str, d: &Bound<'_, PyDict>) -> PyResult<()> {
        let ws = require_sheet(&mut self.inner, sheet)?;
        install_autofilter(ws, d)
    }

    pub fn add_conditional_format(
        &mut self,
        sheet: &str,
        rule_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let dict = rule_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("rule must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "cf_rule")?;
        // Resolve CF (may intern a dxf — borrows styles) before borrowing sheet.
        let Some(cf) = dict_to_conditional_format(&cfg, &mut self.inner.styles)? else {
            return Ok(()); // silent no-op — match oracle
        };
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.conditional_formats.push(cf);
        Ok(())
    }

    pub fn add_data_validation(
        &mut self,
        sheet: &str,
        validation_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let dict = validation_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("validation must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "validation")?;
        let Some(dv) = dict_to_data_validation(&cfg)? else {
            return Ok(()); // silent no-op — match oracle
        };
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.validations.push(dv);
        Ok(())
    }

    pub fn add_named_range(&mut self, sheet: &str, named_range: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = named_range
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("named_range must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "named_range")?;
        // sheet_index_by_name borrows &self.inner immutably — do it before any &mut borrow.
        let Some(dn) = dict_to_defined_name(&self.inner, sheet, &cfg)? else {
            return Ok(()); // silent no-op — match oracle
        };
        self.inner.defined_names.push(dn);
        Ok(())
    }

    pub fn add_table(&mut self, sheet: &str, table: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = table
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("table must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "table")?;
        let Some(tbl) = dict_to_table(&cfg)? else {
            return Ok(()); // silent no-op — match oracle
        };
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.tables.push(tbl);
        Ok(())
    }

    pub fn set_properties(&mut self, props: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = props
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("props must be a dict"))?;
        let doc_props = dict_to_doc_properties(dict)?;
        self.inner.set_doc_props(doc_props);
        Ok(())
    }

    /// Sprint Ο Pod 1D (RFC-058) — set workbook-level security.
    ///
    /// Accepts the §10 flat dict shape (`workbook_protection` and
    /// `file_sharing` top-level keys, either of which may be `None`).
    /// Replaces any previously-set security (last writer wins).
    pub fn set_workbook_security(&mut self, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = payload
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("workbook security payload must be a dict"))?;
        let security = dict_to_workbook_security(dict)?;
        self.inner.security = security;
        Ok(())
    }

    /// Sprint Λ Pod-β (RFC-045) — queue an image onto a sheet.
    ///
    /// `image_dict` shape (built by Python's
    /// ``Worksheet.add_image``):
    ///
    /// ```python
    /// {
    ///     "data": <bytes>,
    ///     "ext": "png" | "jpeg" | "gif" | "bmp",
    ///     "width": int,
    ///     "height": int,
    ///     "anchor": {
    ///         "type": "one_cell" | "two_cell" | "absolute",
    ///         # one_cell: from_col, from_row, from_col_off, from_row_off (0-based + EMU)
    ///         # two_cell: + to_col, to_row, to_col_off, to_row_off, edit_as
    ///         # absolute: x_emu, y_emu, cx_emu, cy_emu
    ///         ...
    ///     },
    /// }
    /// ```
    pub fn add_image(&mut self, sheet: &str, image_dict: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = image_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("image must be a dict"))?;
        let img = dict_to_sheet_image(dict)?;
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.images.push(img);
        Ok(())
    }

    /// Sprint Μ Pod-α (RFC-046) — queue a chart onto a sheet.
    ///
    /// `chart_dict` shape (built by Python's
    /// ``Worksheet.add_chart``):
    ///
    /// ```python
    /// {
    ///     "kind": "bar" | "line" | "pie" | "doughnut" | "area"
    ///           | "scatter" | "bubble" | "radar",
    ///     "anchor": { "type": "one_cell" | "two_cell" | "absolute", ... },
    ///     "title": { "runs": [{"text": "Sales", "bold": true, ...}],
    ///                "overlay": false } | None,
    ///     "legend": { "position": "r"|"l"|"t"|"b"|"tr",
    ///                 "overlay": false, ... } | None,
    ///     "x_axis": { "kind": "category"|"value"|"date"|"series",
    ///                 "ax_id": 10, "cross_ax": 100,
    ///                 "ax_pos": "b"|"t"|"l"|"r",
    ///                 "orientation": "minMax"|"maxMin",
    ///                 "major_gridlines": false, "minor_gridlines": false,
    ///                 "major_tick_mark": "none"|"in"|"out"|"cross",
    ///                 "title": {...}, "number_format": "0.00",
    ///                 # ValueAxis only:
    ///                 "min": 0.0, "max": 100.0,
    ///                 "major_unit": 10.0, "minor_unit": 1.0,
    ///                 "crosses": "autoZero"|"min"|"max",
    ///                 # CategoryAxis only:
    ///                 "lbl_offset": 100, "lbl_algn": "ctr",
    ///                 # DateAxis only:
    ///                 "base_time_unit": "days"|"months"|"years",
    ///               } | None,
    ///     "y_axis": {...} | None,
    ///     "series": [
    ///         { "idx": 0, "order": 0,
    ///           "title": {"strRef": {"sheet": "Sheet1", "range": "B1"}}
    ///                 | {"literal": "My Series"} | None,
    ///           "categories": {"sheet": "Sheet1", "range": "A2:A6"} | None,
    ///           "values": {"sheet": "Sheet1", "range": "B2:B6"} | None,
    ///           "x_values": {...} | None,
    ///           "bubble_size": {...} | None,
    ///           "graphical_properties": {
    ///               "line_color": "FF000000", "line_width_emu": 12700,
    ///               "line_dash": "solid", "fill_color": "FF0000FF",
    ///               "no_fill": false, "no_line": false,
    ///           } | None,
    ///           "marker": { "symbol": "circle"|"square"|...,
    ///                       "size": 7, "graphical_properties": {...} } | None,
    ///           "data_labels": { "show_val": true, "show_cat_name": true,
    ///                            "show_ser_name": false, "show_percent": false,
    ///                            "show_legend_key": false,
    ///                            "show_bubble_size": false,
    ///                            "position": "outEnd",
    ///                            "number_format": "0.00",
    ///                            "separator": "," } | None,
    ///           "error_bars": [
    ///               { "bar_type": "plus"|"minus"|"both",
    ///                 "val_type": "cust"|"fixedVal"|"percentage"
    ///                          |"stdDev"|"stdErr",
    ///                 "value": 1.5, "no_end_cap": false }
    ///           ],
    ///           "trendlines": [
    ///               { "kind": "linear"|"log"|"power"|"exp"
    ///                       |"poly"|"movingAvg",
    ///                 "order": 2, "period": 3, "forward": 1.0,
    ///                 "backward": 1.0, "display_equation": true,
    ///                 "display_r_squared": true, "name": "fit" }
    ///           ],
    ///           "smooth": true, "invert_if_negative": false,
    ///         },
    ///     ],
    ///     "plot_visible_only": true, "display_blanks_as": "gap"|"span"|"zero",
    ///     "vary_colors": true,
    ///     # Bar:
    ///     "bar_dir": "col"|"bar", "grouping": "clustered"|"stacked"
    ///                                |"percentStacked"|"standard",
    ///     "gap_width": 150, "overlap": -50,
    ///     # Doughnut/Pie:
    ///     "hole_size": 50, "first_slice_ang": 0,
    ///     # Scatter:
    ///     "scatter_style": "line"|"lineMarker"|"marker"|"smooth"
    ///                    |"smoothMarker"|"none",
    ///     # Radar:
    ///     "radar_style": "standard"|"marker"|"filled",
    ///     # Bubble:
    ///     "bubble3d": false, "bubble_scale": 100,
    ///     "show_neg_bubbles": false,
    ///     "smoothing": false, "style": 1,
    /// }
    /// ```
    ///
    /// `anchor_a1` is the A1 reference of the top-left cell where the
    /// chart should be anchored (e.g. `"D2"`). Pod-β-style anchor
    /// dicts inside `chart_dict["anchor"]` override this if present.
    pub fn add_chart_native(
        &mut self,
        sheet: &str,
        chart_dict: &Bound<'_, PyAny>,
        anchor_a1: &str,
    ) -> PyResult<()> {
        let dict = chart_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("chart must be a dict"))?;
        let chart = parse_chart_dict(dict, anchor_a1)?;
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.charts.push(chart);
        Ok(())
    }
}

/// Sprint Ο Pod 1D (RFC-058 §10) — render the workbook security dict to
/// the two XML fragments (`workbookProtection`, `fileSharing`).
///
/// Returns `(workbook_protection_bytes, file_sharing_bytes)`. Either
/// element may be empty bytes when the corresponding source dict is
/// `None` or all-default. Callers splice each fragment at the matching
/// canonical position in `xl/workbook.xml`.
#[pyfunction]
pub fn serialize_workbook_security_dict(
    payload: &Bound<'_, PyDict>,
) -> PyResult<(Vec<u8>, Vec<u8>)> {
    use wolfxl_writer::parse::workbook_security::{emit_file_sharing, emit_workbook_protection};
    let security = dict_to_workbook_security(payload)?;
    let prot_bytes = security
        .workbook_protection
        .as_ref()
        .map(emit_workbook_protection)
        .unwrap_or_default();
    let share_bytes = security
        .file_sharing
        .as_ref()
        .map(emit_file_sharing)
        .unwrap_or_default();
    Ok((prot_bytes, share_bytes))
}
