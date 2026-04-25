//! Native xlsx writer pyclass — Wave 4A.
//!
//! Sibling to [`crate::rust_xlsxwriter_backend::RustXlsxWriterBook`]. Mirrors
//! the same 22 pymethods so [`python::wolfxl::_backend.make_writer`] can swap
//! between the two without any other Python-side changes. The 13
//! cell/format/structure pymethods + `save()` are real implementations
//! that drive [`wolfxl_writer::Workbook`] and emit a complete xlsx via
//! [`wolfxl_writer::emit_xlsx`]. The 8 rich-feature pymethods raise
//! `NotImplementedError` until Wave 4B fills them in.
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

use pyo3::exceptions::{PyIOError, PyNotImplementedError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use wolfxl_writer::model::date::{date_to_excel_serial, datetime_to_excel_serial};
use wolfxl_writer::model::{
    AlignmentSpec, BorderSideSpec, BorderSpec, FillSpec, FontSpec, FormatSpec, Worksheet,
    WriteCellValue,
};
use wolfxl_writer::refs;
use wolfxl_writer::Workbook;

use crate::util::{parse_iso_date, parse_iso_datetime};

// ---------------------------------------------------------------------------
// PyClass
// ---------------------------------------------------------------------------

#[pyclass(unsendable)]
pub struct NativeWorkbook {
    inner: Workbook,
    saved: bool,
}

// ---------------------------------------------------------------------------
// Color helpers
// ---------------------------------------------------------------------------

/// Normalize a Python-side color string to OOXML's 8-char ARGB form
/// (`"FFRRGGBB"`, uppercase, no `#`).
///
/// Accepts:
///
/// - `"#RRGGBB"`        → prefix alpha `FF`, uppercase
/// - `"RRGGBB"`         → prefix alpha `FF`, uppercase
/// - `"#AARRGGBB"`      → strip `#`, uppercase
/// - `"AARRGGBB"`       → uppercase as-is
/// - `"#RGB"` / `"RGB"` → expand each digit (CSS shorthand), then `FF` prefix
///
/// Returns `None` for any other shape (empty, non-hex, wrong length).
fn parse_hex_color(input: &str) -> Option<String> {
    let s = input.strip_prefix('#').unwrap_or(input);
    let upper: String = s.chars().map(|c| c.to_ascii_uppercase()).collect();

    // Validate hex digits up front; cheaper than a per-char check below.
    if !upper.chars().all(|c| c.is_ascii_hexdigit()) {
        return None;
    }

    match upper.len() {
        3 => {
            // CSS shorthand: each digit is doubled.
            let mut expanded = String::with_capacity(8);
            expanded.push_str("FF");
            for ch in upper.chars() {
                expanded.push(ch);
                expanded.push(ch);
            }
            Some(expanded)
        }
        6 => Some(format!("FF{upper}")),
        8 => Some(upper),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Python value → WriteCellValue
// ---------------------------------------------------------------------------

/// Convert oracle-shape cell payload dict into a `WriteCellValue`.
///
/// The Python flush path passes `python_value_to_payload(value)` which
/// always returns a dict like `{"type": "string", "value": "x"}` (see
/// `python/wolfxl/_cell.py:518`). We extract the `type` and decode
/// `value` / `formula` per the oracle's mapping in `write_cell` at
/// `src/rust_xlsxwriter_backend.rs:464`. Stays in lockstep with the
/// oracle so flushed payloads work identically under both backends.
fn payload_to_write_cell_value(payload: &Bound<'_, PyAny>) -> PyResult<WriteCellValue> {
    let dict = payload
        .downcast::<PyDict>()
        .map_err(|_| PyValueError::new_err("payload must be a dict"))?;

    let type_str: String = dict
        .get_item("type")?
        .ok_or_else(|| PyValueError::new_err("payload missing 'type'"))?
        .extract()?;

    // The oracle accepts strings or numbers for `value` and converts to
    // String; we mirror that so flush-path round-trips stay identical.
    let value_str: Option<String> = dict.get_item("value")?.and_then(|v| {
        v.extract::<String>().ok().or_else(|| {
            v.extract::<f64>()
                .map(|n| n.to_string())
                .ok()
                .or_else(|| v.extract::<bool>().map(|b| b.to_string()).ok())
        })
    });
    let formula_str: Option<String> = dict.get_item("formula")?.and_then(|v| v.extract().ok());

    match type_str.as_str() {
        "blank" => Ok(WriteCellValue::Blank),
        "string" => Ok(WriteCellValue::String(value_str.unwrap_or_default())),
        "number" => {
            let n: f64 = value_str
                .as_deref()
                .unwrap_or("0")
                .parse()
                .map_err(|_| PyValueError::new_err("number parse failed"))?;
            Ok(WriteCellValue::Number(n))
        }
        "boolean" => {
            let b = value_str.as_deref().map(parse_python_bool).unwrap_or(false);
            Ok(WriteCellValue::Boolean(b))
        }
        "formula" => {
            let expr = formula_str
                .or(value_str)
                .map(|s| s.trim_start_matches('=').to_string())
                .unwrap_or_default();
            Ok(WriteCellValue::Formula {
                expr,
                result: None,
            })
        }
        "error" => {
            // Native model has no Error variant yet — fall back to the
            // raw token as a string so the cell isn't lost. Mirrors the
            // oracle's last-resort branch (line 530-534).
            Ok(WriteCellValue::String(value_str.unwrap_or_default()))
        }
        "date" => {
            let s = value_str.unwrap_or_default();
            if let Some(d) = parse_iso_date(&s) {
                if let Some(serial) = date_to_excel_serial(d) {
                    return Ok(WriteCellValue::DateSerial(serial));
                }
            }
            Ok(WriteCellValue::String(s))
        }
        "datetime" => {
            let s = value_str.unwrap_or_default();
            if let Some(dt) = parse_iso_datetime(&s) {
                if let Some(serial) = datetime_to_excel_serial(dt) {
                    return Ok(WriteCellValue::DateSerial(serial));
                }
            }
            Ok(WriteCellValue::String(s))
        }
        other => Err(PyValueError::new_err(format!(
            "Unsupported cell type: {other}"
        ))),
    }
}

fn parse_python_bool(s: &str) -> bool {
    matches!(
        s.trim().to_ascii_lowercase().as_str(),
        "true" | "1" | "t" | "yes" | "y"
    )
}

/// Coerce a raw Python value (from `write_sheet_values`'s 2-D list) to a
/// `WriteCellValue`. Mirrors the oracle's order-of-attempts at
/// `src/rust_xlsxwriter_backend.rs:1417`, but fixes a subtle bug: the
/// oracle tries `f64` before `bool`, so `True`/`False` (which extract
/// as `1.0`/`0.0`) silently become numbers. The Python flush path
/// avoids this by routing booleans through `write_cell_value` instead,
/// but we tighten the rule here for correctness — bool first.
fn raw_python_to_write_cell_value(value: &Bound<'_, PyAny>) -> Option<WriteCellValue> {
    if value.is_none() {
        return None;
    }
    // Boolean check via `is_instance_of` (rather than `extract`) since
    // `extract::<bool>()` would succeed on `0`/`1` ints too.
    let py = value.py();
    let bool_type = py.get_type::<pyo3::types::PyBool>();
    if value.is_instance(&bool_type).unwrap_or(false) {
        let b = value.extract::<bool>().ok()?;
        return Some(WriteCellValue::Boolean(b));
    }
    if let Ok(i) = value.extract::<i64>() {
        return Some(WriteCellValue::Number(i as f64));
    }
    if let Ok(f) = value.extract::<f64>() {
        return Some(WriteCellValue::Number(f));
    }
    if let Ok(s) = value.extract::<String>() {
        if s.starts_with('=') {
            return Some(WriteCellValue::Formula {
                expr: s.trim_start_matches('=').to_string(),
                result: None,
            });
        }
        return Some(WriteCellValue::String(s));
    }
    // Datetime / date — best-effort via isoformat() if exposed.
    if let Ok(iso) = value.call_method0("isoformat") {
        if let Ok(s) = iso.extract::<String>() {
            if let Some(dt) = parse_iso_datetime(&s) {
                if let Some(serial) = datetime_to_excel_serial(dt) {
                    return Some(WriteCellValue::DateSerial(serial));
                }
            }
            if let Some(d) = parse_iso_date(&s) {
                if let Some(serial) = date_to_excel_serial(d) {
                    return Some(WriteCellValue::DateSerial(serial));
                }
            }
        }
    }
    None
}

// ---------------------------------------------------------------------------
// Format / border dict → spec
// ---------------------------------------------------------------------------

/// Read the format dict the Python side builds via
/// `python/wolfxl/_cell.py:font_to_format_dict` etc. and turn it into a
/// `FormatSpec`. Keys mirrored from oracle's `extract_format_fields`
/// (line 352) and `build_format` (line 239) — same names, same coercions.
fn dict_to_format_spec(dict: &Bound<'_, PyDict>) -> PyResult<FormatSpec> {
    let mut spec = FormatSpec::default();

    let mut font = FontSpec::default();
    let mut font_touched = false;

    if let Some(v) = dict.get_item("bold")? {
        if let Ok(b) = v.extract::<bool>() {
            font.bold = b;
            font_touched |= b;
        }
    }
    if let Some(v) = dict.get_item("italic")? {
        if let Ok(b) = v.extract::<bool>() {
            font.italic = b;
            font_touched |= b;
        }
    }
    if let Some(v) = dict.get_item("underline")? {
        // Python sends a string ("single", "double", ...). Coerce to
        // boolean: if the field is present and non-empty, set underline.
        if let Ok(s) = v.extract::<String>() {
            if !s.is_empty() {
                font.underline = true;
                font_touched = true;
            }
        } else if let Ok(b) = v.extract::<bool>() {
            font.underline = b;
            font_touched |= b;
        }
    }
    if let Some(v) = dict.get_item("strikethrough")? {
        if let Ok(b) = v.extract::<bool>() {
            font.strikethrough = b;
            font_touched |= b;
        }
    }
    if let Some(v) = dict.get_item("font_name")? {
        if let Ok(s) = v.extract::<String>() {
            if !s.is_empty() {
                font.name = Some(s);
                font_touched = true;
            }
        }
    }
    if let Some(v) = dict.get_item("font_size")? {
        if let Ok(f) = v.extract::<f64>() {
            // FontSpec stores u32 points. Clamp non-negative; the oracle
            // (rust_xlsxwriter) accepts f64 but our schema is whole points.
            if f.is_finite() && f >= 0.0 {
                font.size = Some(f.round() as u32);
                font_touched = true;
            }
        }
    }
    if let Some(v) = dict.get_item("font_color")? {
        if let Ok(s) = v.extract::<String>() {
            if let Some(rgb) = parse_hex_color(&s) {
                font.color_rgb = Some(rgb);
                font_touched = true;
            }
        }
    }
    if font_touched {
        spec.font = Some(font);
    }

    // Fill — only `bg_color` is wired through the Python flush path
    // (`fill_to_format_dict` only emits this one key).
    if let Some(v) = dict.get_item("bg_color")? {
        if let Ok(s) = v.extract::<String>() {
            if let Some(rgb) = parse_hex_color(&s) {
                spec.fill = Some(FillSpec {
                    pattern_type: "solid".to_string(),
                    fg_color_rgb: Some(rgb.clone()),
                    bg_color_rgb: Some(rgb),
                });
            }
        }
    }

    if let Some(v) = dict.get_item("number_format")? {
        if let Ok(s) = v.extract::<String>() {
            if !s.is_empty() {
                spec.number_format = Some(s);
            }
        }
    }

    // Alignment — `h_align`, `v_align`, `wrap`, `rotation`, `indent`.
    let mut align = AlignmentSpec::default();
    let mut align_touched = false;
    if let Some(v) = dict.get_item("h_align")? {
        if let Ok(s) = v.extract::<String>() {
            if !s.is_empty() {
                align.horizontal = Some(s);
                align_touched = true;
            }
        }
    }
    if let Some(v) = dict.get_item("v_align")? {
        if let Ok(s) = v.extract::<String>() {
            if !s.is_empty() {
                align.vertical = Some(s);
                align_touched = true;
            }
        }
    }
    if let Some(v) = dict.get_item("wrap")? {
        if let Ok(b) = v.extract::<bool>() {
            align.wrap_text = b;
            align_touched |= b;
        }
    }
    if let Some(v) = dict.get_item("rotation")? {
        if let Ok(i) = v.extract::<i32>() {
            // OOXML rotation field is u32 (0-180; 255 = vertical text).
            // Oracle clamps to i16 then assigns; we follow suit minus the
            // bound check since u32 fits all valid values.
            if i >= 0 {
                align.text_rotation = i as u32;
                align_touched = true;
            }
        }
    }
    if let Some(v) = dict.get_item("indent")? {
        if let Ok(i) = v.extract::<i32>() {
            if i >= 0 {
                align.indent = i as u32;
                align_touched = true;
            }
        }
    }
    if align_touched {
        spec.alignment = Some(align);
    }

    Ok(spec)
}

/// Edge dict (`{"style": "...", "color": "..."}`) → BorderSideSpec.
/// Returns `(side_spec, touched)` so the caller can decide whether the
/// border block contributes anything to the final FormatSpec.
fn edge_to_side_spec(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<(BorderSideSpec, bool)> {
    let mut side = BorderSideSpec::default();
    let mut touched = false;

    let Some(sub) = dict.get_item(key)? else {
        return Ok((side, false));
    };
    let Ok(d) = sub.downcast::<PyDict>() else {
        return Ok((side, false));
    };

    if let Some(v) = d.get_item("style")? {
        if let Ok(s) = v.extract::<String>() {
            if !s.is_empty() {
                side.style = Some(s);
                touched = true;
            }
        }
    }
    if let Some(v) = d.get_item("color")? {
        if let Ok(s) = v.extract::<String>() {
            if let Some(rgb) = parse_hex_color(&s) {
                side.color_rgb = Some(rgb);
                touched = true;
            }
        }
    }
    Ok((side, touched))
}

/// Read the border dict the Python side builds via
/// `border_to_rust_dict` (`python/wolfxl/_cell.py:581`) and turn it into
/// a `BorderSpec`. Edge keys: `top`, `bottom`, `left`, `right`,
/// `diagonal_up`, `diagonal_down`. Each edge value is a sub-dict with
/// `style` and (optional) `color`.
fn dict_to_border_spec(dict: &Bound<'_, PyDict>) -> PyResult<BorderSpec> {
    let mut border = BorderSpec::default();
    let mut any = false;

    let (top, t1) = edge_to_side_spec(dict, "top")?;
    let (bottom, t2) = edge_to_side_spec(dict, "bottom")?;
    let (left, t3) = edge_to_side_spec(dict, "left")?;
    let (right, t4) = edge_to_side_spec(dict, "right")?;
    let (diag_up, t5) = edge_to_side_spec(dict, "diagonal_up")?;
    let (diag_down, t6) = edge_to_side_spec(dict, "diagonal_down")?;

    if t1 {
        border.top = top;
        any = true;
    }
    if t2 {
        border.bottom = bottom;
        any = true;
    }
    if t3 {
        border.left = left;
        any = true;
    }
    if t4 {
        border.right = right;
        any = true;
    }
    if t5 || t6 {
        // The native model has a single `diagonal` slot — preferring
        // `down` matches oracle's `build_format` (line 322-330) which
        // applies `down` second when both are set.
        if t6 {
            border.diagonal = diag_down;
        } else {
            border.diagonal = diag_up;
        }
        border.diagonal_up = t5;
        border.diagonal_down = t6;
        any = true;
    }

    let _ = any; // returned spec is always meaningful even if all edges empty
    Ok(border)
}

// ---------------------------------------------------------------------------
// Helpers tying conversions to the model
// ---------------------------------------------------------------------------

fn intern_format_from_dict(
    wb: &mut Workbook,
    dict: &Bound<'_, PyDict>,
) -> PyResult<u32> {
    let spec = dict_to_format_spec(dict)?;
    Ok(wb.styles.intern_format(&spec))
}

fn intern_border_only(wb: &mut Workbook, dict: &Bound<'_, PyDict>) -> PyResult<u32> {
    let border = dict_to_border_spec(dict)?;
    let spec = FormatSpec {
        border: Some(border),
        ..Default::default()
    };
    Ok(wb.styles.intern_format(&spec))
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
                .downcast::<PyDict>()
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
                if let Some(value) = raw_python_to_write_cell_value(val) {
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
            .downcast::<PyDict>()
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
            .downcast::<PyDict>()
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
                    .downcast::<PyDict>()
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
                    .downcast::<PyDict>()
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
        ws.merge_cells(&cleaned)
            .map_err(PyValueError::new_err)
    }

    /// Mirrors oracle: dict with keys `mode` ("freeze" | "split") and
    /// either `top_left_cell` (freeze) or `x_split` / `y_split` (split).
    /// Optional wrapper key `freeze` is also accepted.
    pub fn set_freeze_panes(&mut self, sheet: &str, settings: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = settings
            .downcast::<PyDict>()
            .map_err(|_| PyValueError::new_err("settings must be a dict"))?;

        let inner: Option<Bound<'_, PyAny>> = dict.get_item("freeze")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let mode: String = cfg
            .get_item("mode")?
            .and_then(|v| v.extract::<String>().ok())
            .unwrap_or_else(|| "freeze".to_string());

        let ws = require_sheet(&mut self.inner, sheet)?;

        if mode == "freeze" {
            let top_left: Option<String> = cfg
                .get_item("top_left_cell")?
                .and_then(|v| v.extract::<String>().ok());
            if let Some(cell) = top_left {
                let (row, col) = parse_a1_to_row_col(&cell)?;
                // freeze_row/col semantics in the model: rows above
                // `freeze_row` and columns left of `freeze_col` stay
                // pinned; the top-left cell's (row, col) IS the
                // freeze split point.
                ws.set_freeze(row, col, Some((row, col)));
            }
        } else if mode == "split" {
            let x_split: f64 = cfg
                .get_item("x_split")?
                .and_then(|v| v.extract::<f64>().ok())
                .unwrap_or(0.0);
            let y_split: f64 = cfg
                .get_item("y_split")?
                .and_then(|v| v.extract::<f64>().ok())
                .unwrap_or(0.0);
            ws.set_split(x_split, y_split, None);
        }

        Ok(())
    }

    pub fn save(&mut self, path: &str) -> PyResult<()> {
        if self.saved {
            return Err(PyValueError::new_err(
                "Workbook already saved (NativeWorkbook is consumed-on-save)",
            ));
        }
        let bytes = wolfxl_writer::emit_xlsx(&mut self.inner);
        fs::write(path, bytes)
            .map_err(|e| PyIOError::new_err(format!("failed to write {path}: {e}")))?;
        self.saved = true;
        Ok(())
    }

    // =========================================================================
    // 4B stubs — raise NotImplementedError; existence keeps Python imports working.
    // =========================================================================

    pub fn add_hyperlink(
        &mut self,
        _sheet: &str,
        _link_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        Err(PyNotImplementedError::new_err(
            "NativeWorkbook.add_hyperlink — Wave 4B (rich features). \
             Set WOLFXL_WRITER=oracle to use the rust_xlsxwriter backend for now.",
        ))
    }

    pub fn add_comment(
        &mut self,
        _sheet: &str,
        _comment_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        Err(PyNotImplementedError::new_err(
            "NativeWorkbook.add_comment — Wave 4B (rich features). \
             Set WOLFXL_WRITER=oracle to use the rust_xlsxwriter backend for now.",
        ))
    }

    pub fn set_print_area(&mut self, _sheet: &str, _range_str: &str) -> PyResult<()> {
        Err(PyNotImplementedError::new_err(
            "NativeWorkbook.set_print_area — Wave 4B (rich features). \
             Set WOLFXL_WRITER=oracle to use the rust_xlsxwriter backend for now.",
        ))
    }

    pub fn add_conditional_format(
        &mut self,
        _sheet: &str,
        _rule_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        Err(PyNotImplementedError::new_err(
            "NativeWorkbook.add_conditional_format — Wave 4B (rich features). \
             Set WOLFXL_WRITER=oracle to use the rust_xlsxwriter backend for now.",
        ))
    }

    pub fn add_data_validation(
        &mut self,
        _sheet: &str,
        _validation_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        Err(PyNotImplementedError::new_err(
            "NativeWorkbook.add_data_validation — Wave 4B (rich features). \
             Set WOLFXL_WRITER=oracle to use the rust_xlsxwriter backend for now.",
        ))
    }

    pub fn add_named_range(
        &mut self,
        _sheet: &str,
        _named_range: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        Err(PyNotImplementedError::new_err(
            "NativeWorkbook.add_named_range — Wave 4B (rich features). \
             Set WOLFXL_WRITER=oracle to use the rust_xlsxwriter backend for now.",
        ))
    }

    pub fn add_table(&mut self, _sheet: &str, _table: &Bound<'_, PyAny>) -> PyResult<()> {
        Err(PyNotImplementedError::new_err(
            "NativeWorkbook.add_table — Wave 4B (rich features). \
             Set WOLFXL_WRITER=oracle to use the rust_xlsxwriter backend for now.",
        ))
    }

    pub fn set_properties(&mut self, _props: &Bound<'_, PyAny>) -> PyResult<()> {
        Err(PyNotImplementedError::new_err(
            "NativeWorkbook.set_properties — Wave 4B (rich features). \
             Set WOLFXL_WRITER=oracle to use the rust_xlsxwriter backend for now.",
        ))
    }
}

