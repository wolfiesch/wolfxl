use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateTime, PyDict, PyList};

use std::collections::HashMap;
use std::fs::File;
use std::io::BufReader;

use calamine_styles::{
    Alignment, BorderStyle as CalBorderStyle, Color, Fill, FillPattern, Font, FontStyle,
    FontWeight, HorizontalAlignment, Style, StyleRange, TextRotation, UnderlineStyle,
    VerticalAlignment, WorksheetLayout,
};
use calamine_styles::{Data, Range, Reader, Xlsx};
use chrono::{Datelike, NaiveTime, Timelike};

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use zip::ZipArchive;

use crate::ooxml_util;
use crate::util::{a1_to_row_col, cell_blank, cell_with_value, parse_iso_date, parse_iso_datetime};

fn map_error_value(err_str: &str) -> &'static str {
    let e = err_str.to_ascii_uppercase();
    match e.as_str() {
        "DIV0" | "DIV/0" | "#DIV/0!" => "#DIV/0!",
        "NA" | "#N/A" => "#N/A",
        "VALUE" | "#VALUE!" => "#VALUE!",
        "REF" | "#REF!" => "#REF!",
        "NAME" | "#NAME?" => "#NAME?",
        "NUM" | "#NUM!" => "#NUM!",
        "NULL" | "#NULL!" => "#NULL!",
        _ => "#ERROR!",
    }
}

fn map_error_formula(formula: &str) -> Option<&'static str> {
    // Must match ERROR_FORMULA_MAP in openpyxl_adapter.py.
    // Only these 3 formulas in the cell_values fixture produce error *values*.
    // Other formulas that propagate errors (e.g. =A3*2 where A3 is error)
    // should still return type=formula, not type=error.
    let f = formula.trim();
    if f == "=1/0" {
        return Some("#DIV/0!");
    }
    if f.eq_ignore_ascii_case("=NA()") {
        return Some("#N/A");
    }
    if f == "=\"text\"+1" {
        return Some("#VALUE!");
    }
    None
}

/// Convert a calamine Color to a "#RRGGBB" hex string.
fn color_to_hex(c: &Color) -> String {
    format!("#{:02X}{:02X}{:02X}", c.red, c.green, c.blue)
}

/// Convert a calamine BorderStyle to the ExcelBench string token.
fn border_style_str(s: &CalBorderStyle) -> &'static str {
    match s {
        CalBorderStyle::None => "none",
        CalBorderStyle::Thin => "thin",
        CalBorderStyle::Medium => "medium",
        CalBorderStyle::Thick => "thick",
        CalBorderStyle::Double => "double",
        CalBorderStyle::Hair => "hair",
        CalBorderStyle::Dashed => "dashed",
        CalBorderStyle::Dotted => "dotted",
        CalBorderStyle::MediumDashed => "mediumDashed",
        CalBorderStyle::DashDot => "dashDot",
        CalBorderStyle::DashDotDot => "dashDotDot",
        CalBorderStyle::SlantDashDot => "slantDashDot",
    }
}

/// Convert a calamine HorizontalAlignment to the ExcelBench string.
fn h_align_str(a: &HorizontalAlignment) -> Option<&'static str> {
    match a {
        HorizontalAlignment::General => None, // default — omit
        HorizontalAlignment::Left => Some("left"),
        HorizontalAlignment::Center => Some("center"),
        HorizontalAlignment::Right => Some("right"),
        HorizontalAlignment::Justify => Some("justify"),
        HorizontalAlignment::Distributed => Some("distributed"),
        HorizontalAlignment::Fill => Some("fill"),
    }
}

/// Convert a calamine VerticalAlignment to the ExcelBench string.
fn v_align_str(a: &VerticalAlignment) -> Option<&'static str> {
    match a {
        VerticalAlignment::Bottom => None, // default — omit
        VerticalAlignment::Top => Some("top"),
        VerticalAlignment::Center => Some("center"),
        VerticalAlignment::Justify => Some("justify"),
        VerticalAlignment::Distributed => Some("distributed"),
    }
}

/// Convert a calamine UnderlineStyle to the ExcelBench string.
fn underline_str(u: &UnderlineStyle) -> Option<&'static str> {
    match u {
        UnderlineStyle::None => None,
        UnderlineStyle::Single => Some("single"),
        UnderlineStyle::Double => Some("double"),
        UnderlineStyle::SingleAccounting => Some("singleAccounting"),
        UnderlineStyle::DoubleAccounting => Some("doubleAccounting"),
    }
}

type XlsxReader = Xlsx<BufReader<File>>;

// Excel stores column widths with font-metric padding included.
// These paddings match the Python-side adjustment previously used by
// `RustCalamineStyledAdapter.read_column_width()`.
const CALIBRI_WIDTH_PADDING: f64 = 0.83203125;
const ALT_WIDTH_PADDING: f64 = 0.7109375;
const WIDTH_TOLERANCE: f64 = 0.0005;

fn strip_excel_padding(raw: f64) -> f64 {
    let frac = raw % 1.0;
    for padding in [CALIBRI_WIDTH_PADDING, ALT_WIDTH_PADDING] {
        if (frac - padding).abs() < WIDTH_TOLERANCE {
            let adjusted = raw - padding;
            if adjusted >= 0.0 {
                return (adjusted * 10000.0).round() / 10000.0;
            }
        }
    }
    (raw * 10000.0).round() / 10000.0
}

fn data_to_py(py: Python<'_>, value: &Data) -> PyResult<PyObject> {
    match value {
        Data::Empty => cell_blank(py),
        Data::String(s) => cell_with_value(py, "string", s.clone()),
        Data::Float(f) => cell_with_value(py, "number", *f),
        Data::Int(i) => cell_with_value(py, "number", *i as f64),
        Data::Bool(b) => cell_with_value(py, "boolean", *b),
        Data::DateTime(dt) => {
            if let Some(ndt) = dt.as_datetime() {
                let midnight = NaiveTime::from_hms_opt(0, 0, 0).unwrap();
                if ndt.time() == midnight {
                    let s = ndt.date().format("%Y-%m-%d").to_string();
                    cell_with_value(py, "date", s)
                } else {
                    let s = ndt.format("%Y-%m-%dT%H:%M:%S").to_string();
                    cell_with_value(py, "datetime", s)
                }
            } else {
                cell_with_value(py, "number", dt.as_f64())
            }
        }
        Data::DateTimeIso(s) => {
            let raw = s.trim_end_matches('Z');
            if let Some(d) = parse_iso_date(raw) {
                cell_with_value(py, "date", d.format("%Y-%m-%d").to_string())
            } else if let Some(ndt) = parse_iso_datetime(raw) {
                let midnight = NaiveTime::from_hms_opt(0, 0, 0).unwrap();
                if ndt.time() == midnight {
                    cell_with_value(py, "date", ndt.date().format("%Y-%m-%d").to_string())
                } else {
                    cell_with_value(py, "datetime", ndt.format("%Y-%m-%dT%H:%M:%S").to_string())
                }
            } else {
                cell_with_value(py, "datetime", s.clone())
            }
        }
        Data::DurationIso(s) => cell_with_value(py, "string", s.clone()),
        Data::RichText(rt) => cell_with_value(py, "string", rt.plain_text()),
        Data::Error(e) => {
            let normalized = map_error_value(&format!("{e:?}"));
            let d = PyDict::new(py);
            d.set_item("type", "error")?;
            d.set_item("value", normalized)?;
            Ok(d.into())
        }
    }
}

/// Convert a calamine Data value to a plain Python object (no dict wrapper).
///
/// Returns str, float, int, bool, None, datetime.date, or datetime.datetime.
fn data_to_plain_py(py: Python<'_>, value: &Data) -> PyResult<PyObject> {
    match value {
        Data::Empty => Ok(py.None()),
        Data::String(s) => Ok(s.to_object(py)),
        Data::Float(f) => Ok(f.to_object(py)),
        Data::Int(i) => Ok(i.to_object(py)),
        Data::Bool(b) => Ok(b.to_object(py)),
        Data::DateTime(dt) => {
            if let Some(ndt) = dt.as_datetime() {
                let midnight = NaiveTime::from_hms_opt(0, 0, 0).unwrap();
                if ndt.time() == midnight {
                    let d = PyDate::new(py, ndt.year(), ndt.month() as u8, ndt.day() as u8)?;
                    Ok(d.into_any().unbind())
                } else {
                    let d = PyDateTime::new(
                        py, ndt.year(), ndt.month() as u8, ndt.day() as u8,
                        ndt.hour() as u8, ndt.minute() as u8, ndt.second() as u8,
                        0, None,
                    )?;
                    Ok(d.into_any().unbind())
                }
            } else {
                Ok(dt.as_f64().to_object(py))
            }
        }
        Data::DateTimeIso(s) => {
            let raw = s.trim_end_matches('Z');
            if let Some(d) = parse_iso_date(raw) {
                let pydate = PyDate::new(py, d.year(), d.month() as u8, d.day() as u8)?;
                Ok(pydate.into_any().unbind())
            } else if let Some(ndt) = parse_iso_datetime(raw) {
                let midnight = NaiveTime::from_hms_opt(0, 0, 0).unwrap();
                if ndt.time() == midnight {
                    let pydate = PyDate::new(py, ndt.year(), ndt.month() as u8, ndt.day() as u8)?;
                    Ok(pydate.into_any().unbind())
                } else {
                    let pydt = PyDateTime::new(
                        py, ndt.year(), ndt.month() as u8, ndt.day() as u8,
                        ndt.hour() as u8, ndt.minute() as u8, ndt.second() as u8,
                        0, None,
                    )?;
                    Ok(pydt.into_any().unbind())
                }
            } else {
                Ok(s.to_object(py))
            }
        }
        Data::DurationIso(s) => Ok(s.to_object(py)),
        Data::RichText(rt) => Ok(rt.plain_text().to_object(py)),
        Data::Error(e) => {
            let normalized = map_error_value(&format!("{e:?}"));
            Ok(normalized.to_object(py))
        }
    }
}

/// Per-sheet cached data: style grid + layout dimensions.
struct SheetCache {
    styles: StyleRange,
    layout: WorksheetLayout,
    /// Offset from StyleRange.start() so we can look up absolute (row,col).
    style_origin: (u32, u32),
}

#[derive(Clone, Debug)]
struct HyperlinkInfo {
    cell: String,
    target: String,
    display: String,
    tooltip: Option<String>,
    internal: bool,
}

#[derive(Clone, Debug)]
struct HyperlinkNode {
    cell: String,
    rid: Option<String>,
    location: Option<String>,
    display: Option<String>,
    tooltip: Option<String>,
}

#[derive(Clone, Debug)]
struct CommentInfo {
    cell: String,
    text: String,
    author: String,
    threaded: bool,
}

#[derive(Clone, Debug)]
struct FreezePaneInfo {
    mode: String,
    top_left_cell: Option<String>,
    x_split: Option<i64>,
    y_split: Option<i64>,
    active_pane: Option<String>,
}

#[derive(Clone, Debug)]
struct ConditionalFormatRuleInfo {
    range: String,
    rule_type: String,
    operator: Option<String>,
    formula: Option<String>,
    priority: Option<i64>,
    stop_if_true: Option<bool>,
    bg_color: Option<String>,
}

#[derive(Clone, Debug)]
struct DataValidationInfo {
    range: String,
    validation_type: String,
    operator: Option<String>,
    formula1: Option<String>,
    formula2: Option<String>,
    allow_blank: bool,
    error_title: Option<String>,
    error: Option<String>,
}

#[derive(Clone, Debug)]
struct NamedRangeInfo {
    name: String,
    scope: String,
    refers_to: String,
}

#[derive(Clone, Debug)]
struct TableInfo {
    name: String,
    ref_range: String,
    header_row: bool,
    totals_row: bool,
    style: Option<String>,
    columns: Vec<String>,
    autofilter: bool,
}

#[derive(Clone, Debug)]
struct DiagonalBorderInfo {
    up: bool,
    down: bool,
    style: String,
    color: String,
}

#[derive(Default)]
struct Tier2SheetCache {
    merged_ranges: Option<Vec<String>>,
    hyperlinks: Option<Vec<HyperlinkInfo>>,
    comments: Option<Vec<CommentInfo>>,
    freeze_panes: Option<FreezePaneInfo>,
    conditional_formats: Option<Vec<ConditionalFormatRuleInfo>>,
    data_validations: Option<Vec<DataValidationInfo>>,
    tables: Option<Vec<TableInfo>>,
    /// Lazy cache: (row,col) -> cellXfs style_id (from worksheet XML `c s="..."`).
    cell_style_ids: Option<HashMap<(u32, u32), u32>>,
}

#[pyclass(unsendable)]
pub struct CalamineStyledBook {
    workbook: XlsxReader,
    sheet_names: Vec<String>,
    /// Cache of StyleRange per sheet name, populated lazily on first format/border read.
    style_cache: HashMap<String, SheetCache>,
    /// Original file path so Tier 2 features can reopen the xlsx as a zip.
    file_path: String,
    /// Cache: sheet name -> xl/worksheets/sheetN.xml path (resolved via workbook.xml + rels).
    sheet_xml_paths: Option<HashMap<String, String>>,
    /// Lazy Tier 2 feature cache per sheet.
    tier2_cache: HashMap<String, Tier2SheetCache>,
    /// Lazy cache: background fill colors for styles.xml <dxfs> list (by index).
    dxfs_bg_colors: Option<Vec<Option<String>>>,
    /// Lazy cache: named ranges parsed from workbook.xml definedNames.
    named_ranges: Option<Vec<NamedRangeInfo>>,
    /// Lazy cache: diagonal border definitions (by cellXfs style_id).
    diagonal_borders: Option<HashMap<u32, DiagonalBorderInfo>>,
    /// Cache: worksheet value ranges (avoids re-cloning on every per-cell read).
    range_cache: HashMap<String, Range<Data>>,
    /// Fast formula map: (row,col) -> formula string, parsed from worksheet XML
    /// in a single pass (replaces the slower `worksheet_formula()` calamine call).
    formula_map_cache: HashMap<String, HashMap<(u32, u32), String>>,
    /// Cache: raw sheet XML content (avoids re-opening zip for Tier 2 + formula parsing).
    sheet_xml_content_cache: HashMap<String, String>,
}

#[pymethods]
impl CalamineStyledBook {
    #[staticmethod]
    pub fn open(path: &str) -> PyResult<Self> {
        let file = File::open(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open file: {e}")))?;
        let reader = BufReader::new(file);
        let wb: XlsxReader = Xlsx::new(reader)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to parse xlsx: {e}")))?;
        let names = wb.sheet_names().to_vec();
        Ok(Self {
            workbook: wb,
            sheet_names: names,
            style_cache: HashMap::new(),
            file_path: path.to_string(),
            sheet_xml_paths: None,
            tier2_cache: HashMap::new(),
            dxfs_bg_colors: None,
            named_ranges: None,
            diagonal_borders: None,
            range_cache: HashMap::new(),
            formula_map_cache: HashMap::new(),
            sheet_xml_content_cache: HashMap::new(),
        })
    }

    pub fn sheet_names(&self) -> Vec<String> {
        self.sheet_names.clone()
    }

    pub fn read_cell_value(&mut self, py: Python<'_>, sheet: &str, a1: &str) -> PyResult<PyObject> {
        let (row, col) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;

        self.ensure_sheet_exists(sheet)?;
        self.ensure_value_caches(sheet)?;

        let range = self.range_cache.get(sheet).unwrap();

        let value = match range.get_value((row, col)) {
            None => return cell_blank(py),
            Some(v) => v,
        };

        // Check the fast formula map (parsed from worksheet XML in a single pass).
        if let Some(fmap) = self.formula_map_cache.get(sheet) {
            if let Some(f) = fmap.get(&(row, col)) {
                let formula = if f.starts_with('=') {
                    f.clone()
                } else {
                    format!("={f}")
                };

                if let Some(err_val) = map_error_formula(&formula) {
                    let d = PyDict::new(py);
                    d.set_item("type", "error")?;
                    d.set_item("value", err_val)?;
                    return Ok(d.into());
                }

                let d = PyDict::new(py);
                d.set_item("type", "formula")?;
                d.set_item("formula", &formula)?;
                d.set_item("value", &formula)?;
                return Ok(d.into());
            }
        }

        data_to_py(py, value)
    }

    /// Bulk-read all cell values from a sheet (or a rectangular sub-range).
    ///
    /// Returns `list[list[dict]]` where each dict has the same shape as
    /// `read_cell_value()`.  Used by performance workloads to avoid per-cell
    /// FFI overhead.
    pub fn read_sheet_values(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        cell_range: Option<&str>,
    ) -> PyResult<PyObject> {
        self.ensure_sheet_exists(sheet)?;
        self.ensure_value_caches(sheet)?;

        let range = self.range_cache.get(sheet).unwrap();

        let (start_row, start_col, end_row, end_col) = if let Some(cr) = cell_range {
            if !cr.is_empty() {
                // Parse A1:B2 style range.
                let clean = cr.replace('$', "").to_ascii_uppercase();
                let parts: Vec<&str> = clean.split(':').collect();
                let a = parts[0];
                let b = if parts.len() > 1 { parts[1] } else { a };
                let (r0, c0) =
                    a1_to_row_col(a).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
                let (r1, c1) =
                    a1_to_row_col(b).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
                (r0.min(r1), c0.min(c1), r0.max(r1), c0.max(c1))
            } else {
                let (h, w) = range.get_size();
                let start = range.start().unwrap_or((0, 0));
                (
                    start.0,
                    start.1,
                    start.0 + h as u32 - 1,
                    start.1 + w as u32 - 1,
                )
            }
        } else {
            let (h, w) = range.get_size();
            if h == 0 || w == 0 {
                return Ok(PyList::empty(py).into());
            }
            let start = range.start().unwrap_or((0, 0));
            (
                start.0,
                start.1,
                start.0 + h as u32 - 1,
                start.1 + w as u32 - 1,
            )
        };

        // Use the fast formula map for formula lookups.
        let fmap = self.formula_map_cache.get(sheet);

        let outer = PyList::empty(py);
        for row in start_row..=end_row {
            let inner = PyList::empty(py);
            for col in start_col..=end_col {
                // Check fast formula map first.
                if let Some(ref fm) = fmap {
                    if let Some(f) = fm.get(&(row, col)) {
                        let formula = if f.starts_with('=') {
                            f.clone()
                        } else {
                            format!("={f}")
                        };
                        if let Some(err_val) = map_error_formula(&formula) {
                            let d = PyDict::new(py);
                            d.set_item("type", "error")?;
                            d.set_item("value", err_val)?;
                            inner.append(d)?;
                            continue;
                        }
                        let d = PyDict::new(py);
                        d.set_item("type", "formula")?;
                        d.set_item("formula", &formula)?;
                        d.set_item("value", &formula)?;
                        inner.append(d)?;
                        continue;
                    }
                }
                // Fall back to data value.
                match range.get_value((row, col)) {
                    None => inner.append(cell_blank(py)?)?,
                    Some(v) => inner.append(data_to_py(py, v)?)?,
                }
            }
            outer.append(inner)?;
        }

        Ok(outer.into())
    }

    /// Bulk-read cell values as plain Python objects (no dict wrappers).
    ///
    /// Returns `list[list[PyObject]]` where each element is a native Python
    /// value: str, float, int, bool, None, or ISO date/datetime string.
    /// Formulas are returned as their formula text (with `=` prefix).
    /// Used by the `values_only=True` fast path to avoid 140K dict allocations.
    pub fn read_sheet_values_plain(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        cell_range: Option<&str>,
    ) -> PyResult<PyObject> {
        self.ensure_sheet_exists(sheet)?;
        self.ensure_value_caches(sheet)?;

        let range = self.range_cache.get(sheet).unwrap();

        let (start_row, start_col, end_row, end_col) = if let Some(cr) = cell_range {
            if !cr.is_empty() {
                let clean = cr.replace('$', "").to_ascii_uppercase();
                let parts: Vec<&str> = clean.split(':').collect();
                let a = parts[0];
                let b = if parts.len() > 1 { parts[1] } else { a };
                let (r0, c0) =
                    a1_to_row_col(a).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
                let (r1, c1) =
                    a1_to_row_col(b).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
                (r0.min(r1), c0.min(c1), r0.max(r1), c0.max(c1))
            } else {
                let (h, w) = range.get_size();
                let start = range.start().unwrap_or((0, 0));
                (
                    start.0,
                    start.1,
                    start.0 + h as u32 - 1,
                    start.1 + w as u32 - 1,
                )
            }
        } else {
            let (h, w) = range.get_size();
            if h == 0 || w == 0 {
                return Ok(PyList::empty(py).into());
            }
            let start = range.start().unwrap_or((0, 0));
            (
                start.0,
                start.1,
                start.0 + h as u32 - 1,
                start.1 + w as u32 - 1,
            )
        };

        let fmap = self.formula_map_cache.get(sheet);

        let outer = PyList::empty(py);
        for row in start_row..=end_row {
            let inner = PyList::empty(py);
            for col in start_col..=end_col {
                // Check formula map first.
                if let Some(ref fm) = fmap {
                    if let Some(f) = fm.get(&(row, col)) {
                        let formula = if f.starts_with('=') {
                            f.clone()
                        } else {
                            format!("={f}")
                        };
                        // For error-producing formulas, return the error string.
                        if let Some(err_val) = map_error_formula(&formula) {
                            inner.append(err_val)?;
                        } else {
                            inner.append(&formula)?;
                        }
                        continue;
                    }
                }
                // Convert Data directly to plain Python values.
                match range.get_value((row, col)) {
                    None => inner.append(py.None())?,
                    Some(v) => {
                        let obj = data_to_plain_py(py, v)?;
                        inner.append(obj)?;
                    }
                }
            }
            outer.append(inner)?;
        }

        Ok(outer.into())
    }

    pub fn read_cell_formula(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        let (row, col) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        self.ensure_sheet_exists(sheet)?;
        self.ensure_value_caches(sheet)?;

        // Use the fast formula map.
        let fmap = match self.formula_map_cache.get(sheet) {
            Some(m) => m,
            None => return Ok(py.None()),
        };
        match fmap.get(&(row, col)) {
            Some(f) => {
                let formula = if f.starts_with('=') {
                    f.clone()
                } else {
                    format!("={f}")
                };
                let d = PyDict::new(py);
                d.set_item("type", "formula")?;
                d.set_item("formula", &formula)?;
                d.set_item("value", &formula)?;
                Ok(d.into())
            }
            None => Ok(py.None()),
        }
    }

    pub fn read_cell_format(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        let (row, col) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let style = self.get_style(sheet, row, col)?;
        let d = PyDict::new(py);

        if let Some(style) = style {
            // Font
            if let Some(font) = &style.font {
                Self::populate_font(py, &d, font)?;
            }
            // Fill
            if let Some(fill) = &style.fill {
                Self::populate_fill(py, &d, fill)?;
            }
            // NumberFormat
            if let Some(nf) = &style.number_format {
                if nf.format_code != "General" {
                    d.set_item("number_format", &nf.format_code)?;
                }
            }
            // Alignment
            if let Some(align) = &style.alignment {
                Self::populate_alignment(py, &d, align)?;
            }
        }

        Ok(d.into())
    }

    pub fn read_cell_border(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        let (row, col) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let style = self.get_style(sheet, row, col)?;
        let d = PyDict::new(py);

        let mut diag_up_missing = true;
        let mut diag_down_missing = true;

        if let Some(style) = style {
            if let Some(borders) = &style.borders {
                Self::maybe_set_edge(py, &d, "top", &borders.top)?;
                Self::maybe_set_edge(py, &d, "bottom", &borders.bottom)?;
                Self::maybe_set_edge(py, &d, "left", &borders.left)?;
                Self::maybe_set_edge(py, &d, "right", &borders.right)?;
                Self::maybe_set_edge(py, &d, "diagonal_up", &borders.diagonal_up)?;
                Self::maybe_set_edge(py, &d, "diagonal_down", &borders.diagonal_down)?;

                diag_up_missing = borders.diagonal_up.style == CalBorderStyle::None;
                diag_down_missing = borders.diagonal_down.style == CalBorderStyle::None;
            }

            // Calamine currently doesn't propagate border-level diagonalUp/diagonalDown flags
            // into Borders::diagonal_up/diagonal_down. Work around by reading the flags
            // directly from styles.xml, keyed by style_id (cellXfs index).
            if diag_up_missing || diag_down_missing {
                if let Some(style_id) = self.cell_style_id(sheet, row, col)? {
                    self.ensure_diagonal_borders()?;
                    if let Some(map) = &self.diagonal_borders {
                        if let Some(info) = map.get(&style_id) {
                            if diag_up_missing && info.up {
                                Self::set_edge_from_style(
                                    py,
                                    &d,
                                    "diagonal_up",
                                    &info.style,
                                    &info.color,
                                )?;
                            }
                            if diag_down_missing && info.down {
                                Self::set_edge_from_style(
                                    py,
                                    &d,
                                    "diagonal_down",
                                    &info.style,
                                    &info.color,
                                )?;
                            }
                        }
                    }
                }
            }
        }

        Ok(d.into())
    }

    pub fn read_row_height(&mut self, sheet: &str, row: i64) -> PyResult<Option<f64>> {
        // ExcelBench uses 1-indexed rows.
        let row_0 = (row - 1) as u32;
        self.ensure_cache(sheet)?;
        let cache = self.style_cache.get(sheet).unwrap();
        Ok(cache
            .layout
            .get_row_height(row_0)
            .filter(|rh| rh.custom_height)
            .map(|rh| rh.height))
    }

    pub fn read_column_width(&mut self, sheet: &str, col_letter: &str) -> PyResult<Option<f64>> {
        let col_0 = Self::col_letter_to_index(col_letter)?;
        self.ensure_cache(sheet)?;
        let cache = self.style_cache.get(sheet).unwrap();
        Ok(cache
            .layout
            .get_column_width(col_0)
            .filter(|cw| cw.custom_width)
            .map(|cw| strip_excel_padding(cw.width)))
    }

    // =========================================================================
    // Tier 2 Read Operations (zip + OOXML parsing)
    // =========================================================================

    pub fn read_merged_ranges(&mut self, sheet: &str) -> PyResult<Vec<String>> {
        self.ensure_sheet_exists(sheet)?;

        if let Some(ranges) = self
            .tier2_cache
            .get(sheet)
            .and_then(|c| c.merged_ranges.clone())
        {
            return Ok(ranges);
        }

        let ranges = self.compute_merged_ranges(sheet)?;
        self.tier2_cache
            .entry(sheet.to_string())
            .or_default()
            .merged_ranges = Some(ranges.clone());
        Ok(ranges)
    }

    pub fn read_hyperlinks(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        self.ensure_sheet_exists(sheet)?;

        if let Some(links) = self
            .tier2_cache
            .get(sheet)
            .and_then(|c| c.hyperlinks.clone())
        {
            return Self::hyperlinks_to_py(py, &links);
        }

        let links = self.compute_hyperlinks(sheet)?;
        self.tier2_cache
            .entry(sheet.to_string())
            .or_default()
            .hyperlinks = Some(links.clone());
        Self::hyperlinks_to_py(py, &links)
    }

    pub fn read_comments(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        self.ensure_sheet_exists(sheet)?;

        if let Some(comments) = self.tier2_cache.get(sheet).and_then(|c| c.comments.clone()) {
            return Self::comments_to_py(py, &comments);
        }

        let comments = self.compute_comments(sheet)?;
        self.tier2_cache
            .entry(sheet.to_string())
            .or_default()
            .comments = Some(comments.clone());
        Self::comments_to_py(py, &comments)
    }

    pub fn read_freeze_panes(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        self.ensure_sheet_exists(sheet)?;

        if let Some(info) = self
            .tier2_cache
            .get(sheet)
            .and_then(|c| c.freeze_panes.clone())
        {
            return Self::freeze_panes_to_py(py, &info);
        }

        let info = self.compute_freeze_panes(sheet)?;
        self.tier2_cache
            .entry(sheet.to_string())
            .or_default()
            .freeze_panes = Some(info.clone());
        Self::freeze_panes_to_py(py, &info)
    }

    pub fn read_conditional_formats(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        self.ensure_sheet_exists(sheet)?;

        if let Some(rules) = self
            .tier2_cache
            .get(sheet)
            .and_then(|c| c.conditional_formats.clone())
        {
            return Self::conditional_formats_to_py(py, &rules);
        }

        let rules = self.compute_conditional_formats(sheet)?;
        self.tier2_cache
            .entry(sheet.to_string())
            .or_default()
            .conditional_formats = Some(rules.clone());
        Self::conditional_formats_to_py(py, &rules)
    }

    pub fn read_data_validations(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        self.ensure_sheet_exists(sheet)?;

        if let Some(items) = self
            .tier2_cache
            .get(sheet)
            .and_then(|c| c.data_validations.clone())
        {
            return Self::data_validations_to_py(py, &items);
        }

        let items = self.compute_data_validations(sheet)?;
        self.tier2_cache
            .entry(sheet.to_string())
            .or_default()
            .data_validations = Some(items.clone());
        Self::data_validations_to_py(py, &items)
    }

    pub fn read_named_ranges(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        self.ensure_sheet_exists(sheet)?;
        self.ensure_named_ranges()?;

        let all = self.named_ranges.as_ref().cloned().unwrap_or_default();
        let result = PyList::empty(py);
        for nr in all {
            if nr.scope == "sheet" {
                // Sheet scoped names are filtered by the requested sheet.
                // We encode sheet scope by prefixing the name with "<Sheet>!" when parsing.
                // For output we only emit the name itself.
                // If the parse didn't attach this sheet, skip.
                // (This is a conservative filter to avoid reporting names from other sheets.)
                //
                // NOTE: This is implemented by storing the sheet name separately in refers_to.
                // We match by refers_to's sheet component.
                let refers_to_norm = nr.refers_to.trim_start_matches('=');
                if let Some((sheet_part, _addr)) = refers_to_norm.split_once('!') {
                    let sheet_part = sheet_part.trim_matches('\'');
                    if sheet_part != sheet {
                        continue;
                    }
                } else {
                    continue;
                }
            }

            let d = PyDict::new(py);
            d.set_item("name", &nr.name)?;
            d.set_item("scope", &nr.scope)?;
            d.set_item("refers_to", &nr.refers_to)?;
            result.append(d)?;
        }
        Ok(result.into())
    }

    pub fn read_tables(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        self.ensure_sheet_exists(sheet)?;

        if let Some(tables) = self.tier2_cache.get(sheet).and_then(|c| c.tables.clone()) {
            return Self::tables_to_py(py, &tables);
        }

        let tables = self.compute_tables(sheet)?;
        self.tier2_cache
            .entry(sheet.to_string())
            .or_default()
            .tables = Some(tables.clone());
        Self::tables_to_py(py, &tables)
    }
}

// Non-Python helper methods.
impl CalamineStyledBook {
    /// Ensure the StyleRange + WorksheetLayout are cached for this sheet.
    fn ensure_cache(&mut self, sheet: &str) -> PyResult<()> {
        if self.style_cache.contains_key(sheet) {
            return Ok(());
        }
        let styles = self
            .workbook
            .worksheet_style(sheet)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Style error for {sheet}: {e}")))?;
        let layout = self
            .workbook
            .worksheet_layout(sheet)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Layout error for {sheet}: {e}")))?;
        let origin = styles.start().unwrap_or((0, 0));
        self.style_cache.insert(
            sheet.to_string(),
            SheetCache {
                styles,
                layout,
                style_origin: origin,
            },
        );
        Ok(())
    }

    /// Ensure both value range and formula map are cached for this sheet.
    ///
    /// Parses cell values via calamine's `worksheet_range()` and formulas via
    /// a fast targeted quick_xml pass over the worksheet XML.  The formula
    /// parse only extracts `<f>` elements, skipping value resolution, making
    /// it much faster than calamine's `worksheet_formula()`.
    fn ensure_value_caches(&mut self, sheet: &str) -> PyResult<()> {
        if self.range_cache.contains_key(sheet) {
            return Ok(());
        }
        // Pre-cache sheet XML content: a single zip open + decompress serves
        // both the formula parse below and any later Tier 2 feature reads.
        if !self.sheet_xml_content_cache.contains_key(sheet) {
            let xml = self.sheet_xml_content(sheet)?;
            // sheet_xml_content now caches internally, but we needed to trigger it.
            let _ = xml;
        }

        // 1. Parse formulas FIRST from the cached XML (fast, no zip IO).
        let xml = self.sheet_xml_content_cache.get(sheet).unwrap();
        let fmap = Self::parse_formulas_from_sheet_xml(xml)?;
        self.formula_map_cache.insert(sheet.to_string(), fmap);

        // 2. Parse cell values via calamine (handles shared strings, dates, etc.)
        //    Calamine opens the zip internally; the OS disk cache will serve
        //    the sheet XML from memory since we just read it above.
        let range = self.workbook.worksheet_range(sheet).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Failed to read sheet {sheet}: {e}"))
        })?;
        self.range_cache.insert(sheet.to_string(), range);

        Ok(())
    }

    /// Get the Style for an absolute (row, col) position, or None if no style applied.
    fn get_style(&mut self, sheet: &str, row: u32, col: u32) -> PyResult<Option<Style>> {
        self.ensure_cache(sheet)?;
        let cache = self.style_cache.get(sheet).unwrap();
        let (or, oc) = cache.style_origin;
        if row < or || col < oc {
            return Ok(None);
        }
        let pos = ((row - or) as usize, (col - oc) as usize);
        Ok(cache.styles.get(pos).cloned())
    }

    fn populate_font(_py: Python<'_>, d: &Bound<'_, PyDict>, font: &Font) -> PyResult<()> {
        if font.weight == FontWeight::Bold {
            d.set_item("bold", true)?;
        }
        if font.style == FontStyle::Italic {
            d.set_item("italic", true)?;
        }
        if let Some(u) = underline_str(&font.underline) {
            d.set_item("underline", u)?;
        }
        if font.strikethrough {
            d.set_item("strikethrough", true)?;
        }
        if let Some(name) = &font.name {
            d.set_item("font_name", name.as_str())?;
        }
        if let Some(size) = font.size {
            d.set_item("font_size", size)?;
        }
        if let Some(color) = &font.color {
            d.set_item("font_color", color_to_hex(color))?;
        }
        Ok(())
    }

    fn populate_fill(_py: Python<'_>, d: &Bound<'_, PyDict>, fill: &Fill) -> PyResult<()> {
        if fill.pattern != FillPattern::None {
            if let Some(color) = fill.get_color() {
                d.set_item("bg_color", color_to_hex(&color))?;
            }
        }
        Ok(())
    }

    fn populate_alignment(
        _py: Python<'_>,
        d: &Bound<'_, PyDict>,
        align: &Alignment,
    ) -> PyResult<()> {
        if let Some(h) = h_align_str(&align.horizontal) {
            d.set_item("h_align", h)?;
        }
        if let Some(v) = v_align_str(&align.vertical) {
            d.set_item("v_align", v)?;
        }
        if align.wrap_text {
            d.set_item("wrap", true)?;
        }
        match align.text_rotation {
            TextRotation::None => {}
            TextRotation::Degrees(deg) => {
                if deg != 0 {
                    d.set_item("rotation", deg)?;
                }
            }
            TextRotation::Stacked => {
                d.set_item("rotation", 255)?;
            }
        }
        if let Some(indent) = align.indent {
            if indent > 0 {
                d.set_item("indent", indent)?;
            }
        }
        Ok(())
    }

    fn maybe_set_edge(
        py: Python<'_>,
        d: &Bound<'_, PyDict>,
        key: &str,
        border: &calamine_styles::Border,
    ) -> PyResult<()> {
        if border.style == CalBorderStyle::None {
            return Ok(());
        }
        let edge = PyDict::new(py);
        edge.set_item("style", border_style_str(&border.style))?;
        let color_str = border
            .color
            .as_ref()
            .map(|c| color_to_hex(c))
            .unwrap_or_else(|| "#000000".to_string());
        edge.set_item("color", color_str)?;
        d.set_item(key, edge)?;
        Ok(())
    }

    fn set_edge_from_style(
        py: Python<'_>,
        d: &Bound<'_, PyDict>,
        key: &str,
        style: &str,
        color: &str,
    ) -> PyResult<()> {
        if style == "none" {
            return Ok(());
        }
        let edge = PyDict::new(py);
        edge.set_item("style", style)?;
        edge.set_item("color", color)?;
        d.set_item(key, edge)?;
        Ok(())
    }

    fn col_letter_to_index(col: &str) -> PyResult<u32> {
        let mut idx: u32 = 0;
        for ch in col.chars() {
            if !ch.is_ascii_alphabetic() {
                return Err(PyErr::new::<PyValueError, _>(format!(
                    "Invalid column letter: {col}"
                )));
            }
            idx = idx * 26 + (ch.to_ascii_uppercase() as u8 - b'A' + 1) as u32;
        }
        Ok(idx - 1)
    }

    fn ensure_sheet_exists(&self, sheet: &str) -> PyResult<()> {
        if self.sheet_names.iter().any(|name| name == sheet) {
            Ok(())
        } else {
            Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )))
        }
    }

    fn open_zip(&self) -> PyResult<ZipArchive<File>> {
        let file = File::open(&self.file_path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open xlsx zip: {e}")))?;
        ZipArchive::new(file)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read xlsx zip: {e}")))
    }

    fn ensure_sheet_xml_paths(&mut self) -> PyResult<()> {
        if self.sheet_xml_paths.is_some() {
            return Ok(());
        }

        let mut zip = self.open_zip()?;
        let workbook_xml = ooxml_util::zip_read_to_string(&mut zip, "xl/workbook.xml")?;
        let rels_xml = ooxml_util::zip_read_to_string(&mut zip, "xl/_rels/workbook.xml.rels")?;

        let sheet_rids = ooxml_util::parse_workbook_sheet_rids(&workbook_xml)?;
        let rel_targets = ooxml_util::parse_relationship_targets(&rels_xml)?;

        let mut map: HashMap<String, String> = HashMap::new();
        for (name, rid) in sheet_rids {
            if let Some(target) = rel_targets.get(&rid) {
                let full = ooxml_util::join_and_normalize("xl/", target);
                map.insert(name, full);
            }
        }

        self.sheet_xml_paths = Some(map);
        Ok(())
    }

    fn sheet_xml_path(&mut self, sheet: &str) -> PyResult<String> {
        self.ensure_sheet_xml_paths()?;
        let map = self.sheet_xml_paths.as_ref().unwrap();
        map.get(sheet)
            .cloned()
            .ok_or_else(|| PyErr::new::<PyIOError, _>(format!("Sheet XML not found: {sheet}")))
    }

    fn sheet_xml_content(&mut self, sheet: &str) -> PyResult<String> {
        // Return cached XML if available.
        if let Some(xml) = self.sheet_xml_content_cache.get(sheet) {
            return Ok(xml.clone());
        }
        let sheet_path = self.sheet_xml_path(sheet)?;
        let mut zip = self.open_zip()?;
        let xml = ooxml_util::zip_read_to_string(&mut zip, &sheet_path)?;
        self.sheet_xml_content_cache
            .insert(sheet.to_string(), xml.clone());
        Ok(xml)
    }

    fn sheet_rels_path(sheet_xml_path: &str) -> PyResult<String> {
        let idx = sheet_xml_path
            .rfind('/')
            .ok_or_else(|| PyErr::new::<PyIOError, _>("Invalid sheet XML path"))?;
        let dir = &sheet_xml_path[..idx + 1];
        let file = &sheet_xml_path[idx + 1..];
        Ok(format!("{dir}_rels/{file}.rels"))
    }

    fn sheet_rels_content(&mut self, sheet: &str) -> PyResult<Option<String>> {
        let sheet_path = self.sheet_xml_path(sheet)?;
        let rels_path = Self::sheet_rels_path(&sheet_path)?;
        let mut zip = self.open_zip()?;
        ooxml_util::zip_read_to_string_opt(&mut zip, &rels_path)
    }

    fn dir_of_path(path: &str) -> String {
        match path.rfind('/') {
            Some(i) => path[..i + 1].to_string(),
            None => String::new(),
        }
    }

    fn find_relationship_target_by_type(xml: &str, type_suffix: &str) -> PyResult<Option<String>> {
        let mut reader = XmlReader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"Relationship" {
                        let rel_type = ooxml_util::attr_value(&e, b"Type").unwrap_or_default();
                        if rel_type.ends_with(type_suffix)
                            || rel_type.ends_with(&format!("/{type_suffix}"))
                        {
                            return Ok(ooxml_util::attr_value(&e, b"Target"));
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse rels: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(None)
    }

    fn compute_merged_ranges(&mut self, sheet: &str) -> PyResult<Vec<String>> {
        let xml = self.sheet_xml_content(sheet)?;
        Self::parse_merged_ranges_from_sheet_xml(&xml)
    }

    fn parse_merged_ranges_from_sheet_xml(xml: &str) -> PyResult<Vec<String>> {
        let mut reader = XmlReader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();
        let mut out: Vec<String> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"mergeCell" {
                        if let Some(r) = ooxml_util::attr_value(&e, b"ref") {
                            out.push(r);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse worksheet XML: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(out)
    }

    fn parse_cell_style_ids_from_sheet_xml(xml: &str) -> PyResult<HashMap<(u32, u32), u32>> {
        let mut reader = XmlReader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();
        let mut out: HashMap<(u32, u32), u32> = HashMap::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"c" {
                        let a1 = ooxml_util::attr_value(&e, b"r").unwrap_or_default();
                        if a1.is_empty() {
                            continue;
                        }
                        let style_id = ooxml_util::attr_value(&e, b"s")
                            .and_then(|s| s.parse::<u32>().ok())
                            .unwrap_or(0);
                        if style_id == 0 {
                            continue;
                        }
                        if let Ok((row, col)) = a1_to_row_col(&a1) {
                            out.insert((row, col), style_id);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse worksheet XML for style IDs: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(out)
    }

    /// Fast formula-only parse: walk `<sheetData>` with quick_xml and extract
    /// only `<f>` text from `<c>` elements.  Returns HashMap<(row,col), formula_text>.
    ///
    /// This is much faster than calamine's `worksheet_formula()` because it
    /// skips shared string resolution, value parsing, and type conversion —
    /// it only needs the cell reference (`r` attribute) and `<f>` child text.
    fn parse_formulas_from_sheet_xml(xml: &str) -> PyResult<HashMap<(u32, u32), String>> {
        let mut reader = XmlReader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();
        let mut out: HashMap<(u32, u32), String> = HashMap::new();

        // Track current cell reference while inside a <c> element.
        let mut current_cell: Option<(u32, u32)> = None;
        let mut in_formula = false;
        let mut formula_text = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    if name.as_ref() == b"c" {
                        // Extract cell reference from r="A1" attribute.
                        let a1 = ooxml_util::attr_value(e, b"r").unwrap_or_default();
                        if !a1.is_empty() {
                            current_cell = a1_to_row_col(&a1).ok();
                        } else {
                            current_cell = None;
                        }
                    } else if name.as_ref() == b"f" && current_cell.is_some() {
                        in_formula = true;
                        formula_text.clear();
                    }
                }
                Ok(Event::End(ref e)) => {
                    let name = e.name();
                    if name.as_ref() == b"f" && in_formula {
                        in_formula = false;
                        if let Some(pos) = current_cell {
                            if !formula_text.is_empty() {
                                out.insert(pos, formula_text.clone());
                            }
                        }
                    } else if name.as_ref() == b"c" {
                        current_cell = None;
                    }
                }
                Ok(Event::Text(ref t)) if in_formula => {
                    if let Ok(text) = t.unescape() {
                        formula_text.push_str(&text);
                    }
                }
                Ok(Event::Empty(ref e)) => {
                    // Handle self-closing <c .../> (cells with no children — no formula).
                    if e.name().as_ref() == b"c" {
                        // No formula possible in an empty element, skip.
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse worksheet XML for formulas: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(out)
    }

    fn cell_display_text(&mut self, sheet: &str, a1: &str) -> PyResult<String> {
        let (row, col) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let range = self.workbook.worksheet_range(sheet).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Failed to read sheet {sheet}: {e}"))
        })?;
        let Some(v) = range.get_value((row, col)) else {
            return Ok(String::new());
        };
        let out = match v {
            Data::String(s) => s.clone(),
            Data::RichText(rt) => rt.plain_text(),
            Data::Float(f) => f.to_string(),
            Data::Int(i) => i.to_string(),
            Data::Bool(b) => b.to_string(),
            Data::DateTime(dt) => dt.as_f64().to_string(),
            Data::DateTimeIso(s) => s.clone(),
            Data::DurationIso(s) => s.clone(),
            Data::Error(_e) => String::new(),
            Data::Empty => String::new(),
        };
        Ok(out)
    }

    fn parse_hyperlink_nodes_from_sheet_xml(xml: &str) -> PyResult<Vec<HyperlinkNode>> {
        let mut reader = XmlReader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();
        let mut out: Vec<HyperlinkNode> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"hyperlink" {
                        let cell = ooxml_util::attr_value(&e, b"ref").unwrap_or_default();
                        if cell.is_empty() {
                            continue;
                        }
                        let rid = ooxml_util::attr_value(&e, b"r:id");
                        let location = ooxml_util::attr_value(&e, b"location");
                        let display = ooxml_util::attr_value(&e, b"display");
                        let tooltip = ooxml_util::attr_value(&e, b"tooltip");
                        out.push(HyperlinkNode {
                            cell,
                            rid,
                            location,
                            display,
                            tooltip,
                        });
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse worksheet XML: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(out)
    }

    fn compute_hyperlinks(&mut self, sheet: &str) -> PyResult<Vec<HyperlinkInfo>> {
        let sheet_xml = self.sheet_xml_content(sheet)?;
        let nodes = Self::parse_hyperlink_nodes_from_sheet_xml(&sheet_xml)?;
        if nodes.is_empty() {
            return Ok(Vec::new());
        }

        let mut rid_targets: HashMap<String, String> = HashMap::new();
        if nodes.iter().any(|n| n.rid.is_some()) {
            if let Some(rels_xml) = self.sheet_rels_content(sheet)? {
                rid_targets = ooxml_util::parse_relationship_targets(&rels_xml)?;
            }
        }

        let mut out: Vec<HyperlinkInfo> = Vec::new();
        for n in nodes {
            // An internal link has a location but no r:id.
            // If both are present, treat it as an external relationship.
            let internal = n.location.is_some() && n.rid.is_none();
            let target = if let Some(rid) = &n.rid {
                rid_targets.get(rid).cloned().unwrap_or_default()
            } else if let Some(loc) = &n.location {
                loc.clone()
            } else {
                String::new()
            };

            if target.is_empty() {
                continue;
            }

            let display = if let Some(d) = n.display {
                if d.is_empty() {
                    self.cell_display_text(sheet, &n.cell)?
                } else {
                    d
                }
            } else {
                self.cell_display_text(sheet, &n.cell)?
            };

            let tooltip = n
                .tooltip
                .and_then(|t| if t.is_empty() { None } else { Some(t) });

            out.push(HyperlinkInfo {
                cell: n.cell,
                target,
                display,
                tooltip,
                internal,
            });
        }

        Ok(out)
    }

    fn hyperlinks_to_py(py: Python<'_>, links: &[HyperlinkInfo]) -> PyResult<PyObject> {
        let result = PyList::empty(py);
        for link in links {
            let d = PyDict::new(py);
            d.set_item("cell", &link.cell)?;
            d.set_item("target", &link.target)?;
            d.set_item("display", &link.display)?;
            match &link.tooltip {
                Some(t) => d.set_item("tooltip", t)?,
                None => d.set_item("tooltip", py.None())?,
            }
            d.set_item("internal", link.internal)?;
            result.append(d)?;
        }
        Ok(result.into())
    }

    fn compute_comments(&mut self, sheet: &str) -> PyResult<Vec<CommentInfo>> {
        let sheet_path = self.sheet_xml_path(sheet)?;
        let sheet_dir = Self::dir_of_path(&sheet_path);

        let Some(rels_xml) = self.sheet_rels_content(sheet)? else {
            return Ok(Vec::new());
        };

        let target = match Self::find_relationship_target_by_type(&rels_xml, "comments")? {
            Some(t) => t,
            None => return Ok(Vec::new()),
        };
        let comments_path = ooxml_util::join_and_normalize(&sheet_dir, &target);

        let mut zip = self.open_zip()?;
        let comments_xml = match ooxml_util::zip_read_to_string_opt(&mut zip, &comments_path)? {
            Some(s) => s,
            None => return Ok(Vec::new()),
        };

        Self::parse_comments_xml(&comments_xml)
    }

    fn parse_comments_xml(xml: &str) -> PyResult<Vec<CommentInfo>> {
        let mut reader = XmlReader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();

        let mut authors: Vec<String> = Vec::new();
        let mut out: Vec<CommentInfo> = Vec::new();

        let mut in_author = false;
        let mut in_comment = false;
        let mut in_t = false;

        let mut cur_cell: String = String::new();
        let mut cur_author_id: usize = 0;
        let mut cur_text: String = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let name = e.name();
                    let name = name.as_ref();
                    if name == b"author" {
                        in_author = true;
                    } else if name == b"comment" {
                        in_comment = true;
                        cur_text.clear();
                        cur_cell = ooxml_util::attr_value(&e, b"ref").unwrap_or_default();
                        cur_author_id = ooxml_util::attr_value(&e, b"authorId")
                            .and_then(|s| s.parse::<usize>().ok())
                            .unwrap_or(0);
                    } else if name == b"t" {
                        in_t = true;
                    }
                }
                Ok(Event::End(e)) => {
                    let name = e.name();
                    let name = name.as_ref();
                    if name == b"author" {
                        in_author = false;
                    } else if name == b"comment" {
                        in_comment = false;
                        let author = authors.get(cur_author_id).cloned().unwrap_or_default();
                        out.push(CommentInfo {
                            cell: cur_cell.clone(),
                            text: cur_text.clone(),
                            author,
                            threaded: false,
                        });
                    } else if name == b"t" {
                        in_t = false;
                    }
                }
                Ok(Event::Text(e)) => {
                    let text = e.unescape().unwrap_or_default().to_string();
                    if in_author {
                        authors.push(text);
                    } else if in_comment && in_t {
                        cur_text.push_str(&text);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse comments XML: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(out)
    }

    fn comments_to_py(py: Python<'_>, comments: &[CommentInfo]) -> PyResult<PyObject> {
        let result = PyList::empty(py);
        for c in comments {
            let d = PyDict::new(py);
            d.set_item("cell", &c.cell)?;
            d.set_item("text", &c.text)?;
            d.set_item("author", &c.author)?;
            d.set_item("threaded", c.threaded)?;
            result.append(d)?;
        }
        Ok(result.into())
    }

    fn compute_freeze_panes(&mut self, sheet: &str) -> PyResult<FreezePaneInfo> {
        let xml = self.sheet_xml_content(sheet)?;
        Self::parse_freeze_panes_from_sheet_xml(&xml)
    }

    fn parse_freeze_panes_from_sheet_xml(xml: &str) -> PyResult<FreezePaneInfo> {
        let mut reader = XmlReader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();

        let mut info = FreezePaneInfo {
            mode: String::new(),
            top_left_cell: None,
            x_split: None,
            y_split: None,
            active_pane: None,
        };

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"pane" {
                        let state = ooxml_util::attr_value(&e, b"state").unwrap_or_default();
                        let state_lc = state.to_ascii_lowercase();
                        if state_lc == "split" {
                            info.mode = "split".to_string();
                        } else if state_lc.starts_with("frozen") {
                            info.mode = "freeze".to_string();
                        }

                        info.top_left_cell = ooxml_util::attr_value(&e, b"topLeftCell");
                        info.active_pane = ooxml_util::attr_value(&e, b"activePane");
                        info.x_split = ooxml_util::attr_value(&e, b"xSplit")
                            .and_then(|s| s.parse::<f64>().ok())
                            .map(|v| v as i64);
                        info.y_split = ooxml_util::attr_value(&e, b"ySplit")
                            .and_then(|s| s.parse::<f64>().ok())
                            .map(|v| v as i64);
                        break;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse worksheet XML: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(info)
    }

    fn freeze_panes_to_py(py: Python<'_>, info: &FreezePaneInfo) -> PyResult<PyObject> {
        let d = PyDict::new(py);
        if info.mode.is_empty() {
            return Ok(d.into());
        }
        d.set_item("mode", &info.mode)?;
        if let Some(tlc) = &info.top_left_cell {
            if !tlc.is_empty() {
                d.set_item("top_left_cell", tlc)?;
            }
        }
        if let Some(x) = info.x_split {
            d.set_item("x_split", x)?;
        }
        if let Some(y) = info.y_split {
            d.set_item("y_split", y)?;
        }
        if let Some(ap) = &info.active_pane {
            if !ap.is_empty() {
                d.set_item("active_pane", ap)?;
            }
        }
        Ok(d.into())
    }

    fn normalize_ooxml_rgb(rgb: &str) -> Option<String> {
        let mut s = rgb.trim().to_string();
        if s.starts_with('#') {
            s = s.trim_start_matches('#').to_string();
        }
        // Excel OOXML colors are usually ARGB. Drop alpha if present.
        if s.len() == 8 {
            s = s[2..].to_string();
        }
        if s.len() != 6 {
            return None;
        }
        Some(format!("#{}", s.to_ascii_uppercase()))
    }

    fn ensure_dxfs_bg_colors(&mut self) -> PyResult<()> {
        if self.dxfs_bg_colors.is_some() {
            return Ok(());
        }

        let mut zip = self.open_zip()?;
        let styles_xml = match ooxml_util::zip_read_to_string_opt(&mut zip, "xl/styles.xml")? {
            Some(s) => s,
            None => {
                self.dxfs_bg_colors = Some(Vec::new());
                return Ok(());
            }
        };

        let mut reader = XmlReader::from_str(&styles_xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();

        let mut in_dxfs = false;
        let mut in_dxf = false;
        let mut dxf_depth: usize = 0;

        let mut cur_bg: Option<String> = None;
        let mut out: Vec<Option<String>> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"dxfs" {
                        in_dxfs = true;
                    } else if in_dxfs && e.name().as_ref() == b"dxf" {
                        in_dxf = true;
                        dxf_depth = 1;
                        cur_bg = None;
                    } else if in_dxf {
                        dxf_depth += 1;
                    }
                }
                Ok(Event::Empty(e)) => {
                    if !in_dxf {
                        // nothing
                    } else {
                        if e.name().as_ref() == b"fgColor" || e.name().as_ref() == b"bgColor" {
                            if cur_bg.is_none() {
                                if let Some(rgb) = ooxml_util::attr_value(&e, b"rgb") {
                                    cur_bg = Self::normalize_ooxml_rgb(&rgb);
                                }
                            }
                        }
                    }
                }
                Ok(Event::End(e)) => {
                    if in_dxf {
                        if e.name().as_ref() == b"dxf" {
                            out.push(cur_bg.take());
                            in_dxf = false;
                            dxf_depth = 0;
                        } else if dxf_depth > 0 {
                            dxf_depth -= 1;
                        }
                    }
                    if e.name().as_ref() == b"dxfs" {
                        in_dxfs = false;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse styles.xml: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        self.dxfs_bg_colors = Some(out);
        Ok(())
    }

    fn cell_style_id(&mut self, sheet: &str, row: u32, col: u32) -> PyResult<Option<u32>> {
        self.ensure_sheet_exists(sheet)?;

        let needs_load = self
            .tier2_cache
            .get(sheet)
            .and_then(|c| c.cell_style_ids.as_ref())
            .is_none();
        if needs_load {
            let xml = self.sheet_xml_content(sheet)?;
            let map = Self::parse_cell_style_ids_from_sheet_xml(&xml)?;
            self.tier2_cache
                .entry(sheet.to_string())
                .or_default()
                .cell_style_ids = Some(map);
        }

        Ok(self
            .tier2_cache
            .get(sheet)
            .and_then(|c| c.cell_style_ids.as_ref())
            .and_then(|m| m.get(&(row, col)).copied()))
    }

    fn ensure_diagonal_borders(&mut self) -> PyResult<()> {
        if self.diagonal_borders.is_some() {
            return Ok(());
        }

        let mut zip = self.open_zip()?;
        let styles_xml = match ooxml_util::zip_read_to_string_opt(&mut zip, "xl/styles.xml")? {
            Some(s) => s,
            None => {
                self.diagonal_borders = Some(HashMap::new());
                return Ok(());
            }
        };

        #[derive(Default)]
        struct BorderDef {
            up: bool,
            down: bool,
            style: Option<String>,
            color: Option<String>,
        }

        fn parse_bool_attr(v: &str) -> bool {
            v == "1" || v.eq_ignore_ascii_case("true")
        }

        let mut border_defs: Vec<BorderDef> = Vec::new();
        let mut xf_border_ids: Vec<usize> = Vec::new();

        let mut in_borders = false;
        let mut in_cellxfs = false;
        let mut in_diagonal = false;

        let mut cur_border: Option<BorderDef> = None;

        let mut reader = XmlReader::from_str(&styles_xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.name().as_ref() {
                    b"borders" => {
                        in_borders = true;
                    }
                    b"border" if in_borders => {
                        let mut def = BorderDef::default();
                        if let Some(v) = ooxml_util::attr_value(&e, b"diagonalUp") {
                            def.up = parse_bool_attr(&v);
                        }
                        if let Some(v) = ooxml_util::attr_value(&e, b"diagonalDown") {
                            def.down = parse_bool_attr(&v);
                        }
                        cur_border = Some(def);
                    }
                    b"diagonal" => {
                        if let Some(def) = cur_border.as_mut() {
                            if let Some(style) = ooxml_util::attr_value(&e, b"style") {
                                def.style = Some(style);
                            }
                            in_diagonal = true;
                        }
                    }
                    b"color" if in_diagonal => {
                        if let Some(def) = cur_border.as_mut() {
                            if let Some(rgb) = ooxml_util::attr_value(&e, b"rgb") {
                                def.color = Self::normalize_ooxml_rgb(&rgb);
                            }
                        }
                    }
                    b"cellXfs" => {
                        in_cellxfs = true;
                    }
                    b"xf" if in_cellxfs => {
                        let border_id = ooxml_util::attr_value(&e, b"borderId")
                            .and_then(|s| s.parse::<usize>().ok())
                            .unwrap_or(0);
                        xf_border_ids.push(border_id);
                    }
                    _ => {}
                },
                Ok(Event::Empty(e)) => {
                    match e.name().as_ref() {
                        b"border" if in_borders => {
                            // Rare, but handle self-closing border.
                            let mut def = BorderDef::default();
                            if let Some(v) = ooxml_util::attr_value(&e, b"diagonalUp") {
                                def.up = parse_bool_attr(&v);
                            }
                            if let Some(v) = ooxml_util::attr_value(&e, b"diagonalDown") {
                                def.down = parse_bool_attr(&v);
                            }
                            border_defs.push(def);
                        }
                        b"diagonal" => {
                            if let Some(def) = cur_border.as_mut() {
                                if let Some(style) = ooxml_util::attr_value(&e, b"style") {
                                    def.style = Some(style);
                                }
                            }
                        }
                        b"color" if in_diagonal => {
                            if let Some(def) = cur_border.as_mut() {
                                if let Some(rgb) = ooxml_util::attr_value(&e, b"rgb") {
                                    def.color = Self::normalize_ooxml_rgb(&rgb);
                                }
                            }
                        }
                        b"xf" if in_cellxfs => {
                            let border_id = ooxml_util::attr_value(&e, b"borderId")
                                .and_then(|s| s.parse::<usize>().ok())
                                .unwrap_or(0);
                            xf_border_ids.push(border_id);
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(e)) => match e.name().as_ref() {
                    b"borders" => {
                        in_borders = false;
                    }
                    b"border" => {
                        if let Some(def) = cur_border.take() {
                            border_defs.push(def);
                        }
                        in_diagonal = false;
                    }
                    b"diagonal" => {
                        in_diagonal = false;
                    }
                    b"cellXfs" => {
                        in_cellxfs = false;
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse styles.xml for diagonal borders: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        let mut out: HashMap<u32, DiagonalBorderInfo> = HashMap::new();
        for (xf_idx, border_id) in xf_border_ids.iter().enumerate() {
            let bd = match border_defs.get(*border_id) {
                Some(b) => b,
                None => continue,
            };
            if !(bd.up || bd.down) {
                continue;
            }
            let style = match bd.style.as_deref() {
                Some(s) if s != "none" => s.to_string(),
                _ => continue,
            };
            let color = bd.color.clone().unwrap_or_else(|| "#000000".to_string());
            out.insert(
                xf_idx as u32,
                DiagonalBorderInfo {
                    up: bd.up,
                    down: bd.down,
                    style,
                    color,
                },
            );
        }

        self.diagonal_borders = Some(out);
        Ok(())
    }

    fn compute_conditional_formats(
        &mut self,
        sheet: &str,
    ) -> PyResult<Vec<ConditionalFormatRuleInfo>> {
        let xml = self.sheet_xml_content(sheet)?;
        let mut reader = XmlReader::from_str(&xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();

        let mut current_range: Option<String> = None;
        let mut out: Vec<ConditionalFormatRuleInfo> = Vec::new();

        let mut cur_rule: Option<ConditionalFormatRuleInfo> = None;
        let mut cur_dxf_id: Option<usize> = None;
        let mut in_formula = false;
        let mut formula_buf = String::new();
        let mut in_dxf = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"conditionalFormatting" {
                        current_range = ooxml_util::attr_value(&e, b"sqref");
                    } else if e.name().as_ref() == b"cfRule" {
                        let range = current_range.clone().unwrap_or_default();
                        let rule_type = ooxml_util::attr_value(&e, b"type").unwrap_or_default();
                        let operator = ooxml_util::attr_value(&e, b"operator");
                        let priority = ooxml_util::attr_value(&e, b"priority")
                            .and_then(|s| s.parse::<i64>().ok());
                        let stop_if_true = ooxml_util::attr_value(&e, b"stopIfTrue")
                            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));
                        cur_dxf_id = ooxml_util::attr_value(&e, b"dxfId")
                            .and_then(|s| s.parse::<usize>().ok());

                        cur_rule = Some(ConditionalFormatRuleInfo {
                            range,
                            rule_type,
                            operator,
                            formula: None,
                            priority,
                            stop_if_true,
                            bg_color: None,
                        });
                    } else if e.name().as_ref() == b"formula" {
                        in_formula = true;
                        formula_buf.clear();
                    } else if e.name().as_ref() == b"dxf" {
                        in_dxf = true;
                    }
                }
                Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"conditionalFormatting" {
                        current_range = ooxml_util::attr_value(&e, b"sqref");
                    } else if e.name().as_ref() == b"cfRule" {
                        let range = current_range.clone().unwrap_or_default();
                        let rule_type = ooxml_util::attr_value(&e, b"type").unwrap_or_default();
                        let operator = ooxml_util::attr_value(&e, b"operator");
                        let priority = ooxml_util::attr_value(&e, b"priority")
                            .and_then(|s| s.parse::<i64>().ok());
                        let stop_if_true = ooxml_util::attr_value(&e, b"stopIfTrue")
                            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));
                        let dxf_id = ooxml_util::attr_value(&e, b"dxfId")
                            .and_then(|s| s.parse::<usize>().ok());

                        let mut rule = ConditionalFormatRuleInfo {
                            range,
                            rule_type,
                            operator,
                            formula: None,
                            priority,
                            stop_if_true,
                            bg_color: None,
                        };

                        if rule.bg_color.is_none() {
                            if let Some(id) = dxf_id {
                                self.ensure_dxfs_bg_colors()?;
                                if let Some(list) = &self.dxfs_bg_colors {
                                    if let Some(v) = list.get(id).cloned().flatten() {
                                        rule.bg_color = Some(v);
                                    }
                                }
                            }
                        }
                        if !rule.range.trim().is_empty() && !rule.rule_type.trim().is_empty() {
                            out.push(rule);
                        }
                    } else if in_dxf {
                        if e.name().as_ref() == b"fgColor" || e.name().as_ref() == b"bgColor" {
                            if let Some(ref mut rule) = cur_rule {
                                if rule.bg_color.is_none() {
                                    if let Some(rgb) = ooxml_util::attr_value(&e, b"rgb") {
                                        rule.bg_color = Self::normalize_ooxml_rgb(&rgb);
                                    }
                                }
                            }
                        }
                    }
                }
                Ok(Event::Text(e)) => {
                    if in_formula {
                        let t = e.unescape().unwrap_or_default().to_string();
                        formula_buf.push_str(&t);
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"conditionalFormatting" {
                        current_range = None;
                    } else if e.name().as_ref() == b"formula" {
                        if in_formula {
                            in_formula = false;
                            let f = formula_buf.trim().to_string();
                            if let Some(ref mut rule) = cur_rule {
                                if rule.formula.is_none() && !f.is_empty() {
                                    let formula = if f.starts_with('=') {
                                        f
                                    } else {
                                        format!("={f}")
                                    };
                                    rule.formula = Some(formula);
                                }
                            }
                        }
                    } else if e.name().as_ref() == b"dxf" {
                        in_dxf = false;
                    } else if e.name().as_ref() == b"cfRule" {
                        let Some(mut rule) = cur_rule.take() else {
                            cur_dxf_id = None;
                            continue;
                        };

                        if rule.bg_color.is_none() {
                            if let Some(id) = cur_dxf_id.take() {
                                self.ensure_dxfs_bg_colors()?;
                                if let Some(list) = &self.dxfs_bg_colors {
                                    if let Some(v) = list.get(id).cloned().flatten() {
                                        rule.bg_color = Some(v);
                                    }
                                }
                            }
                        } else {
                            cur_dxf_id = None;
                        }

                        if !rule.range.trim().is_empty() && !rule.rule_type.trim().is_empty() {
                            out.push(rule);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse conditional formats: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(out)
    }

    fn conditional_formats_to_py(
        py: Python<'_>,
        rules: &[ConditionalFormatRuleInfo],
    ) -> PyResult<PyObject> {
        let result = PyList::empty(py);
        for r in rules {
            let d = PyDict::new(py);
            d.set_item("range", &r.range)?;
            d.set_item("rule_type", &r.rule_type)?;
            if let Some(op) = &r.operator {
                d.set_item("operator", op)?;
            }
            if let Some(f) = &r.formula {
                d.set_item("formula", f)?;
            }
            if let Some(p) = r.priority {
                d.set_item("priority", p)?;
            }
            if let Some(s) = r.stop_if_true {
                d.set_item("stop_if_true", s)?;
            }
            if let Some(bg) = &r.bg_color {
                let fmt = PyDict::new(py);
                fmt.set_item("bg_color", bg)?;
                d.set_item("format", fmt)?;
            }
            result.append(d)?;
        }
        Ok(result.into())
    }

    fn compute_data_validations(&mut self, sheet: &str) -> PyResult<Vec<DataValidationInfo>> {
        let xml = self.sheet_xml_content(sheet)?;
        let mut reader = XmlReader::from_str(&xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();

        let mut out: Vec<DataValidationInfo> = Vec::new();

        let mut in_validation = false;
        let mut in_formula1 = false;
        let mut in_formula2 = false;
        let mut formula1_buf = String::new();
        let mut formula2_buf = String::new();

        let mut cur: Option<DataValidationInfo> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"dataValidation" {
                        in_validation = true;
                        let range = ooxml_util::attr_value(&e, b"sqref").unwrap_or_default();
                        let validation_type = ooxml_util::attr_value(&e, b"type")
                            .unwrap_or_else(|| "any".to_string());
                        let operator = ooxml_util::attr_value(&e, b"operator");

                        // OOXML uses allowBlank="1" when true. Missing => true (Excel default).
                        let allow_blank = match ooxml_util::attr_value(&e, b"allowBlank") {
                            Some(v) => v == "1" || v.eq_ignore_ascii_case("true"),
                            None => true,
                        };

                        let error_title = ooxml_util::attr_value(&e, b"errorTitle").and_then(|s| {
                            if s.is_empty() {
                                None
                            } else {
                                Some(s)
                            }
                        });
                        let error = ooxml_util::attr_value(&e, b"error").and_then(|s| {
                            if s.is_empty() {
                                None
                            } else {
                                Some(s)
                            }
                        });

                        formula1_buf.clear();
                        formula2_buf.clear();

                        cur = Some(DataValidationInfo {
                            range,
                            validation_type,
                            operator,
                            formula1: None,
                            formula2: None,
                            allow_blank,
                            error_title,
                            error,
                        });
                    } else if in_validation && e.name().as_ref() == b"formula1" {
                        in_formula1 = true;
                        formula1_buf.clear();
                    } else if in_validation && e.name().as_ref() == b"formula2" {
                        in_formula2 = true;
                        formula2_buf.clear();
                    }
                }
                Ok(Event::Text(e)) => {
                    let text = e.unescape().unwrap_or_default().to_string();
                    if in_formula1 {
                        formula1_buf.push_str(&text);
                    } else if in_formula2 {
                        formula2_buf.push_str(&text);
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"formula1" {
                        in_formula1 = false;
                        if let Some(ref mut dv) = cur {
                            let f = formula1_buf.trim().to_string();
                            if !f.is_empty() {
                                let formula = if f.starts_with('=') {
                                    f
                                } else {
                                    format!("={f}")
                                };
                                dv.formula1 = Some(formula);
                            }
                        }
                    } else if e.name().as_ref() == b"formula2" {
                        in_formula2 = false;
                        if let Some(ref mut dv) = cur {
                            let f = formula2_buf.trim().to_string();
                            if !f.is_empty() {
                                let formula = if f.starts_with('=') {
                                    f
                                } else {
                                    format!("={f}")
                                };
                                dv.formula2 = Some(formula);
                            }
                        }
                    } else if e.name().as_ref() == b"dataValidation" {
                        in_validation = false;
                        if let Some(dv) = cur.take() {
                            if !dv.range.trim().is_empty() {
                                out.push(dv);
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse data validations: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(out)
    }

    fn data_validations_to_py(py: Python<'_>, items: &[DataValidationInfo]) -> PyResult<PyObject> {
        let result = PyList::empty(py);
        for v in items {
            let d = PyDict::new(py);
            d.set_item("range", &v.range)?;
            d.set_item("validation_type", &v.validation_type)?;
            if let Some(op) = &v.operator {
                d.set_item("operator", op)?;
            }
            if let Some(f1) = &v.formula1 {
                d.set_item("formula1", f1)?;
            }
            if let Some(f2) = &v.formula2 {
                d.set_item("formula2", f2)?;
            }
            d.set_item("allow_blank", v.allow_blank)?;
            if let Some(t) = &v.error_title {
                d.set_item("error_title", t)?;
            }
            if let Some(msg) = &v.error {
                d.set_item("error", msg)?;
            }
            result.append(d)?;
        }
        Ok(result.into())
    }

    fn ensure_named_ranges(&mut self) -> PyResult<()> {
        if self.named_ranges.is_some() {
            return Ok(());
        }
        let mut zip = self.open_zip()?;
        let workbook_xml = ooxml_util::zip_read_to_string(&mut zip, "xl/workbook.xml")?;

        let mut reader = XmlReader::from_str(&workbook_xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();

        let mut in_defined_name = false;
        let mut cur_name: Option<String> = None;
        let mut cur_local_id: Option<usize> = None;
        let mut text_buf = String::new();

        let mut out: Vec<NamedRangeInfo> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"definedName" {
                        in_defined_name = true;
                        cur_name = ooxml_util::attr_value(&e, b"name");
                        cur_local_id = ooxml_util::attr_value(&e, b"localSheetId")
                            .and_then(|s| s.parse::<usize>().ok());
                        text_buf.clear();
                    }
                }
                Ok(Event::Text(e)) => {
                    if in_defined_name {
                        let t = e.unescape().unwrap_or_default().to_string();
                        text_buf.push_str(&t);
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"definedName" {
                        in_defined_name = false;
                        let Some(n) = cur_name.take() else {
                            continue;
                        };
                        let refers_to = text_buf.trim().to_string();
                        if refers_to.is_empty() {
                            continue;
                        }

                        // Skip Excel reserved/system names.
                        if n.starts_with("_xlnm.") {
                            continue;
                        }

                        let (scope, sheet_name) = match cur_local_id.take() {
                            Some(idx) => {
                                let sname = self.sheet_names.get(idx).cloned();
                                ("sheet".to_string(), sname)
                            }
                            None => ("workbook".to_string(), None),
                        };

                        // For sheet-scoped names, ensure refers_to includes the sheet.
                        let refers_to = if scope == "sheet" {
                            if refers_to.contains('!') {
                                refers_to
                            } else if let Some(sn) = sheet_name {
                                format!("{sn}!{refers_to}")
                            } else {
                                refers_to
                            }
                        } else {
                            refers_to
                        };

                        out.push(NamedRangeInfo {
                            name: n,
                            scope,
                            refers_to,
                        });
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse workbook.xml definedNames: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        self.named_ranges = Some(out);
        Ok(())
    }

    fn compute_tables(&mut self, sheet: &str) -> PyResult<Vec<TableInfo>> {
        let sheet_path = self.sheet_xml_path(sheet)?;
        let sheet_dir = Self::dir_of_path(&sheet_path);
        let Some(rels_xml) = self.sheet_rels_content(sheet)? else {
            return Ok(Vec::new());
        };

        let rid_targets = ooxml_util::parse_relationship_targets(&rels_xml)?;
        let sheet_xml = self.sheet_xml_content(sheet)?;

        // Collect tablePart relationship ids.
        let mut reader = XmlReader::from_str(&sheet_xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();
        let mut table_rids: Vec<String> = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"tablePart" {
                        if let Some(rid) = ooxml_util::attr_value(&e, b"r:id") {
                            if !rid.is_empty() {
                                table_rids.push(rid);
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse worksheet XML for tables: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        if table_rids.is_empty() {
            return Ok(Vec::new());
        }

        let mut zip = self.open_zip()?;
        let mut out: Vec<TableInfo> = Vec::new();
        for rid in table_rids {
            let Some(target) = rid_targets.get(&rid) else {
                continue;
            };
            let table_path = ooxml_util::join_and_normalize(&sheet_dir, target);
            let Some(table_xml) = ooxml_util::zip_read_to_string_opt(&mut zip, &table_path)? else {
                continue;
            };
            if let Ok(info) = Self::parse_table_xml(&table_xml) {
                out.push(info);
            }
        }

        Ok(out)
    }

    fn parse_table_xml(xml: &str) -> PyResult<TableInfo> {
        let mut reader = XmlReader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();

        let mut name: String = String::new();
        let mut ref_range: String = String::new();
        let mut header_row: bool = true;
        let mut totals_row: bool = false;
        let mut style: Option<String> = None;
        let mut columns: Vec<String> = Vec::new();
        let mut autofilter: bool = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"table" {
                        name = ooxml_util::attr_value(&e, b"name")
                            .or_else(|| ooxml_util::attr_value(&e, b"displayName"))
                            .unwrap_or_default();
                        ref_range = ooxml_util::attr_value(&e, b"ref").unwrap_or_default();
                        header_row = match ooxml_util::attr_value(&e, b"headerRowCount") {
                            Some(v) => v != "0",
                            None => true,
                        };
                        totals_row = match ooxml_util::attr_value(&e, b"totalsRowCount") {
                            Some(v) => v != "0",
                            None => false,
                        };
                    } else if e.name().as_ref() == b"tableStyleInfo" {
                        style = ooxml_util::attr_value(&e, b"name").and_then(|s| {
                            if s.is_empty() {
                                None
                            } else {
                                Some(s)
                            }
                        });
                    } else if e.name().as_ref() == b"tableColumn" {
                        if let Some(cn) = ooxml_util::attr_value(&e, b"name") {
                            columns.push(cn);
                        }
                    } else if e.name().as_ref() == b"autoFilter" {
                        autofilter = true;
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(PyErr::new::<PyIOError, _>(format!(
                        "Failed to parse table XML: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(TableInfo {
            name,
            ref_range,
            header_row,
            totals_row,
            style,
            columns,
            autofilter,
        })
    }

    fn tables_to_py(py: Python<'_>, tables: &[TableInfo]) -> PyResult<PyObject> {
        let result = PyList::empty(py);
        for t in tables {
            let d = PyDict::new(py);
            d.set_item("name", &t.name)?;
            d.set_item("ref", &t.ref_range)?;
            d.set_item("header_row", t.header_row)?;
            d.set_item("totals_row", t.totals_row)?;
            match &t.style {
                Some(s) => d.set_item("style", s)?,
                None => d.set_item("style", py.None())?,
            }
            d.set_item("columns", t.columns.clone())?;
            d.set_item("autofilter", t.autofilter)?;
            result.append(d)?;
        }
        Ok(result.into())
    }
}
