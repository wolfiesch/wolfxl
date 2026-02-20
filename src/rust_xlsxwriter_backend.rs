use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{Read, Write};

use indexmap::IndexMap;

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader as XmlReader, Writer as XmlWriter};

use rust_xlsxwriter::{
    Color, ConditionalFormat3ColorScale, ConditionalFormatCell, ConditionalFormatCellRule,
    ConditionalFormatDataBar, ConditionalFormatFormula, DataValidation, DataValidationRule, Format,
    FormatAlign, FormatBorder, FormatPattern, Formula, Note, Table, TableColumn, TableStyle, Url,
    Workbook, Worksheet,
};

use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::ooxml_util;
use crate::util::{a1_to_row_col, parse_iso_date, parse_iso_datetime};

// ---------------------------------------------------------------------------
// Queued operation types
// ---------------------------------------------------------------------------

/// Stored cell value payload (mirrors the Python dict contract).
struct CellPayload {
    type_str: String,
    value: Option<String>,
    formula: Option<String>,
}

/// Queued per-cell format dict fields.
struct FormatFields {
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<String>,
    strikethrough: Option<bool>,
    font_name: Option<String>,
    font_size: Option<f64>,
    font_color: Option<String>,
    bg_color: Option<String>,
    number_format: Option<String>,
    h_align: Option<String>,
    v_align: Option<String>,
    wrap: Option<bool>,
    rotation: Option<i32>,
    indent: Option<i32>,
}

/// Queued per-cell border dict fields.
struct BorderFields {
    top_style: Option<String>,
    top_color: Option<String>,
    bottom_style: Option<String>,
    bottom_color: Option<String>,
    left_style: Option<String>,
    left_color: Option<String>,
    right_style: Option<String>,
    right_color: Option<String>,
    diagonal_up_style: Option<String>,
    diagonal_up_color: Option<String>,
    diagonal_down_style: Option<String>,
    diagonal_down_color: Option<String>,
}

struct MergeRange {
    sheet: String,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
}

struct HyperlinkPayload {
    sheet: String,
    row: u32,
    col: u16,
    url: String,
    display: Option<String>,
    tooltip: Option<String>,
}

struct CommentPayload {
    sheet: String,
    row: u32,
    col: u16,
    text: String,
    author: Option<String>,
}

struct ConditionalFormatPayload {
    sheet: String,
    range: String,
    rule_type: String,
    operator: Option<String>,
    formula: Option<String>,
    stop_if_true: bool,
    bg_color: Option<String>,
}

struct DataValidationPayload {
    sheet: String,
    range: String,
    validation_type: String,
    operator: Option<String>,
    formula1: Option<String>,
    formula2: Option<String>,
    allow_blank: Option<bool>,
    error_title: Option<String>,
    error: Option<String>,
}

struct NamedRangePayload {
    name: String,
    scope: String,
    sheet: Option<String>,
    refers_to: String,
}

struct TablePayload {
    sheet: String,
    name: String,
    ref_range: String,
    style: Option<String>,
    columns: Vec<String>,
    totals_row: bool,
    autofilter: Option<bool>,
    header_row: bool,
}

enum PaneSetting {
    Freeze { row: u32, col: u16 },
    Split { x_split: f64, y_split: f64 },
}

type CellKey = (String, u32, u16); // (sheet, row, col)

#[pyclass(unsendable)]
pub struct RustXlsxWriterBook {
    sheet_names: Vec<String>,
    values: IndexMap<CellKey, CellPayload>,
    formats: HashMap<CellKey, FormatFields>,
    borders: HashMap<CellKey, BorderFields>,
    row_heights: HashMap<(String, u32), f64>,
    col_widths: HashMap<(String, u16), f64>,
    merge_ranges: Vec<MergeRange>,
    hyperlinks: Vec<HyperlinkPayload>,
    comments: Vec<CommentPayload>,
    panes: HashMap<String, PaneSetting>,
    print_areas: HashMap<String, String>,
    conditional_formats: Vec<ConditionalFormatPayload>,
    data_validations: Vec<DataValidationPayload>,
    named_ranges: Vec<NamedRangePayload>,
    tables: Vec<TablePayload>,
    saved: bool,
}

// ---------------------------------------------------------------------------
// Color / enum helpers
// ---------------------------------------------------------------------------

fn parse_hex_color(hex: &str) -> Color {
    let s = hex.strip_prefix('#').unwrap_or(hex);
    if let Ok(n) = u32::from_str_radix(s, 16) {
        Color::RGB(n)
    } else {
        Color::Black
    }
}

fn map_h_align(s: &str) -> FormatAlign {
    match s.to_ascii_lowercase().as_str() {
        "left" => FormatAlign::Left,
        "center" | "centre" => FormatAlign::Center,
        "right" => FormatAlign::Right,
        "fill" => FormatAlign::Fill,
        "justify" => FormatAlign::Justify,
        "distributed" | "centercontinuous" => FormatAlign::CenterAcross,
        _ => FormatAlign::Left,
    }
}

fn map_v_align(s: &str) -> FormatAlign {
    match s.to_ascii_lowercase().as_str() {
        "top" => FormatAlign::Top,
        "center" | "centre" => FormatAlign::VerticalCenter,
        "bottom" => FormatAlign::Bottom,
        "justify" => FormatAlign::VerticalJustify,
        "distributed" => FormatAlign::VerticalDistributed,
        _ => FormatAlign::Bottom,
    }
}

fn map_border_style(s: &str) -> FormatBorder {
    match s.to_ascii_lowercase().as_str() {
        "thin" => FormatBorder::Thin,
        "medium" => FormatBorder::Medium,
        "thick" => FormatBorder::Thick,
        "double" => FormatBorder::Double,
        "dashed" => FormatBorder::Dashed,
        "dotted" => FormatBorder::Dotted,
        "hair" => FormatBorder::Hair,
        "mediumdashed" => FormatBorder::MediumDashed,
        "dashdot" => FormatBorder::DashDot,
        "mediumdashdot" => FormatBorder::MediumDashDot,
        "dashdotdot" => FormatBorder::DashDotDot,
        "mediumdashdotdot" => FormatBorder::MediumDashDotDot,
        "slantdashdot" => FormatBorder::SlantDashDot,
        "none" | "" => FormatBorder::None,
        _ => FormatBorder::Thin,
    }
}

fn map_underline(s: &str) -> rust_xlsxwriter::FormatUnderline {
    match s.to_ascii_lowercase().as_str() {
        "single" => rust_xlsxwriter::FormatUnderline::Single,
        "double" => rust_xlsxwriter::FormatUnderline::Double,
        "singleaccounting" => rust_xlsxwriter::FormatUnderline::SingleAccounting,
        "doubleaccounting" => rust_xlsxwriter::FormatUnderline::DoubleAccounting,
        _ => rust_xlsxwriter::FormatUnderline::Single,
    }
}

// ---------------------------------------------------------------------------
// Build a Format from optional FormatFields + BorderFields
// ---------------------------------------------------------------------------

fn build_format(fmt: Option<&FormatFields>, bdr: Option<&BorderFields>) -> PyResult<Format> {
    let mut f = Format::new();

    if let Some(ff) = fmt {
        if ff.bold == Some(true) {
            f = f.set_bold();
        }
        if ff.italic == Some(true) {
            f = f.set_italic();
        }
        if let Some(ref ul) = ff.underline {
            f = f.set_underline(map_underline(ul));
        }
        if ff.strikethrough == Some(true) {
            f = f.set_font_strikethrough();
        }
        if let Some(ref name) = ff.font_name {
            f = f.set_font_name(name);
        }
        if let Some(size) = ff.font_size {
            f = f.set_font_size(size);
        }
        if let Some(ref color) = ff.font_color {
            f = f.set_font_color(parse_hex_color(color));
        }
        if let Some(ref bg) = ff.bg_color {
            f = f.set_background_color(parse_hex_color(bg));
        }
        if let Some(ref nf) = ff.number_format {
            f = f.set_num_format(nf);
        }
        if let Some(ref h) = ff.h_align {
            f = f.set_align(map_h_align(h));
        }
        if let Some(ref v) = ff.v_align {
            f = f.set_align(map_v_align(v));
        }
        if ff.wrap == Some(true) {
            f = f.set_text_wrap();
        }
        if let Some(rot) = ff.rotation {
            let r: i16 = rot.try_into().map_err(|_| {
                PyErr::new::<PyValueError, _>(format!("Rotation value {rot} out of range for i16"))
            })?;
            f = f.set_rotation(r);
        }
        if let Some(indent) = ff.indent {
            if indent >= 0 && indent <= u8::MAX as i32 {
                f = f.set_indent(indent as u8);
            }
        }
    }

    if let Some(bb) = bdr {
        if let Some(ref s) = bb.top_style {
            f = f.set_border_top(map_border_style(s));
            if let Some(ref c) = bb.top_color {
                f = f.set_border_top_color(parse_hex_color(c));
            }
        }
        if let Some(ref s) = bb.bottom_style {
            f = f.set_border_bottom(map_border_style(s));
            if let Some(ref c) = bb.bottom_color {
                f = f.set_border_bottom_color(parse_hex_color(c));
            }
        }
        if let Some(ref s) = bb.left_style {
            f = f.set_border_left(map_border_style(s));
            if let Some(ref c) = bb.left_color {
                f = f.set_border_left_color(parse_hex_color(c));
            }
        }
        if let Some(ref s) = bb.right_style {
            f = f.set_border_right(map_border_style(s));
            if let Some(ref c) = bb.right_color {
                f = f.set_border_right_color(parse_hex_color(c));
            }
        }

        // Diagonal borders: if both up+down are present, use BorderUpDown.
        let has_up = bb.diagonal_up_style.is_some();
        let has_down = bb.diagonal_down_style.is_some();
        if has_up || has_down {
            // Use whichever is set (prefer down if both, since it's applied second).
            let (style_ref, color_ref) = if has_down {
                (&bb.diagonal_down_style, &bb.diagonal_down_color)
            } else {
                (&bb.diagonal_up_style, &bb.diagonal_up_color)
            };
            if let Some(ref s) = style_ref {
                f = f.set_border_diagonal(map_border_style(s));
            }
            if let Some(ref c) = color_ref {
                f = f.set_border_diagonal_color(parse_hex_color(c));
            }
            let diag_type = if has_up && has_down {
                rust_xlsxwriter::FormatDiagonalBorder::BorderUpDown
            } else if has_up {
                rust_xlsxwriter::FormatDiagonalBorder::BorderUp
            } else {
                rust_xlsxwriter::FormatDiagonalBorder::BorderDown
            };
            f = f.set_border_diagonal_type(diag_type);
        }
    }

    Ok(f)
}

// ---------------------------------------------------------------------------
// Extract fields from Python dicts
// ---------------------------------------------------------------------------

fn extract_format_fields(dict: &Bound<'_, PyDict>) -> PyResult<FormatFields> {
    Ok(FormatFields {
        bold: dict.get_item("bold")?.and_then(|v| v.extract().ok()),
        italic: dict.get_item("italic")?.and_then(|v| v.extract().ok()),
        underline: dict.get_item("underline")?.and_then(|v| v.extract().ok()),
        strikethrough: dict
            .get_item("strikethrough")?
            .and_then(|v| v.extract().ok()),
        font_name: dict.get_item("font_name")?.and_then(|v| v.extract().ok()),
        font_size: dict.get_item("font_size")?.and_then(|v| v.extract().ok()),
        font_color: dict.get_item("font_color")?.and_then(|v| v.extract().ok()),
        bg_color: dict.get_item("bg_color")?.and_then(|v| v.extract().ok()),
        number_format: dict
            .get_item("number_format")?
            .and_then(|v| v.extract().ok()),
        h_align: dict.get_item("h_align")?.and_then(|v| v.extract().ok()),
        v_align: dict.get_item("v_align")?.and_then(|v| v.extract().ok()),
        wrap: dict.get_item("wrap")?.and_then(|v| v.extract().ok()),
        rotation: dict.get_item("rotation")?.and_then(|v| v.extract().ok()),
        indent: dict.get_item("indent")?.and_then(|v| v.extract().ok()),
    })
}

fn extract_border_fields(dict: &Bound<'_, PyDict>) -> PyResult<BorderFields> {
    fn edge(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<(Option<String>, Option<String>)> {
        if let Some(sub) = dict.get_item(key)? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                let style: Option<String> = d.get_item("style")?.and_then(|v| v.extract().ok());
                let color: Option<String> = d.get_item("color")?.and_then(|v| v.extract().ok());
                return Ok((style, color));
            }
        }
        Ok((None, None))
    }

    let (ts, tc) = edge(dict, "top")?;
    let (bs, bc) = edge(dict, "bottom")?;
    let (ls, lc) = edge(dict, "left")?;
    let (rs, rc) = edge(dict, "right")?;
    let (dus, duc) = edge(dict, "diagonal_up")?;
    let (dds, ddc) = edge(dict, "diagonal_down")?;

    Ok(BorderFields {
        top_style: ts,
        top_color: tc,
        bottom_style: bs,
        bottom_color: bc,
        left_style: ls,
        left_color: lc,
        right_style: rs,
        right_color: rc,
        diagonal_up_style: dus,
        diagonal_up_color: duc,
        diagonal_down_style: dds,
        diagonal_down_color: ddc,
    })
}

fn resolve_key(sheet: &str, a1: &str) -> PyResult<CellKey> {
    let (row, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let col: u16 = col0.try_into().map_err(|_| {
        PyErr::new::<PyValueError, _>(format!("Column out of range for Excel: {a1}"))
    })?;
    Ok((sheet.to_string(), row, col))
}

fn parse_a1_range(range_str: &str) -> PyResult<(u32, u16, u32, u16)> {
    let clean = range_str.replace('$', "");
    let mut parts = clean.split(':');
    let a = parts.next().unwrap_or("");
    let b = parts.next().unwrap_or(a);

    let (r1, c1_32) = a1_to_row_col(a).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let (r2, c2_32) = a1_to_row_col(b).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;

    let c1: u16 = c1_32.try_into().map_err(|_| {
        PyErr::new::<PyValueError, _>(format!("Column out of range for Excel: {a}"))
    })?;
    let c2: u16 = c2_32.try_into().map_err(|_| {
        PyErr::new::<PyValueError, _>(format!("Column out of range for Excel: {b}"))
    })?;

    let (first_row, last_row) = if r1 <= r2 { (r1, r2) } else { (r2, r1) };
    let (first_col, last_col) = if c1 <= c2 { (c1, c2) } else { (c2, c1) };

    Ok((first_row, first_col, last_row, last_col))
}

fn col_letter_to_index(col_str: &str) -> PyResult<u16> {
    let mut col: u32 = 0;
    for ch in col_str.chars() {
        if !ch.is_ascii_alphabetic() {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Invalid column letter: {col_str}"
            )));
        }
        let uc = ch.to_ascii_uppercase() as u8;
        col = col * 26 + (uc - b'A' + 1) as u32;
    }
    if col == 0 {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Invalid column letter: {col_str}"
        )));
    }
    let idx = (col - 1) as u16;
    Ok(idx)
}

// ---------------------------------------------------------------------------
// Helper: write a single cell's value+format to a Worksheet
// ---------------------------------------------------------------------------

fn write_cell(
    ws: &mut Worksheet,
    row: u32,
    col: u16,
    payload: &CellPayload,
    format: &Format,
) -> PyResult<()> {
    fn parse_bool(s: &str) -> bool {
        match s.trim().to_ascii_lowercase().as_str() {
            "true" | "1" | "t" | "yes" | "y" => true,
            _ => false,
        }
    }

    match payload.type_str.as_str() {
        "blank" => {
            // Write blank with format so the format is preserved.
            ws.write_blank(row, col, format)
                .map(|_| ())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_blank failed: {e}")))
        }
        "string" => {
            let s = payload.value.as_deref().unwrap_or("");
            ws.write_string_with_format(row, col, s, format)
                .map(|_| ())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_string failed: {e}")))
        }
        "number" => {
            let f_val: f64 = payload
                .value
                .as_deref()
                .unwrap_or("0")
                .parse()
                .map_err(|_| PyErr::new::<PyValueError, _>("number parse failed"))?;
            ws.write_number_with_format(row, col, f_val, format)
                .map(|_| ())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_number failed: {e}")))
        }
        "boolean" => {
            let b = payload.value.as_deref().map(parse_bool).unwrap_or(false);
            ws.write_boolean_with_format(row, col, b, format)
                .map(|_| ())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_boolean failed: {e}")))
        }
        "formula" => {
            let formula = payload
                .formula
                .as_deref()
                .or(payload.value.as_deref())
                .unwrap_or("");
            ws.write_formula_with_format(row, col, formula, format)
                .map(|_| ())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_formula failed: {e}")))
        }
        "error" => {
            let token = payload.value.as_deref().unwrap_or("");
            let formula = match token {
                "#DIV/0!" => Some("=1/0"),
                "#N/A" => Some("=NA()"),
                "#VALUE!" => Some("=\"text\"+1"),
                _ => None,
            };
            if let Some(f) = formula {
                ws.write_formula_with_format(row, col, f, format)
                    .map(|_| ())
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_formula failed: {e}")))
            } else {
                ws.write_string_with_format(row, col, token, format)
                    .map(|_| ())
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_string failed: {e}")))
            }
        }
        "date" => {
            let s = payload.value.as_deref().unwrap_or("");
            if let Some(d) = parse_iso_date(s) {
                ws.write_datetime_with_format(row, col, d, format)
                    .map(|_| ())
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_datetime failed: {e}")))
            } else {
                ws.write_string_with_format(row, col, s, format)
                    .map(|_| ())
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_string failed: {e}")))
            }
        }
        "datetime" => {
            let s = payload.value.as_deref().unwrap_or("");
            if let Some(dt) = parse_iso_datetime(s) {
                ws.write_datetime_with_format(row, col, dt, format)
                    .map(|_| ())
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_datetime failed: {e}")))
            } else {
                ws.write_string_with_format(row, col, s, format)
                    .map(|_| ())
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_string failed: {e}")))
            }
        }
        other => Err(PyErr::new::<PyValueError, _>(format!(
            "Unsupported cell type: {other}"
        ))),
    }
}

// ---------------------------------------------------------------------------
// PyO3 implementation
// ---------------------------------------------------------------------------

impl RustXlsxWriterBook {
    fn ensure_sheet_exists(&self, sheet: &str) -> PyResult<()> {
        if self.sheet_names.contains(&sheet.to_string()) {
            Ok(())
        } else {
            Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )))
        }
    }
}

fn quote_sheet_name(sheet: &str) -> String {
    if sheet.contains(' ') || sheet.contains('\'') {
        let escaped = sheet.replace('\'', "''");
        format!("'{escaped}'")
    } else {
        sheet.to_string()
    }
}

fn map_table_style(style_name: Option<&str>) -> TableStyle {
    let Some(s) = style_name else {
        return TableStyle::None;
    };
    let raw = s.trim();
    if raw.is_empty() {
        return TableStyle::None;
    }
    let token = raw.strip_prefix("TableStyle").unwrap_or(raw);
    match token {
        "Medium9" => TableStyle::Medium9,
        "Medium2" => TableStyle::Medium2,
        "Light1" => TableStyle::Light1,
        _ => TableStyle::Medium9,
    }
}

fn map_cf_cell_rule(operator: &str, value_str: &str) -> PyResult<ConditionalFormatCellRule<i32>> {
    let value: i32 = value_str.trim().parse().map_err(|_| {
        PyErr::new::<PyValueError, _>(format!(
            "Conditional format numeric value parse failed: {value_str}"
        ))
    })?;

    let op = operator.to_ascii_lowercase();
    let rule = match op.as_str() {
        "greaterthan" => ConditionalFormatCellRule::GreaterThan(value),
        "greaterthanorequal" | "greaterthanorequalto" => {
            ConditionalFormatCellRule::GreaterThanOrEqualTo(value)
        }
        "lessthan" => ConditionalFormatCellRule::LessThan(value),
        "lessthanorequal" | "lessthanorequalto" => {
            ConditionalFormatCellRule::LessThanOrEqualTo(value)
        }
        "equal" | "equalto" => ConditionalFormatCellRule::EqualTo(value),
        "notequal" | "notequalto" => ConditionalFormatCellRule::NotEqualTo(value),
        other => {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unsupported conditional format operator: {other}"
            )))
        }
    };
    Ok(rule)
}

fn map_dv_rule_between_i32(formula1: &str, formula2: &str) -> PyResult<DataValidationRule<i32>> {
    let a: i32 = formula1.trim().parse().map_err(|_| {
        PyErr::new::<PyValueError, _>(format!("Data validation formula1 parse failed: {formula1}"))
    })?;
    let b: i32 = formula2.trim().parse().map_err(|_| {
        PyErr::new::<PyValueError, _>(format!("Data validation formula2 parse failed: {formula2}"))
    })?;
    Ok(DataValidationRule::Between(a, b))
}

// ---------------------------------------------------------------------------
// OOXML post-processing (split panes)
// ---------------------------------------------------------------------------

fn patch_sheet_xml_split_panes(xml: &str, x_split: i32, y_split: i32) -> PyResult<String> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::new());
    let mut buf: Vec<u8> = Vec::new();

    let mut in_sheet_view = false;
    let mut replaced_pane = false;
    let mut wrote_selection = false;
    let mut skip_depth: usize = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let is_pane = e.name().as_ref() == b"pane";
                let is_selection = e.name().as_ref() == b"selection";
                let is_sheet_view = e.name().as_ref() == b"sheetView";
                if skip_depth > 0 {
                    skip_depth += 1;
                } else if in_sheet_view && is_pane {
                    write_split_pane(&mut writer, x_split, y_split).map_err(|err| {
                        PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                    })?;
                    replaced_pane = true;
                    skip_depth = 1; // skip original <pane>...</pane>
                } else if in_sheet_view && is_selection {
                    skip_depth = 1; // drop all selections; we'll add one later
                } else {
                    if is_sheet_view {
                        in_sheet_view = true;
                    }
                    writer
                        .write_event(Event::Start(e.into_owned()))
                        .map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                        })?;
                }
            }
            Ok(Event::Empty(e)) => {
                let is_pane = e.name().as_ref() == b"pane";
                let is_selection = e.name().as_ref() == b"selection";
                let is_sheet_view = e.name().as_ref() == b"sheetView";
                if skip_depth > 0 {
                    // skip
                } else if in_sheet_view && is_pane {
                    write_split_pane(&mut writer, x_split, y_split).map_err(|err| {
                        PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                    })?;
                    replaced_pane = true;
                } else if in_sheet_view && is_selection {
                    // drop
                } else {
                    if is_sheet_view {
                        in_sheet_view = true;
                    }
                    writer
                        .write_event(Event::Empty(e.into_owned()))
                        .map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                        })?;
                    if is_sheet_view {
                        in_sheet_view = false;
                    }
                }
            }
            Ok(Event::End(e)) => {
                let is_sheet_view = e.name().as_ref() == b"sheetView";
                if skip_depth > 0 {
                    skip_depth -= 1;
                } else if is_sheet_view {
                    if !replaced_pane {
                        write_split_pane(&mut writer, x_split, y_split).map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                        })?;
                        replaced_pane = true;
                    }
                    if !wrote_selection {
                        write_default_split_selection(&mut writer).map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                        })?;
                        wrote_selection = true;
                    }
                    in_sheet_view = false;
                    writer
                        .write_event(Event::End(e.into_owned()))
                        .map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                        })?;
                } else {
                    writer
                        .write_event(Event::End(e.into_owned()))
                        .map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                        })?;
                }
            }
            Ok(Event::Text(e)) => {
                if skip_depth == 0 {
                    writer
                        .write_event(Event::Text(e.into_owned()))
                        .map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                        })?;
                }
            }
            Ok(Event::CData(e)) => {
                if skip_depth == 0 {
                    writer
                        .write_event(Event::CData(e.into_owned()))
                        .map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                        })?;
                }
            }
            Ok(Event::Comment(e)) => {
                if skip_depth == 0 {
                    writer
                        .write_event(Event::Comment(e.into_owned()))
                        .map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                        })?;
                }
            }
            Ok(Event::Decl(e)) => {
                writer
                    .write_event(Event::Decl(e.into_owned()))
                    .map_err(|err| PyErr::new::<PyIOError, _>(format!("XML write error: {err}")))?;
            }
            Ok(Event::PI(e)) => {
                if skip_depth == 0 {
                    writer
                        .write_event(Event::PI(e.into_owned()))
                        .map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                        })?;
                }
            }
            Ok(Event::DocType(e)) => {
                if skip_depth == 0 {
                    writer
                        .write_event(Event::DocType(e.into_owned()))
                        .map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                        })?;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(PyErr::new::<PyIOError, _>(format!(
                    "Failed to parse worksheet XML: {e}"
                )))
            }
        }
        buf.clear();
    }

    let out = writer.into_inner();
    String::from_utf8(out)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Worksheet XML not UTF-8: {e}")))
}

fn write_split_pane(
    writer: &mut XmlWriter<Vec<u8>>,
    x_split: i32,
    y_split: i32,
) -> std::io::Result<()> {
    let mut elem = BytesStart::new("pane");
    if x_split > 0 {
        let x_str = x_split.to_string();
        elem.push_attribute(("xSplit", x_str.as_str()));
    }
    if y_split > 0 {
        let y_str = y_split.to_string();
        elem.push_attribute(("ySplit", y_str.as_str()));
    }
    elem.push_attribute(("activePane", "topLeft"));
    elem.push_attribute(("state", "split"));
    writer.write_event(Event::Empty(elem))
}

fn write_default_split_selection(writer: &mut XmlWriter<Vec<u8>>) -> std::io::Result<()> {
    let mut elem = BytesStart::new("selection");
    elem.push_attribute(("activeCell", "A1"));
    elem.push_attribute(("sqref", "A1"));
    writer.write_event(Event::Empty(elem))
}

fn patch_split_panes_xlsx(path: &str, split_patches: &[(String, i32, i32)]) -> PyResult<()> {
    if split_patches.is_empty() {
        return Ok(());
    }

    // Build sheet name -> sheet XML path mapping.
    let f = File::open(path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open '{path}': {e}")))?;
    let mut zip = ZipArchive::new(f)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read xlsx zip: {e}")))?;

    let workbook_xml = ooxml_util::zip_read_to_string(&mut zip, "xl/workbook.xml")?;
    let rels_xml = ooxml_util::zip_read_to_string(&mut zip, "xl/_rels/workbook.xml.rels")?;
    let sheet_rids = ooxml_util::parse_workbook_sheet_rids(&workbook_xml)?;
    let rel_targets = ooxml_util::parse_relationship_targets(&rels_xml)?;

    let mut sheet_to_path: HashMap<String, String> = HashMap::new();
    for (name, rid) in sheet_rids {
        if let Some(target) = rel_targets.get(&rid) {
            sheet_to_path.insert(name, ooxml_util::join_and_normalize("xl/", target));
        }
    }

    // Generate patched worksheet XML contents.
    let mut file_patches: HashMap<String, Vec<u8>> = HashMap::new();
    for (sheet_name, x_split, y_split) in split_patches {
        let Some(sheet_path) = sheet_to_path.get(sheet_name) else {
            continue;
        };
        let xml = ooxml_util::zip_read_to_string(&mut zip, sheet_path)?;
        let patched = patch_sheet_xml_split_panes(&xml, *x_split, *y_split)?;
        file_patches.insert(sheet_path.clone(), patched.into_bytes());
    }
    drop(zip);

    if file_patches.is_empty() {
        return Ok(());
    }

    // Rewrite the zip with patched entries.
    let src = File::open(path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open '{path}': {e}")))?;
    let mut zip = ZipArchive::new(src)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read xlsx zip: {e}")))?;

    let tmp_path = format!("{path}.tmp");
    let dst = File::create(&tmp_path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to create '{tmp_path}': {e}")))?;
    let mut out = ZipWriter::new(dst);

    for i in 0..zip.len() {
        let mut file = zip.by_index(i).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Failed to read zip entry {i}: {e}"))
        })?;
        let name = file.name().to_string();

        let mut opts = SimpleFileOptions::default().compression_method(file.compression());
        if let Some(dt) = file.last_modified() {
            opts = opts.last_modified_time(dt);
        }
        if let Some(mode) = file.unix_mode() {
            opts = opts.unix_permissions(mode);
        }

        if file.is_dir() {
            out.add_directory(name, opts).map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("Failed to add directory to zip: {e}"))
            })?;
            continue;
        }

        let mut data: Vec<u8> = Vec::new();
        file.read_to_end(&mut data)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read zip entry: {e}")))?;
        if let Some(patched) = file_patches.get(&name) {
            data = patched.clone();
        }

        out.start_file(name, opts)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to write zip entry: {e}")))?;
        out.write_all(&data)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to write zip entry: {e}")))?;
    }

    out.finish()
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to finalize zip: {e}")))?;

    if let Err(e) = std::fs::rename(&tmp_path, path) {
        // On some platforms rename() may not replace; retry with explicit remove.
        let _ = std::fs::remove_file(path);
        std::fs::rename(&tmp_path, path).map_err(|e2| {
            PyErr::new::<PyIOError, _>(format!("Failed to replace file: {e}; {e2}"))
        })?;
    }

    Ok(())
}

fn extract_table_name(xml: &str) -> Option<String> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf: Vec<u8> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"table" {
                    let name = ooxml_util::attr_value(&e, b"name")
                        .or_else(|| ooxml_util::attr_value(&e, b"displayName"));
                    return name;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    None
}

fn patch_table_xml_ref(xml: &str, new_ref: &str) -> PyResult<String> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::new());
    let mut buf: Vec<u8> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name_bytes = e.name().as_ref().to_vec();
                let name_str = std::str::from_utf8(&name_bytes).unwrap_or("");

                if name_bytes.as_slice() == b"table" || name_bytes.as_slice() == b"autoFilter" {
                    let mut elem = BytesStart::new(name_str);
                    for a in e.attributes().with_checks(false) {
                        let a = a.map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML attr parse error: {err}"))
                        })?;
                        if a.key.as_ref() == b"ref" {
                            continue;
                        }
                        let k = std::str::from_utf8(a.key.as_ref()).unwrap_or("");
                        let v = a
                            .unescape_value()
                            .map_err(|err| {
                                PyErr::new::<PyIOError, _>(format!("XML attr decode error: {err}"))
                            })?
                            .into_owned();
                        if !k.is_empty() {
                            elem.push_attribute((k, v.as_str()));
                        }
                    }
                    elem.push_attribute(("ref", new_ref));
                    writer.write_event(Event::Start(elem)).map_err(|err| {
                        PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                    })?;
                } else {
                    writer
                        .write_event(Event::Start(e.into_owned()))
                        .map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                        })?;
                }
            }
            Ok(Event::Empty(e)) => {
                let name_bytes = e.name().as_ref().to_vec();
                let name_str = std::str::from_utf8(&name_bytes).unwrap_or("");

                if name_bytes.as_slice() == b"table" || name_bytes.as_slice() == b"autoFilter" {
                    let mut elem = BytesStart::new(name_str);
                    for a in e.attributes().with_checks(false) {
                        let a = a.map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML attr parse error: {err}"))
                        })?;
                        if a.key.as_ref() == b"ref" {
                            continue;
                        }
                        let k = std::str::from_utf8(a.key.as_ref()).unwrap_or("");
                        let v = a
                            .unescape_value()
                            .map_err(|err| {
                                PyErr::new::<PyIOError, _>(format!("XML attr decode error: {err}"))
                            })?
                            .into_owned();
                        if !k.is_empty() {
                            elem.push_attribute((k, v.as_str()));
                        }
                    }
                    elem.push_attribute(("ref", new_ref));
                    writer.write_event(Event::Empty(elem)).map_err(|err| {
                        PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                    })?;
                } else {
                    writer
                        .write_event(Event::Empty(e.into_owned()))
                        .map_err(|err| {
                            PyErr::new::<PyIOError, _>(format!("XML write error: {err}"))
                        })?;
                }
            }
            Ok(Event::End(e)) => {
                writer
                    .write_event(Event::End(e.into_owned()))
                    .map_err(|err| PyErr::new::<PyIOError, _>(format!("XML write error: {err}")))?;
            }
            Ok(Event::Text(e)) => {
                writer
                    .write_event(Event::Text(e.into_owned()))
                    .map_err(|err| PyErr::new::<PyIOError, _>(format!("XML write error: {err}")))?;
            }
            Ok(Event::CData(e)) => {
                writer
                    .write_event(Event::CData(e.into_owned()))
                    .map_err(|err| PyErr::new::<PyIOError, _>(format!("XML write error: {err}")))?;
            }
            Ok(Event::Comment(e)) => {
                writer
                    .write_event(Event::Comment(e.into_owned()))
                    .map_err(|err| PyErr::new::<PyIOError, _>(format!("XML write error: {err}")))?;
            }
            Ok(Event::Decl(e)) => {
                writer
                    .write_event(Event::Decl(e.into_owned()))
                    .map_err(|err| PyErr::new::<PyIOError, _>(format!("XML write error: {err}")))?;
            }
            Ok(Event::PI(e)) => {
                writer
                    .write_event(Event::PI(e.into_owned()))
                    .map_err(|err| PyErr::new::<PyIOError, _>(format!("XML write error: {err}")))?;
            }
            Ok(Event::DocType(e)) => {
                writer
                    .write_event(Event::DocType(e.into_owned()))
                    .map_err(|err| PyErr::new::<PyIOError, _>(format!("XML write error: {err}")))?;
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(PyErr::new::<PyIOError, _>(format!(
                    "Failed to parse table XML: {e}"
                )))
            }
        }
        buf.clear();
    }

    let out = writer.into_inner();
    String::from_utf8(out)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Table XML not UTF-8: {e}")))
}

fn patch_tables_xlsx(path: &str, ref_patches: &[(String, String)]) -> PyResult<()> {
    if ref_patches.is_empty() {
        return Ok(());
    }
    let patch_map: HashMap<String, String> = ref_patches.iter().cloned().collect();

    let src = File::open(path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open '{path}': {e}")))?;
    let mut zip = ZipArchive::new(src)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read xlsx zip: {e}")))?;

    let mut file_patches: HashMap<String, Vec<u8>> = HashMap::new();

    for i in 0..zip.len() {
        let mut file = zip.by_index(i).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Failed to read zip entry {i}: {e}"))
        })?;
        let name = file.name().to_string();
        if !name.starts_with("xl/tables/") || !name.ends_with(".xml") {
            continue;
        }

        let mut data: Vec<u8> = Vec::new();
        file.read_to_end(&mut data)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read zip entry: {e}")))?;
        let xml = String::from_utf8_lossy(&data).to_string();
        let Some(tname) = extract_table_name(&xml) else {
            continue;
        };
        let Some(new_ref) = patch_map.get(&tname) else {
            continue;
        };
        let patched = patch_table_xml_ref(&xml, new_ref)?;
        file_patches.insert(name, patched.into_bytes());
    }

    drop(zip);
    if file_patches.is_empty() {
        return Ok(());
    }

    // Rewrite the zip with patched table entries.
    let src = File::open(path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open '{path}': {e}")))?;
    let mut zip = ZipArchive::new(src)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read xlsx zip: {e}")))?;

    let tmp_path = format!("{path}.tmp");
    let dst = File::create(&tmp_path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to create '{tmp_path}': {e}")))?;
    let mut out = ZipWriter::new(dst);

    for i in 0..zip.len() {
        let mut file = zip.by_index(i).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Failed to read zip entry {i}: {e}"))
        })?;
        let name = file.name().to_string();

        let mut opts = SimpleFileOptions::default().compression_method(file.compression());
        if let Some(dt) = file.last_modified() {
            opts = opts.last_modified_time(dt);
        }
        if let Some(mode) = file.unix_mode() {
            opts = opts.unix_permissions(mode);
        }

        if file.is_dir() {
            out.add_directory(name, opts).map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("Failed to add directory to zip: {e}"))
            })?;
            continue;
        }

        let mut data: Vec<u8> = Vec::new();
        file.read_to_end(&mut data)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read zip entry: {e}")))?;
        if let Some(patched) = file_patches.get(&name) {
            data = patched.clone();
        }

        out.start_file(name, opts)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to write zip entry: {e}")))?;
        out.write_all(&data)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to write zip entry: {e}")))?;
    }

    out.finish()
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to finalize zip: {e}")))?;

    if let Err(e) = std::fs::rename(&tmp_path, path) {
        let _ = std::fs::remove_file(path);
        std::fs::rename(&tmp_path, path).map_err(|e2| {
            PyErr::new::<PyIOError, _>(format!("Failed to replace file: {e}; {e2}"))
        })?;
    }

    Ok(())
}

#[pymethods]
impl RustXlsxWriterBook {
    #[new]
    pub fn new() -> Self {
        Self {
            sheet_names: Vec::new(),
            values: IndexMap::new(),
            formats: HashMap::new(),
            borders: HashMap::new(),
            row_heights: HashMap::new(),
            col_widths: HashMap::new(),
            merge_ranges: Vec::new(),
            hyperlinks: Vec::new(),
            comments: Vec::new(),
            panes: HashMap::new(),
            print_areas: HashMap::new(),
            conditional_formats: Vec::new(),
            data_validations: Vec::new(),
            named_ranges: Vec::new(),
            tables: Vec::new(),
            saved: false,
        }
    }

    pub fn add_sheet(&mut self, name: &str) -> PyResult<()> {
        if self.sheet_names.contains(&name.to_string()) {
            return Ok(());
        }
        self.sheet_names.push(name.to_string());
        Ok(())
    }

    /// Rename a sheet, updating sheet_names and re-keying all stored data.
    pub fn rename_sheet(&mut self, old_name: &str, new_name: &str) -> PyResult<()> {
        let idx = self
            .sheet_names
            .iter()
            .position(|n| n == old_name)
            .ok_or_else(|| {
                PyErr::new::<PyValueError, _>(format!("Unknown sheet: {old_name}"))
            })?;
        if self.sheet_names.contains(&new_name.to_string()) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Sheet '{new_name}' already exists"
            )));
        }
        self.sheet_names[idx] = new_name.to_string();

        // Re-key values (IndexMap — preserves insertion order).
        let old_keys: Vec<CellKey> = self
            .values
            .keys()
            .filter(|(s, _, _)| s == old_name)
            .cloned()
            .collect();
        for key in old_keys {
            if let Some(payload) = self.values.swap_remove(&key) {
                let new_key = (new_name.to_string(), key.1, key.2);
                self.values.insert(new_key, payload);
            }
        }

        // Re-key formats (HashMap).
        let fmt_keys: Vec<CellKey> = self
            .formats
            .keys()
            .filter(|(s, _, _)| s == old_name)
            .cloned()
            .collect();
        for key in fmt_keys {
            if let Some(fields) = self.formats.remove(&key) {
                let new_key = (new_name.to_string(), key.1, key.2);
                self.formats.insert(new_key, fields);
            }
        }

        // Re-key borders (HashMap).
        let bdr_keys: Vec<CellKey> = self
            .borders
            .keys()
            .filter(|(s, _, _)| s == old_name)
            .cloned()
            .collect();
        for key in bdr_keys {
            if let Some(fields) = self.borders.remove(&key) {
                let new_key = (new_name.to_string(), key.1, key.2);
                self.borders.insert(new_key, fields);
            }
        }

        // Re-key row heights.
        let rh_keys: Vec<(String, u32)> = self
            .row_heights
            .keys()
            .filter(|(s, _)| s == old_name)
            .cloned()
            .collect();
        for key in rh_keys {
            if let Some(h) = self.row_heights.remove(&key) {
                self.row_heights.insert((new_name.to_string(), key.1), h);
            }
        }

        // Re-key column widths.
        let cw_keys: Vec<(String, u16)> = self
            .col_widths
            .keys()
            .filter(|(s, _)| s == old_name)
            .cloned()
            .collect();
        for key in cw_keys {
            if let Some(w) = self.col_widths.remove(&key) {
                self.col_widths.insert((new_name.to_string(), key.1), w);
            }
        }

        // Re-key merge ranges.
        for mr in &mut self.merge_ranges {
            if mr.sheet == old_name {
                mr.sheet = new_name.to_string();
            }
        }

        // Re-key hyperlinks, comments, conditional formats, data validations, tables.
        for h in &mut self.hyperlinks {
            if h.sheet == old_name {
                h.sheet = new_name.to_string();
            }
        }
        for c in &mut self.comments {
            if c.sheet == old_name {
                c.sheet = new_name.to_string();
            }
        }
        if let Some(pane) = self.panes.remove(old_name) {
            self.panes.insert(new_name.to_string(), pane);
        }
        if let Some(pa) = self.print_areas.remove(old_name) {
            self.print_areas.insert(new_name.to_string(), pa);
        }
        for cf in &mut self.conditional_formats {
            if cf.sheet == old_name {
                cf.sheet = new_name.to_string();
            }
        }
        for dv in &mut self.data_validations {
            if dv.sheet == old_name {
                dv.sheet = new_name.to_string();
            }
        }
        for t in &mut self.tables {
            if t.sheet == old_name {
                t.sheet = new_name.to_string();
            }
        }

        Ok(())
    }

    pub fn write_cell_value(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;

        let key = resolve_key(sheet, a1)?;

        let dict = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let type_str: String = dict
            .get_item("type")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("payload missing 'type'"))?
            .extract()?;

        // Store the value for deferred writing.
        let value_str: Option<String> = dict.get_item("value")?.and_then(|v| {
            v.extract::<String>().ok().or_else(|| {
                // Handle numeric/bool values by converting to string.
                v.extract::<f64>()
                    .map(|n| n.to_string())
                    .ok()
                    .or_else(|| v.extract::<bool>().map(|b| b.to_string()).ok())
            })
        });
        let formula_str: Option<String> = dict.get_item("formula")?.and_then(|v| v.extract().ok());

        self.values.insert(
            key,
            CellPayload {
                type_str,
                value: value_str,
                formula: formula_str,
            },
        );

        Ok(())
    }

    /// Bulk-write a rectangular grid of values starting at `start_a1`.
    ///
    /// `values` is a 2-D Python list of raw values (int/float/str/None).
    /// None values are skipped (no cell written).  Used by performance
    /// workloads to avoid per-cell FFI overhead.
    pub fn write_sheet_values(
        &mut self,
        sheet: &str,
        start_a1: &str,
        values: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;

        let (base_row, base_col_32) =
            a1_to_row_col(start_a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let base_col: u16 = base_col_32.try_into().map_err(|_| {
            PyErr::new::<PyValueError, _>(format!("Column out of range: {start_a1}"))
        })?;

        let rows: Vec<Bound<'_, PyAny>> = values.extract()?;
        for (ri, row_obj) in rows.iter().enumerate() {
            let cols: Vec<Bound<'_, PyAny>> = row_obj.extract()?;
            for (ci, val) in cols.iter().enumerate() {
                if val.is_none() {
                    continue;
                }
                let row = base_row + ri as u32;
                let col = base_col + ci as u16;
                let key = (sheet.to_string(), row, col);

                // Infer type from Python object.
                if let Ok(f) = val.extract::<f64>() {
                    self.values.insert(
                        key,
                        CellPayload {
                            type_str: "number".to_string(),
                            value: Some(f.to_string()),
                            formula: None,
                        },
                    );
                } else if let Ok(i) = val.extract::<i64>() {
                    self.values.insert(
                        key,
                        CellPayload {
                            type_str: "number".to_string(),
                            value: Some((i as f64).to_string()),
                            formula: None,
                        },
                    );
                } else if let Ok(s) = val.extract::<String>() {
                    self.values.insert(
                        key,
                        CellPayload {
                            type_str: "string".to_string(),
                            value: Some(s),
                            formula: None,
                        },
                    );
                } else if let Ok(b) = val.extract::<bool>() {
                    self.values.insert(
                        key,
                        CellPayload {
                            type_str: "boolean".to_string(),
                            value: Some(b.to_string()),
                            formula: None,
                        },
                    );
                }
                // else: skip unsupported types silently.
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
        self.ensure_sheet_exists(sheet)?;
        let key = resolve_key(sheet, a1)?;
        let dict = format_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("format_dict must be a dict"))?;
        let fields = extract_format_fields(dict)?;
        self.formats.insert(key, fields);
        Ok(())
    }

    pub fn write_cell_border(
        &mut self,
        sheet: &str,
        a1: &str,
        border_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;
        let key = resolve_key(sheet, a1)?;
        let dict = border_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("border_dict must be a dict"))?;
        let fields = extract_border_fields(dict)?;
        self.borders.insert(key, fields);
        Ok(())
    }

    /// Bulk-write a rectangular grid of format dicts starting at `start_a1`.
    ///
    /// `formats` is a 2-D Python list where each element is either a dict
    /// (same schema as `write_cell_format`) or `None` (skip).  Used to
    /// eliminate per-cell FFI overhead for styled grids.
    pub fn write_sheet_formats(
        &mut self,
        sheet: &str,
        start_a1: &str,
        formats: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;

        let (base_row, base_col_32) =
            a1_to_row_col(start_a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let base_col: u16 = base_col_32.try_into().map_err(|_| {
            PyErr::new::<PyValueError, _>(format!("Column out of range: {start_a1}"))
        })?;

        let rows: Vec<Bound<'_, PyAny>> = formats.extract()?;
        for (ri, row_obj) in rows.iter().enumerate() {
            let cols: Vec<Bound<'_, PyAny>> = row_obj.extract()?;
            for (ci, val) in cols.iter().enumerate() {
                if val.is_none() {
                    continue;
                }
                let dict = val
                    .downcast::<PyDict>()
                    .map_err(|_| PyErr::new::<PyValueError, _>("format element must be dict or None"))?;
                if dict.is_empty() {
                    continue;
                }
                let row = base_row + ri as u32;
                let col = base_col + ci as u16;
                let key = (sheet.to_string(), row, col);
                let fields = extract_format_fields(dict)?;
                self.formats.insert(key, fields);
            }
        }

        Ok(())
    }

    /// Bulk-write a rectangular grid of border dicts starting at `start_a1`.
    ///
    /// `borders` is a 2-D Python list where each element is either a dict
    /// (same schema as `write_cell_border`) or `None` (skip).
    pub fn write_sheet_borders(
        &mut self,
        sheet: &str,
        start_a1: &str,
        borders: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;

        let (base_row, base_col_32) =
            a1_to_row_col(start_a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let base_col: u16 = base_col_32.try_into().map_err(|_| {
            PyErr::new::<PyValueError, _>(format!("Column out of range: {start_a1}"))
        })?;

        let rows: Vec<Bound<'_, PyAny>> = borders.extract()?;
        for (ri, row_obj) in rows.iter().enumerate() {
            let cols: Vec<Bound<'_, PyAny>> = row_obj.extract()?;
            for (ci, val) in cols.iter().enumerate() {
                if val.is_none() {
                    continue;
                }
                let dict = val
                    .downcast::<PyDict>()
                    .map_err(|_| PyErr::new::<PyValueError, _>("border element must be dict or None"))?;
                if dict.is_empty() {
                    continue;
                }
                let row = base_row + ri as u32;
                let col = base_col + ci as u16;
                let key = (sheet.to_string(), row, col);
                let fields = extract_border_fields(dict)?;
                self.borders.insert(key, fields);
            }
        }

        Ok(())
    }

    pub fn set_row_height(&mut self, sheet: &str, row: u32, height: f64) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;
        self.row_heights.insert((sheet.to_string(), row), height);
        Ok(())
    }

    pub fn set_column_width(&mut self, sheet: &str, col_str: &str, width: f64) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;
        let col_idx = col_letter_to_index(col_str)?;
        self.col_widths.insert((sheet.to_string(), col_idx), width);
        Ok(())
    }

    // =========================================================================
    // Tier 2 Write Operations
    // =========================================================================

    pub fn merge_cells(&mut self, sheet: &str, range_str: &str) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;
        let (r1, c1, r2, c2) = parse_a1_range(range_str)?;
        self.merge_ranges.push(MergeRange {
            sheet: sheet.to_string(),
            first_row: r1,
            first_col: c1,
            last_row: r2,
            last_col: c2,
        });
        Ok(())
    }

    pub fn add_hyperlink(&mut self, sheet: &str, link_dict: &Bound<'_, PyAny>) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;

        let dict = link_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("link must be a dict"))?;

        // Support optional wrapper key "hyperlink"
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("hyperlink")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let cell: Option<String> = cfg
            .get_item("cell")?
            .and_then(|v| v.extract::<String>().ok())
            .and_then(|s| if s.is_empty() { None } else { Some(s) });
        let target: Option<String> = cfg
            .get_item("target")?
            .and_then(|v| v.extract::<String>().ok())
            .and_then(|s| if s.is_empty() { None } else { Some(s) });

        // No-op for missing/empty payloads to match other adapters.
        let (Some(cell), Some(target)) = (cell, target) else {
            return Ok(());
        };
        let display: Option<String> = cfg
            .get_item("display")?
            .and_then(|v| v.extract::<String>().ok())
            .and_then(|s| if s.is_empty() { None } else { Some(s) });
        let tooltip: Option<String> = cfg
            .get_item("tooltip")?
            .and_then(|v| v.extract::<String>().ok())
            .and_then(|s| if s.is_empty() { None } else { Some(s) });
        let internal: bool = cfg
            .get_item("internal")?
            .and_then(|v| v.extract::<bool>().ok())
            .unwrap_or(false);

        let (row, col0) = a1_to_row_col(&cell).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let col: u16 = col0.try_into().map_err(|_| {
            PyErr::new::<PyValueError, _>(format!("Column out of range for Excel: {cell}"))
        })?;

        let url = if internal {
            format!("internal:{}", target.trim_start_matches('#'))
        } else {
            target
        };

        self.hyperlinks.push(HyperlinkPayload {
            sheet: sheet.to_string(),
            row,
            col,
            url,
            display,
            tooltip,
        });
        Ok(())
    }

    pub fn add_comment(&mut self, sheet: &str, comment_dict: &Bound<'_, PyAny>) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;

        let dict = comment_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("comment must be a dict"))?;

        // Support optional wrapper key "comment"
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("comment")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let cell: Option<String> = cfg
            .get_item("cell")?
            .and_then(|v| v.extract::<String>().ok())
            .and_then(|s| if s.is_empty() { None } else { Some(s) });
        let text: Option<String> = cfg
            .get_item("text")?
            .and_then(|v| v.extract::<String>().ok());

        // No-op for missing/empty payloads to match other adapters.
        let (Some(cell), Some(text)) = (cell, text) else {
            return Ok(());
        };
        let author: Option<String> = cfg
            .get_item("author")?
            .and_then(|v| v.extract::<String>().ok())
            .and_then(|s| if s.is_empty() { None } else { Some(s) });

        let (row, col0) = a1_to_row_col(&cell).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let col: u16 = col0.try_into().map_err(|_| {
            PyErr::new::<PyValueError, _>(format!("Column out of range for Excel: {cell}"))
        })?;

        self.comments.push(CommentPayload {
            sheet: sheet.to_string(),
            row,
            col,
            text,
            author,
        });
        Ok(())
    }

    pub fn set_freeze_panes(&mut self, sheet: &str, settings: &Bound<'_, PyAny>) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;

        let dict = settings
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("settings must be a dict"))?;

        // Support optional wrapper key "freeze"
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("freeze")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let mode: String = cfg
            .get_item("mode")?
            .and_then(|v| v.extract::<String>().ok())
            .unwrap_or_else(|| "freeze".to_string());

        if mode == "freeze" {
            let top_left: Option<String> = cfg
                .get_item("top_left_cell")?
                .and_then(|v| v.extract::<String>().ok());
            if let Some(cell) = top_left {
                let (row, col0) =
                    a1_to_row_col(&cell).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
                let col: u16 = col0.try_into().map_err(|_| {
                    PyErr::new::<PyValueError, _>(format!("Column out of range for Excel: {cell}"))
                })?;
                self.panes
                    .insert(sheet.to_string(), PaneSetting::Freeze { row, col });
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
            self.panes
                .insert(sheet.to_string(), PaneSetting::Split { x_split, y_split });
        }

        Ok(())
    }

    pub fn set_print_area(&mut self, sheet: &str, range_str: &str) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;
        self.print_areas
            .insert(sheet.to_string(), range_str.to_string());
        Ok(())
    }

    // =========================================================================
    // Tier 2/3 Write Operations (Sprint2)
    // =========================================================================

    pub fn add_conditional_format(
        &mut self,
        sheet: &str,
        rule_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;
        let dict = rule_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("rule must be a dict"))?;

        // Support optional wrapper key "cf_rule".
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("cf_rule")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let range: Option<String> = cfg.get_item("range")?.and_then(|v| v.extract().ok());
        let rule_type: Option<String> = cfg.get_item("rule_type")?.and_then(|v| v.extract().ok());
        let (Some(range), Some(rule_type)) = (range, rule_type) else {
            return Ok(());
        };

        let operator: Option<String> = cfg.get_item("operator")?.and_then(|v| v.extract().ok());
        let formula: Option<String> = cfg.get_item("formula")?.and_then(|v| v.extract().ok());
        let stop_if_true: bool = cfg
            .get_item("stop_if_true")?
            .and_then(|v| v.extract::<bool>().ok())
            .unwrap_or(false);

        let mut bg_color: Option<String> = None;
        if let Some(v) = cfg.get_item("format")? {
            if let Ok(fd) = v.downcast::<PyDict>() {
                bg_color = fd.get_item("bg_color")?.and_then(|x| x.extract().ok());
            }
        }

        self.conditional_formats.push(ConditionalFormatPayload {
            sheet: sheet.to_string(),
            range,
            rule_type,
            operator,
            formula,
            stop_if_true,
            bg_color,
        });
        Ok(())
    }

    pub fn add_data_validation(
        &mut self,
        sheet: &str,
        validation_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;
        let dict = validation_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("validation must be a dict"))?;

        // Support optional wrapper key "validation".
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("validation")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let range: Option<String> = cfg.get_item("range")?.and_then(|v| v.extract().ok());
        let validation_type: Option<String> = cfg
            .get_item("validation_type")?
            .and_then(|v| v.extract().ok());
        let (Some(range), Some(validation_type)) = (range, validation_type) else {
            return Ok(());
        };

        let operator: Option<String> = cfg.get_item("operator")?.and_then(|v| v.extract().ok());
        let formula1: Option<String> = cfg.get_item("formula1")?.and_then(|v| v.extract().ok());
        let formula2: Option<String> = cfg.get_item("formula2")?.and_then(|v| v.extract().ok());
        let allow_blank: Option<bool> = cfg
            .get_item("allow_blank")?
            .and_then(|v| v.extract::<bool>().ok());
        let error_title: Option<String> = cfg
            .get_item("error_title")?
            .and_then(|v| v.extract::<String>().ok())
            .and_then(|s| if s.is_empty() { None } else { Some(s) });
        let error: Option<String> = cfg
            .get_item("error")?
            .and_then(|v| v.extract::<String>().ok())
            .and_then(|s| if s.is_empty() { None } else { Some(s) });

        self.data_validations.push(DataValidationPayload {
            sheet: sheet.to_string(),
            range,
            validation_type,
            operator,
            formula1,
            formula2,
            allow_blank,
            error_title,
            error,
        });
        Ok(())
    }

    pub fn add_named_range(&mut self, sheet: &str, named_range: &Bound<'_, PyAny>) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;
        let dict = named_range
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("named_range must be a dict"))?;

        // Support optional wrapper key "named_range".
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("named_range")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let name: Option<String> = cfg.get_item("name")?.and_then(|v| v.extract().ok());
        let scope: String = cfg
            .get_item("scope")?
            .and_then(|v| v.extract::<String>().ok())
            .unwrap_or_else(|| "workbook".to_string());
        let refers_to: Option<String> = cfg.get_item("refers_to")?.and_then(|v| v.extract().ok());

        let (Some(name), Some(refers_to)) = (name, refers_to) else {
            return Ok(());
        };

        let sheet_scope = if scope == "sheet" {
            Some(sheet.to_string())
        } else {
            None
        };

        self.named_ranges.push(NamedRangePayload {
            name,
            scope,
            sheet: sheet_scope,
            refers_to,
        });
        Ok(())
    }

    pub fn add_table(&mut self, sheet: &str, table: &Bound<'_, PyAny>) -> PyResult<()> {
        self.ensure_sheet_exists(sheet)?;
        let dict = table
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("table must be a dict"))?;

        // Support optional wrapper key "table".
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("table")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let name: Option<String> = cfg.get_item("name")?.and_then(|v| v.extract().ok());
        let ref_range: Option<String> = cfg.get_item("ref")?.and_then(|v| v.extract().ok());
        let (Some(name), Some(ref_range)) = (name, ref_range) else {
            return Ok(());
        };

        let style: Option<String> = cfg
            .get_item("style")?
            .and_then(|v| v.extract::<String>().ok())
            .and_then(|s| if s.is_empty() { None } else { Some(s) });
        let totals_row: bool = cfg
            .get_item("totals_row")?
            .and_then(|v| v.extract::<bool>().ok())
            .unwrap_or(false);
        let header_row: bool = cfg
            .get_item("header_row")?
            .and_then(|v| v.extract::<bool>().ok())
            .unwrap_or(true);
        let autofilter: Option<bool> = cfg
            .get_item("autofilter")?
            .and_then(|v| v.extract::<bool>().ok());

        let mut cols: Vec<String> = Vec::new();
        if let Some(v) = cfg.get_item("columns")? {
            if let Ok(list) = v.extract::<Vec<String>>() {
                cols = list;
            }
        }

        self.tables.push(TablePayload {
            sheet: sheet.to_string(),
            name,
            ref_range,
            style,
            columns: cols,
            totals_row,
            autofilter,
            header_row,
        });
        Ok(())
    }

    pub fn save(&mut self, path: &str) -> PyResult<()> {
        if self.saved {
            return Err(PyErr::new::<PyValueError, _>(
                "Workbook already saved (RustXlsxWriterBook is consumed-on-save)",
            ));
        }
        self.saved = true;

        let mut wb = Workbook::new();

        // Create worksheets in insertion order.
        let mut ws_map: IndexMap<String, Worksheet> = IndexMap::new();
        for name in &self.sheet_names {
            let mut ws = Worksheet::new();
            ws.set_name(name)
                .map_err(|e| PyErr::new::<PyValueError, _>(format!("Invalid sheet name: {e}")))?;
            ws_map.insert(name.clone(), ws);
        }

        // Apply row heights.  Python passes 1-indexed rows (openpyxl convention).
        for ((sheet, row), height) in &self.row_heights {
            if let Some(ws) = ws_map.get_mut(sheet) {
                ws.set_row_height(row.saturating_sub(1), *height).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("set_row_height failed: {e}"))
                })?;
            }
        }

        // Apply column widths.
        for ((sheet, col), width) in &self.col_widths {
            if let Some(ws) = ws_map.get_mut(sheet) {
                ws.set_column_width(*col, *width).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("set_column_width failed: {e}"))
                })?;
            }
        }

        let mut split_patches: Vec<(String, i32, i32)> = Vec::new();
        let mut table_ref_patches: Vec<(String, String)> = Vec::new();

        // Freeze/split panes.
        for (sheet, setting) in &self.panes {
            if let Some(ws) = ws_map.get_mut(sheet) {
                match setting {
                    PaneSetting::Freeze { row, col } => {
                        ws.set_freeze_panes(*row, *col).map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!("set_freeze_panes failed: {e}"))
                        })?;
                    }
                    PaneSetting::Split { x_split, y_split } => {
                        // rust_xlsxwriter doesn't currently support split panes (non-freeze)
                        // via its public API. We write an equivalent freeze panes record to
                        // generate a <pane/> element, then patch the worksheet XML to convert
                        // it into a split panes view.
                        let x_raw = if x_split.is_finite() { *x_split } else { 0.0 };
                        let y_raw = if y_split.is_finite() { *y_split } else { 0.0 };
                        let x_i32: i32 = x_raw.round() as i32;
                        let y_i32: i32 = y_raw.round() as i32;

                        let col: u16 = if x_i32 <= 0 {
                            0
                        } else {
                            u16::try_from(x_i32).map_err(|_| {
                                PyErr::new::<PyValueError, _>(format!(
                                    "x_split out of range for Excel: {x_i32}"
                                ))
                            })?
                        };
                        let row: u32 = if y_i32 <= 0 {
                            0
                        } else {
                            u32::try_from(y_i32).map_err(|_| {
                                PyErr::new::<PyValueError, _>(format!(
                                    "y_split out of range for Excel: {y_i32}"
                                ))
                            })?
                        };

                        if row == 0 && col == 0 {
                            continue;
                        }

                        ws.set_freeze_panes(row, col).map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!("set_freeze_panes failed: {e}"))
                        })?;
                        split_patches.push((sheet.clone(), x_i32.max(0), y_i32.max(0)));
                    }
                }
            }
        }

        // Print areas.
        for (sheet, range_str) in &self.print_areas {
            if let Some(ws) = ws_map.get_mut(sheet) {
                // Parse "A1:D10" → (first_row, first_col, last_row, last_col)
                let clean = range_str.replace('$', "");
                let parts: Vec<&str> = clean.split(':').collect();
                if parts.len() == 2 {
                    let (r1, c1) =
                        a1_to_row_col(parts[0]).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
                    let (r2, c2) =
                        a1_to_row_col(parts[1]).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
                    let c1_u16: u16 = c1.try_into().map_err(|_| {
                        PyErr::new::<PyValueError, _>(format!("Column out of range: {range_str}"))
                    })?;
                    let c2_u16: u16 = c2.try_into().map_err(|_| {
                        PyErr::new::<PyValueError, _>(format!("Column out of range: {range_str}"))
                    })?;
                    ws.set_print_area(r1, c1_u16, r2, c2_u16).map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("set_print_area failed: {e}"))
                    })?;
                }
            }
        }

        // Merged ranges.
        let merge_format = Format::new();
        for m in &self.merge_ranges {
            if let Some(ws) = ws_map.get_mut(&m.sheet) {
                ws.merge_range(
                    m.first_row,
                    m.first_col,
                    m.last_row,
                    m.last_col,
                    "",
                    &merge_format,
                )
                .map(|_| ())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("merge_range failed: {e}")))?;
            }
        }

        // Write all cells with merged format+border.
        for (key, payload) in &self.values {
            let (ref sheet, row, col) = *key;
            let fmt_fields = self.formats.get(key);
            let bdr_fields = self.borders.get(key);
            let mut format = build_format(fmt_fields, bdr_fields)?;

            // Apply default date/datetime number format only if the user
            // didn't already provide one via write_cell_format.
            let has_user_nf = fmt_fields.and_then(|f| f.number_format.as_ref()).is_some();
            if !has_user_nf {
                if payload.type_str == "date" {
                    format = format.set_num_format("yyyy-mm-dd");
                } else if payload.type_str == "datetime" {
                    format = format.set_num_format("yyyy-mm-dd hh:mm:ss");
                }
            }

            let ws = ws_map
                .get_mut(sheet)
                .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

            write_cell(ws, row, col, payload, &format)?;
        }

        // Write formats for cells that have format/border but no value
        // (e.g., blank cells with borders).
        let format_only_keys: HashSet<_> = self
            .formats
            .keys()
            .chain(self.borders.keys())
            .filter(|k| !self.values.contains_key(*k))
            .collect();
        for key in format_only_keys {
            let (ref sheet, row, col) = *key;
            let fmt_fields = self.formats.get(key);
            let bdr_fields = self.borders.get(key);
            let format = build_format(fmt_fields, bdr_fields)?;
            if let Some(ws) = ws_map.get_mut(sheet) {
                ws.write_blank(row, col, &format)
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("write_blank failed: {e}")))?;
            }
        }

        // Hyperlinks (apply after cell writes so they win on the final cell record).
        for link in &self.hyperlinks {
            if let Some(ws) = ws_map.get_mut(&link.sheet) {
                let mut url_obj = Url::new(link.url.as_str());
                if let Some(tip) = &link.tooltip {
                    url_obj = url_obj.set_tip(tip.as_str());
                }
                if let Some(text) = &link.display {
                    ws.write_url_with_text(link.row, link.col, &url_obj, text)
                        .map(|_| ())
                        .map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!("write_url_with_text failed: {e}"))
                        })?;
                } else {
                    ws.write_url(link.row, link.col, &url_obj)
                        .map(|_| ())
                        .map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!("write_url failed: {e}"))
                        })?;
                }
            }
        }

        // Comments/notes.
        for comment in &self.comments {
            if let Some(ws) = ws_map.get_mut(&comment.sheet) {
                // rust_xlsxwriter prefixes the author name into the note text by default
                // (e.g. "Author:\n..."). ExcelBench expectations model the text body only,
                // so disable the author prefix.
                let mut note = Note::new(comment.text.as_str()).add_author_prefix(false);
                if let Some(author) = &comment.author {
                    note = note.set_author(author.as_str());
                }
                ws.insert_note(comment.row, comment.col, &note)
                    .map(|_| ())
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("insert_note failed: {e}")))?;
            }
        }

        // Workbook defined names (named ranges).
        for nr in &self.named_ranges {
            let mut name = nr.name.clone();
            if nr.scope == "sheet" {
                if let Some(sheet_name) = &nr.sheet {
                    let quoted = quote_sheet_name(sheet_name);
                    name = format!("{quoted}!{}", name);
                }
            }
            let mut refers = nr.refers_to.clone();
            if !refers.trim_start().starts_with('=') {
                refers = format!("={refers}");
            }
            wb.define_name(name, &refers)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("define_name failed: {e}")))?;
        }

        // Conditional formats.
        for cf in &self.conditional_formats {
            let (r1, c1, r2, c2) = parse_a1_range(&cf.range)?;
            let ws = match ws_map.get_mut(&cf.sheet) {
                Some(w) => w,
                None => continue,
            };
            let rule_type = cf.rule_type.as_str();

            if rule_type == "cellIs" {
                let Some(op) = &cf.operator else {
                    continue;
                };
                let Some(formula) = &cf.formula else {
                    continue;
                };
                let value_str = formula.trim_start_matches('=');
                let rule = map_cf_cell_rule(op, value_str)?;
                let mut fmt = Format::new();
                if let Some(bg) = &cf.bg_color {
                    let c = parse_hex_color(bg);
                    fmt = fmt
                        .set_foreground_color(c)
                        .set_background_color(c)
                        .set_pattern(FormatPattern::Solid);
                }
                let mut cfmt = ConditionalFormatCell::new().set_rule(rule).set_format(fmt);
                if cf.stop_if_true {
                    cfmt = cfmt.set_stop_if_true(true);
                }
                ws.add_conditional_format(r1, c1, r2, c2, &cfmt)
                    .map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("add_conditional_format failed: {e}"))
                    })?;
            } else if rule_type == "expression" {
                let Some(formula) = &cf.formula else {
                    continue;
                };
                let f = if formula.trim_start().starts_with('=') {
                    formula.clone()
                } else {
                    format!("={formula}")
                };
                let mut fmt = Format::new();
                if let Some(bg) = &cf.bg_color {
                    let c = parse_hex_color(bg);
                    fmt = fmt
                        .set_foreground_color(c)
                        .set_background_color(c)
                        .set_pattern(FormatPattern::Solid);
                }
                let mut cfmt = ConditionalFormatFormula::new()
                    .set_rule(f.as_str())
                    .set_format(fmt);
                if cf.stop_if_true {
                    cfmt = cfmt.set_stop_if_true(true);
                }
                ws.add_conditional_format(r1, c1, r2, c2, &cfmt)
                    .map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("add_conditional_format failed: {e}"))
                    })?;
            } else if rule_type == "dataBar" {
                let cfmt = ConditionalFormatDataBar::new();
                ws.add_conditional_format(r1, c1, r2, c2, &cfmt)
                    .map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("add_conditional_format failed: {e}"))
                    })?;
            } else if rule_type == "colorScale" {
                let cfmt = ConditionalFormat3ColorScale::new();
                ws.add_conditional_format(r1, c1, r2, c2, &cfmt)
                    .map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("add_conditional_format failed: {e}"))
                    })?;
            }
        }

        // Data validations.
        for dv in &self.data_validations {
            let (r1, c1, r2, c2) = parse_a1_range(&dv.range)?;
            let ws = match ws_map.get_mut(&dv.sheet) {
                Some(w) => w,
                None => continue,
            };

            let mut v = DataValidation::new();
            let vtype = dv.validation_type.to_ascii_lowercase();
            if vtype == "list" {
                if let Some(f1) = &dv.formula1 {
                    let f1t = f1.trim();
                    if f1t.starts_with('"') && f1t.ends_with('"') && f1t.len() >= 2 {
                        let inner = &f1t[1..f1t.len() - 1];
                        let parts: Vec<&str> = inner.split(',').collect();
                        v = v.allow_list_strings(&parts).map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!("allow_list_strings failed: {e}"))
                        })?;
                    } else {
                        // Use formula list source.
                        v = v.allow_list_formula(Formula::new(f1t));
                    }
                }
            } else if vtype == "custom" {
                if let Some(f1) = &dv.formula1 {
                    v = v.allow_custom(Formula::new(f1.as_str()));
                }
            } else if vtype == "whole" {
                let op = dv.operator.as_deref().unwrap_or("between");
                if op == "between" {
                    if let (Some(f1), Some(f2)) = (&dv.formula1, &dv.formula2) {
                        let rule = map_dv_rule_between_i32(
                            f1.trim_start_matches('='),
                            f2.trim_start_matches('='),
                        )?;
                        v = v.allow_whole_number(rule);
                    }
                }
            }

            if dv.allow_blank == Some(false) {
                v = v.ignore_blank(false);
            }
            if let Some(t) = &dv.error_title {
                v = v.set_error_title(t).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("set_error_title failed: {e}"))
                })?;
            }
            if let Some(msg) = &dv.error {
                v = v.set_error_message(msg).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("set_error_message failed: {e}"))
                })?;
            }

            ws.add_data_validation(r1, c1, r2, c2, &v).map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("add_data_validation failed: {e}"))
            })?;
        }

        // Tables.
        for tbl in &self.tables {
            let (r1, c1, r2, c2) = parse_a1_range(&tbl.ref_range)?;
            let ws = match ws_map.get_mut(&tbl.sheet) {
                Some(w) => w,
                None => continue,
            };

            let mut columns: Vec<TableColumn> = Vec::new();
            if !tbl.columns.is_empty() {
                for c in &tbl.columns {
                    columns.push(TableColumn::new().set_header(c));
                }
            }

            let mut table = Table::new().set_name(&tbl.name);
            if !columns.is_empty() {
                table = table.set_columns(&columns);
            }
            if tbl.totals_row {
                table = table.set_total_row(true);
            }
            table = table.set_header_row(tbl.header_row);
            if let Some(af) = tbl.autofilter {
                table = table.set_autofilter(af);
            }
            table = table.set_style(map_table_style(tbl.style.as_deref()));

            // Header-only table workaround: write a 2-row range, then patch the table XML ref.
            let last_row_for_add = if tbl.header_row && r1 == r2 {
                r2 + 1
            } else {
                r2
            };
            if tbl.header_row && r1 == r2 {
                table_ref_patches.push((tbl.name.clone(), tbl.ref_range.clone()));
            }

            ws.add_table(r1, c1, last_row_for_add, c2, &table)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("add_table failed: {e}")))?;
        }

        for (_name, ws) in ws_map.drain(..) {
            wb.push_worksheet(ws);
        }

        wb.save(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to save workbook: {e}")))?;

        // Post-process split panes (edge case) by patching OOXML.
        if !split_patches.is_empty() {
            if let Err(e) = patch_split_panes_xlsx(path, &split_patches) {
                eprintln!("Failed to patch split panes in {path}: {e}");
            }
        }

        // Post-process header-only tables by patching table XML refs.
        if !table_ref_patches.is_empty() {
            if let Err(e) = patch_tables_xlsx(path, &table_ref_patches) {
                eprintln!("Failed to patch table refs in {path}: {e}");
            }
        }

        Ok(())
    }
}
