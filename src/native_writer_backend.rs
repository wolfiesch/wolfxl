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

use wolfxl_writer::model::chart::{
    Axis, AxisCommon, AxisOrientation, AxisPos, BarDir, BarGrouping, CategoryAxis, Chart,
    ChartKind, DataLabels, DateAxis, DisplayBlanksAs, ErrorBarType, ErrorBarValType, ErrorBars,
    GraphicalProperties, Layout, LayoutTarget, Legend, LegendPosition, Marker, MarkerSymbol,
    RadarStyle, Reference as ChartReference, ScatterStyle, Series, SeriesAxis, SeriesTitle,
    TickMark, Title as ChartTitle, TitleRun, Trendline, TrendlineKind, ValueAxis,
};
use wolfxl_writer::model::date::{date_to_excel_serial, datetime_to_excel_serial};
use wolfxl_writer::model::image::{ImageAnchor, SheetImage};
use wolfxl_writer::model::{
    AlignmentSpec, BorderSideSpec, BorderSpec, CellIsOperator, Comment, CommentAuthorTable,
    ConditionalFormat, ConditionalKind, ConditionalRule, DataValidation, DefinedName, DocProperties,
    DxfRecord, ErrorStyle, FillSpec, FontSpec, FormatSpec, Hyperlink, StylesBuilder, Table,
    TableColumn, TableStyle, ValidationType, ValidationOperator, Worksheet, WriteCellValue,
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
/// `value` / `formula` per the historical oracle mapping (preserved
/// verbatim from the W5-removed `rust_xlsxwriter_backend`'s `write_cell`)
/// so any pre-W5 Python flush payload still routes correctly.
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
            // W4E.R5: reject non-finite values (NaN, +/-Infinity) at the
            // boundary. Strings like "NaN" / "inf" parse to f64 successfully
            // but have no OOXML representation.
            Ok(WriteCellValue::Number(require_finite_f64(n, "number cell")?))
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

/// Reject non-finite floats (NaN, +/-Infinity) at the pyclass boundary so
/// the emitter never has to format them. OOXML has no representation for
/// these values; without the guard, `format_number` in the writer would
/// emit literal `"NaN"`/`"inf"` and Excel/LO would reject the file.
/// Returns the value unchanged if finite, or a `PyValueError` otherwise.
/// W4E.R5.
fn require_finite_f64(f: f64, context: &str) -> PyResult<f64> {
    if !f.is_finite() {
        return Err(PyValueError::new_err(format!(
            "{context}: non-finite floats (NaN, Infinity) are not representable in xlsx; got {f}",
        )));
    }
    Ok(f)
}

/// Coerce a raw Python value (from `write_sheet_values`'s 2-D list) to a
/// `WriteCellValue`. The order-of-attempts mirrors the historical
/// `rust_xlsxwriter_backend` (removed in W5) but fixes a subtle bug: the
/// legacy oracle tried `f64` before `bool`, so `True`/`False` (which
/// extract as `1.0`/`0.0`) silently became numbers. The Python flush path
/// avoids this by routing booleans through `write_cell_value` instead,
/// but we tighten the rule here for correctness — bool first.
///
/// W4E.R5: returns `PyResult<Option<…>>` so non-finite floats can raise
/// rather than silently emitting invalid `"NaN"`/`"inf"` text in the
/// output XML. `Ok(None)` still means "no usable coercion, skip" (oracle
/// parity); `Err(…)` means the value was a finite-violating float.
fn raw_python_to_write_cell_value(
    value: &Bound<'_, PyAny>,
) -> PyResult<Option<WriteCellValue>> {
    if value.is_none() {
        return Ok(None);
    }
    // Boolean check via `is_instance_of` (rather than `extract`) since
    // `extract::<bool>()` would succeed on `0`/`1` ints too.
    let py = value.py();
    let bool_type = py.get_type::<pyo3::types::PyBool>();
    if value.is_instance(&bool_type).unwrap_or(false) {
        let b = value.extract::<bool>()?;
        return Ok(Some(WriteCellValue::Boolean(b)));
    }
    if let Ok(i) = value.extract::<i64>() {
        return Ok(Some(WriteCellValue::Number(i as f64)));
    }
    if let Ok(f) = value.extract::<f64>() {
        return Ok(Some(WriteCellValue::Number(require_finite_f64(f, "cell value")?)));
    }
    if let Ok(s) = value.extract::<String>() {
        if s.starts_with('=') {
            return Ok(Some(WriteCellValue::Formula {
                expr: s.trim_start_matches('=').to_string(),
                result: None,
            }));
        }
        return Ok(Some(WriteCellValue::String(s)));
    }
    // Datetime / date — best-effort via isoformat() if exposed.
    if let Ok(iso) = value.call_method0("isoformat") {
        if let Ok(s) = iso.extract::<String>() {
            if let Some(dt) = parse_iso_datetime(&s) {
                if let Some(serial) = datetime_to_excel_serial(dt) {
                    return Ok(Some(WriteCellValue::DateSerial(serial)));
                }
            }
            if let Some(d) = parse_iso_date(&s) {
                if let Some(serial) = date_to_excel_serial(d) {
                    return Ok(Some(WriteCellValue::DateSerial(serial)));
                }
            }
        }
    }
    Ok(None)
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
            // FontSpec stores u32 points. Clamp non-negative; OOXML/Excel
            // round to whole points even though `<sz val="..."/>` accepts a
            // floating-point literal.
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

/// Sprint Ι Pod-α: same conversion as ``py_runs_to_rust`` in the
/// patcher module, but typed against the writer's own re-exported
/// types.  Kept separate to avoid introducing a cross-module use of
/// the patcher helpers (the native writer doesn't depend on the
/// patcher).
fn py_runs_to_rust_writer(
    runs: &Bound<'_, pyo3::types::PyList>,
) -> PyResult<Vec<wolfxl_writer::rich_text::RichTextRun>> {
    use wolfxl_writer::rich_text::{InlineFontProps, RichTextRun};
    let mut out: Vec<RichTextRun> = Vec::with_capacity(runs.len());
    for entry in runs.iter() {
        let seq: &Bound<'_, pyo3::types::PySequence> = entry.downcast()?;
        if seq.len()? < 2 {
            return Err(PyValueError::new_err(
                "rich-text run must be a (text, font_or_none) pair",
            ));
        }
        let text: String = seq.get_item(0)?.extract()?;
        let font_obj = seq.get_item(1)?;
        let font = if font_obj.is_none() {
            None
        } else {
            let d: &Bound<'_, PyDict> = font_obj.downcast()?;
            let mut props = InlineFontProps::default();
            macro_rules! pull_bool {
                ($k:literal, $field:ident) => {
                    if let Some(v) = d.get_item($k)? {
                        if !v.is_none() {
                            props.$field = Some(v.extract::<bool>()?);
                        }
                    }
                };
            }
            macro_rules! pull_str {
                ($k:literal, $field:ident) => {
                    if let Some(v) = d.get_item($k)? {
                        if !v.is_none() {
                            let s: String = v.extract()?;
                            props.$field = Some(s);
                        }
                    }
                };
            }
            macro_rules! pull_i32 {
                ($k:literal, $field:ident) => {
                    if let Some(v) = d.get_item($k)? {
                        if !v.is_none() {
                            props.$field = Some(v.extract::<i32>()?);
                        }
                    }
                };
            }
            pull_bool!("b", bold);
            pull_bool!("i", italic);
            pull_bool!("strike", strike);
            pull_str!("u", underline);
            if let Some(v) = d.get_item("sz")? {
                if !v.is_none() {
                    props.size = Some(v.extract::<f64>()?);
                }
            }
            pull_str!("color", color);
            pull_str!("rFont", name);
            pull_i32!("family", family);
            pull_i32!("charset", charset);
            pull_str!("vertAlign", vert_align);
            pull_str!("scheme", scheme);
            Some(props)
        };
        out.push(RichTextRun { text, font });
    }
    Ok(out)
}

fn require_sheet<'wb>(wb: &'wb mut Workbook, name: &str) -> PyResult<&'wb mut Worksheet> {
    wb.sheet_mut_by_name(name)
        .ok_or_else(|| PyValueError::new_err(format!("Unknown sheet: {name}")))
}

// ---------------------------------------------------------------------------
// Wave 4B conversion helpers
// ---------------------------------------------------------------------------

/// Unwrap an optional wrapper key: if `dict` has `wrapper_key` as a key whose
/// value is a dict, return that inner dict. Otherwise return the original dict
/// unchanged. Mirrors the oracle's "inner dict may be passed bare or wrapped"
/// idiom used in all 8 rich-feature methods.
fn unwrap_optional_wrapper<'py>(
    dict: &'py Bound<'py, PyDict>,
    wrapper_key: &str,
) -> PyResult<Bound<'py, PyDict>> {
    if let Some(v) = dict.get_item(wrapper_key)? {
        if let Ok(inner) = v.downcast::<PyDict>() {
            return Ok(inner.clone());
        }
    }
    Ok(dict.clone())
}

/// Build a `(a1_ref, Hyperlink)` pair from a cfg dict, or `None` for silent no-op.
fn dict_to_hyperlink(cfg: &Bound<'_, PyDict>) -> PyResult<Option<(String, Hyperlink)>> {
    let cell: Option<String> = cfg
        .get_item("cell")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let target: Option<String> = cfg
        .get_item("target")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });

    let (Some(cell), Some(raw_target)) = (cell, target) else {
        return Ok(None); // silent no-op — match oracle
    };

    let display: Option<String> = cfg
        .get_item("display")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let tooltip: Option<String> = cfg
        .get_item("tooltip")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let is_internal: bool = cfg
        .get_item("internal")?
        .and_then(|v| v.extract::<bool>().ok())
        .unwrap_or(false);

    // The model's ``is_internal`` flag is the source of truth — see the
    // doc comment on ``Hyperlink``. ``target`` always stores the bare
    // form: a URL for external, a ``Sheet2!A1`` location for internal.
    // We strip a stray leading ``#`` for backward compat with callers
    // that wrote it both ways under the old prefix-sniffing convention.
    let target = if is_internal {
        raw_target.trim_start_matches('#').to_string()
    } else {
        raw_target
    };

    Ok(Some((
        cell,
        Hyperlink {
            target,
            is_internal,
            display,
            tooltip,
        },
    )))
}

/// Build a `(a1_ref, Comment)` pair from a cfg dict, or `None` for silent no-op.
/// Interns the author into `authors` before returning.
fn dict_to_comment(
    cfg: &Bound<'_, PyDict>,
    authors: &mut CommentAuthorTable,
) -> PyResult<Option<(String, Comment)>> {
    let cell: Option<String> = cfg
        .get_item("cell")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let text: Option<String> = cfg
        .get_item("text")?
        .and_then(|v| v.extract::<String>().ok());

    let (Some(cell), Some(text)) = (cell, text) else {
        return Ok(None); // silent no-op — match oracle
    };

    let author_name: String = cfg
        .get_item("author")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) })
        .unwrap_or_default();

    let author_id = authors.intern(author_name);

    Ok(Some((
        cell,
        Comment {
            text,
            author_id,
            width_pt: None,
            height_pt: None,
            visible: false,
        },
    )))
}

/// Build a `ConditionalFormat` from a cfg dict, or `None` for silent no-op.
/// May intern a dxf into `styles` when `format.bg_color` is set.
fn dict_to_conditional_format(
    cfg: &Bound<'_, PyDict>,
    styles: &mut StylesBuilder,
) -> PyResult<Option<ConditionalFormat>> {
    let range: Option<String> = cfg.get_item("range")?.and_then(|v| v.extract().ok());
    let rule_type: Option<String> = cfg.get_item("rule_type")?.and_then(|v| v.extract().ok());

    let (Some(range), Some(rule_type)) = (range, rule_type) else {
        return Ok(None); // silent no-op — match oracle
    };

    let operator: Option<String> = cfg.get_item("operator")?.and_then(|v| v.extract().ok());
    let formula: Option<String> = cfg.get_item("formula")?.and_then(|v| v.extract().ok());
    let stop_if_true: bool = cfg
        .get_item("stop_if_true")?
        .and_then(|v| v.extract::<bool>().ok())
        .unwrap_or(false);

    // Optional bg_color → intern a DxfRecord.
    let mut bg_color: Option<String> = None;
    if let Some(v) = cfg.get_item("format")? {
        if let Ok(fd) = v.downcast::<PyDict>() {
            bg_color = fd.get_item("bg_color")?.and_then(|x| x.extract().ok());
        }
    }
    let dxf_id: Option<u32> = if let Some(ref hex) = bg_color {
        parse_hex_color(hex).map(|rgb| {
            let dxf = DxfRecord {
                font: None,
                fill: Some(FillSpec {
                    pattern_type: "solid".to_string(),
                    fg_color_rgb: Some(rgb.clone()),
                    bg_color_rgb: Some(rgb),
                }),
                border: None,
            };
            styles.intern_dxf(&dxf)
        })
    } else {
        None
    };

    // Map rule_type + operator → ConditionalKind.
    let kind = match rule_type.as_str() {
        "cellIs" | "cell_is" => {
            let op_str = operator.as_deref().unwrap_or("equal");
            let op = match op_str {
                "equal" | "==" => CellIsOperator::Equal,
                "notEqual" | "!=" => CellIsOperator::NotEqual,
                "greaterThan" | ">" => CellIsOperator::GreaterThan,
                "greaterThanOrEqual" | ">=" => CellIsOperator::GreaterThanOrEqual,
                "lessThan" | "<" => CellIsOperator::LessThan,
                "lessThanOrEqual" | "<=" => CellIsOperator::LessThanOrEqual,
                "between" => CellIsOperator::Between,
                "notBetween" => CellIsOperator::NotBetween,
                _ => CellIsOperator::Equal,
            };

            let fstr = formula.as_deref().unwrap_or("").trim_start_matches('=');
            let (formula_a, formula_b) =
                if matches!(op, CellIsOperator::Between | CellIsOperator::NotBetween) {
                    // "formula1,formula2" convention — split on the first comma.
                    if let Some(idx) = fstr.find(',') {
                        (fstr[..idx].trim().to_string(), Some(fstr[idx + 1..].trim().to_string()))
                    } else {
                        (fstr.to_string(), None)
                    }
                } else {
                    (fstr.to_string(), None)
                };

            ConditionalKind::CellIs {
                operator: op,
                formula_a,
                formula_b,
            }
        }
        "expression" | "formula" => {
            let fstr = formula
                .as_deref()
                .unwrap_or("")
                .trim_start_matches('=')
                .to_string();
            ConditionalKind::Expression { formula: fstr }
        }
        // TODO: future wave — color_scale, data_bar, icon_set, duplicates,
        // unique, top, bottom, above_average, below_average, text_contains variants
        _ => ConditionalKind::Expression {
            formula: "FALSE()".to_string(),
        },
    };

    let rule = ConditionalRule {
        kind,
        dxf_id,
        stop_if_true,
    };

    Ok(Some(ConditionalFormat {
        sqref: range,
        rules: vec![rule],
    }))
}

/// Build a `DataValidation` from a cfg dict, or `None` for silent no-op.
fn dict_to_data_validation(cfg: &Bound<'_, PyDict>) -> PyResult<Option<DataValidation>> {
    let range: Option<String> = cfg.get_item("range")?.and_then(|v| v.extract().ok());
    let validation_type: Option<String> = cfg
        .get_item("validation_type")?
        .and_then(|v| v.extract().ok());

    let (Some(range), Some(vtype_str)) = (range, validation_type) else {
        return Ok(None); // silent no-op — match oracle
    };

    let validation_type = match vtype_str.as_str() {
        "whole" | "Whole" => ValidationType::Whole,
        "decimal" | "Decimal" => ValidationType::Decimal,
        "list" | "List" => ValidationType::List,
        "date" | "Date" => ValidationType::Date,
        "time" | "Time" => ValidationType::Time,
        "textLength" | "TextLength" | "text_length" => ValidationType::TextLength,
        "custom" | "Custom" => ValidationType::Custom,
        _ => ValidationType::Any,
    };

    let operator: Option<String> = cfg.get_item("operator")?.and_then(|v| v.extract().ok());
    let operator = match operator.as_deref().unwrap_or("between") {
        "between" | "Between" => ValidationOperator::Between,
        "notBetween" | "NotBetween" | "not_between" => ValidationOperator::NotBetween,
        "equal" | "Equal" | "==" => ValidationOperator::Equal,
        "notEqual" | "NotEqual" | "not_equal" | "!=" => ValidationOperator::NotEqual,
        "greaterThan" | "GreaterThan" | "greater_than" | ">" => ValidationOperator::GreaterThan,
        "lessThan" | "LessThan" | "less_than" | "<" => ValidationOperator::LessThan,
        "greaterThanOrEqual" | "GreaterThanOrEqual" | ">=" => {
            ValidationOperator::GreaterThanOrEqual
        }
        "lessThanOrEqual" | "LessThanOrEqual" | "<=" => ValidationOperator::LessThanOrEqual,
        _ => ValidationOperator::Between,
    };

    let formula_a: Option<String> = cfg.get_item("formula1")?.and_then(|v| v.extract().ok());
    let formula_b: Option<String> = cfg.get_item("formula2")?.and_then(|v| v.extract().ok());
    let allow_blank: bool = cfg
        .get_item("allow_blank")?
        .and_then(|v| v.extract::<bool>().ok())
        .unwrap_or(true); // oracle uses .unwrap_or(true)

    let error_title: Option<String> = cfg
        .get_item("error_title")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let error_message: Option<String> = cfg
        .get_item("error")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });

    Ok(Some(DataValidation {
        sqref: range,
        validation_type,
        operator,
        formula_a,
        formula_b,
        allow_blank,
        show_dropdown: true,
        show_error_message: true,
        error_style: ErrorStyle::Stop,
        error_title,
        error_message,
        show_input_message: false,
        input_title: None,
        input_message: None,
    }))
}

/// Build a `Table` from a cfg dict, or `None` for silent no-op.
fn dict_to_table(cfg: &Bound<'_, PyDict>) -> PyResult<Option<Table>> {
    let name: Option<String> = cfg.get_item("name")?.and_then(|v| v.extract().ok());
    let ref_range: Option<String> = cfg.get_item("ref")?.and_then(|v| v.extract().ok());

    let (Some(name), Some(ref_range)) = (name, ref_range) else {
        return Ok(None); // silent no-op — match oracle
    };

    let style: Option<String> = cfg
        .get_item("style")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let header_row: bool = cfg
        .get_item("header_row")?
        .and_then(|v| v.extract::<bool>().ok())
        .unwrap_or(true);
    let totals_row: bool = cfg
        .get_item("totals_row")?
        .and_then(|v| v.extract::<bool>().ok())
        .unwrap_or(false);
    let autofilter: bool = cfg
        .get_item("autofilter")?
        .and_then(|v| v.extract::<bool>().ok())
        .unwrap_or(true);

    let mut columns: Vec<TableColumn> = Vec::new();
    if let Some(v) = cfg.get_item("columns")? {
        if let Ok(list) = v.extract::<Vec<String>>() {
            for col_name in list {
                columns.push(TableColumn {
                    name: col_name,
                    totals_function: None,
                    totals_label: None,
                });
            }
        }
    }

    let table_style: Option<TableStyle> = style.map(|s| TableStyle {
        name: s,
        show_first_column: false,
        show_last_column: false,
        show_row_stripes: true,
        show_column_stripes: false,
    });

    Ok(Some(Table {
        name,
        display_name: None,
        range: ref_range,
        columns,
        header_row,
        totals_row,
        style: table_style,
        autofilter,
    }))
}

/// Build a `DefinedName` from a cfg dict. Returns `None` for silent no-op when
/// required fields are missing. Returns `Err` when scope="sheet" but the sheet
/// doesn't exist (that's a bug, not user input error).
fn dict_to_defined_name(
    wb: &Workbook,
    sheet_name: &str,
    cfg: &Bound<'_, PyDict>,
) -> PyResult<Option<DefinedName>> {
    let name: Option<String> = cfg.get_item("name")?.and_then(|v| v.extract().ok());
    let refers_to: Option<String> = cfg.get_item("refers_to")?.and_then(|v| v.extract().ok());

    let (Some(name), Some(refers_to)) = (name, refers_to) else {
        return Ok(None); // silent no-op — match oracle
    };

    let scope: String = cfg
        .get_item("scope")?
        .and_then(|v| v.extract::<String>().ok())
        .unwrap_or_else(|| "workbook".to_string());

    let scope_sheet_index: Option<usize> = if scope == "sheet" {
        let idx = wb.sheet_index_by_name(sheet_name).ok_or_else(|| {
            PyValueError::new_err(format!(
                "add_named_range: sheet {sheet_name:?} not found (scope=sheet requires the sheet to exist)"
            ))
        })?;
        Some(idx)
    } else {
        None
    };

    Ok(Some(DefinedName {
        name,
        formula: refers_to,
        scope_sheet_index,
        builtin: None,
        hidden: false,
    }))
}

/// Build a `DocProperties` from a flat props dict.
fn dict_to_doc_properties(props: &Bound<'_, PyDict>) -> PyResult<DocProperties> {
    let title: Option<String> = props.get_item("title")?.and_then(|v| v.extract().ok());
    let subject: Option<String> = props.get_item("subject")?.and_then(|v| v.extract().ok());
    let creator: Option<String> = props.get_item("creator")?.and_then(|v| v.extract().ok());
    let keywords: Option<String> = props.get_item("keywords")?.and_then(|v| v.extract().ok());
    let description: Option<String> =
        props.get_item("description")?.and_then(|v| v.extract().ok());
    let category: Option<String> = props.get_item("category")?.and_then(|v| v.extract().ok());
    // Python passes the OOXML-canonical camelCase key. The Python ->
    // emitter -> <cp:contentStatus> path is preserved verbatim from the
    // W5-removed legacy backend.
    let content_status: Option<String> = props
        .get_item("contentStatus")?
        .and_then(|v| v.extract().ok());

    let created: Option<chrono::NaiveDateTime> =
        props.get_item("created")?.and_then(|v| {
            v.extract::<String>().ok().and_then(|s| {
                // Try datetime first, then date-only (with midnight time).
                parse_iso_datetime(&s).or_else(|| {
                    parse_iso_date(&s).and_then(|d| d.and_hms_opt(0, 0, 0))
                })
            })
        });

    Ok(DocProperties {
        title,
        subject,
        creator,
        keywords,
        description,
        category,
        content_status,
        created,
        ..Default::default()
    })
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

    pub fn add_hyperlink(
        &mut self,
        sheet: &str,
        link_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let dict = link_dict
            .downcast::<PyDict>()
            .map_err(|_| PyValueError::new_err("link must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "hyperlink")?;
        let Some((a1, hyperlink)) = dict_to_hyperlink(&cfg)? else {
            return Ok(()); // silent no-op — match oracle
        };
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.hyperlinks.insert(a1, hyperlink);
        Ok(())
    }

    pub fn add_comment(
        &mut self,
        sheet: &str,
        comment_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let dict = comment_dict
            .downcast::<PyDict>()
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

    pub fn add_conditional_format(
        &mut self,
        sheet: &str,
        rule_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let dict = rule_dict
            .downcast::<PyDict>()
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
            .downcast::<PyDict>()
            .map_err(|_| PyValueError::new_err("validation must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "validation")?;
        let Some(dv) = dict_to_data_validation(&cfg)? else {
            return Ok(()); // silent no-op — match oracle
        };
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.validations.push(dv);
        Ok(())
    }

    pub fn add_named_range(
        &mut self,
        sheet: &str,
        named_range: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let dict = named_range
            .downcast::<PyDict>()
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
            .downcast::<PyDict>()
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
            .downcast::<PyDict>()
            .map_err(|_| PyValueError::new_err("props must be a dict"))?;
        let doc_props = dict_to_doc_properties(dict)?;
        self.inner.set_doc_props(doc_props);
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
            .downcast::<PyDict>()
            .map_err(|_| PyValueError::new_err("image must be a dict"))?;

        let data: Vec<u8> = dict
            .get_item("data")?
            .ok_or_else(|| PyValueError::new_err("image dict missing 'data'"))?
            .extract()?;
        let ext: String = dict
            .get_item("ext")?
            .ok_or_else(|| PyValueError::new_err("image dict missing 'ext'"))?
            .extract()?;
        let width: u32 = dict
            .get_item("width")?
            .ok_or_else(|| PyValueError::new_err("image dict missing 'width'"))?
            .extract()?;
        let height: u32 = dict
            .get_item("height")?
            .ok_or_else(|| PyValueError::new_err("image dict missing 'height'"))?
            .extract()?;
        let anchor_obj = dict
            .get_item("anchor")?
            .ok_or_else(|| PyValueError::new_err("image dict missing 'anchor'"))?;
        let anchor_dict = anchor_obj
            .downcast::<PyDict>()
            .map_err(|_| PyValueError::new_err("anchor must be a dict"))?;

        let anchor = parse_image_anchor(anchor_dict)?;

        let img = SheetImage {
            data,
            ext: ext.to_ascii_lowercase(),
            width_px: width,
            height_px: height,
            anchor,
        };

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
            .downcast::<PyDict>()
            .map_err(|_| PyValueError::new_err("chart must be a dict"))?;
        let chart = parse_chart_dict(dict, anchor_a1)?;
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.charts.push(chart);
        Ok(())
    }
}

fn parse_image_anchor(d: &Bound<'_, PyDict>) -> PyResult<ImageAnchor> {
    let kind: String = d
        .get_item("type")?
        .ok_or_else(|| PyValueError::new_err("anchor dict missing 'type'"))?
        .extract()?;
    match kind.as_str() {
        "one_cell" => {
            let from_col: u32 = anchor_int(d, "from_col", 0)?;
            let from_row: u32 = anchor_int(d, "from_row", 0)?;
            let from_col_off: i64 = anchor_int_i64(d, "from_col_off", 0)?;
            let from_row_off: i64 = anchor_int_i64(d, "from_row_off", 0)?;
            Ok(ImageAnchor::OneCell {
                from_col,
                from_row,
                from_col_off,
                from_row_off,
            })
        }
        "two_cell" => {
            let from_col: u32 = anchor_int(d, "from_col", 0)?;
            let from_row: u32 = anchor_int(d, "from_row", 0)?;
            let from_col_off: i64 = anchor_int_i64(d, "from_col_off", 0)?;
            let from_row_off: i64 = anchor_int_i64(d, "from_row_off", 0)?;
            let to_col: u32 = anchor_int(d, "to_col", 0)?;
            let to_row: u32 = anchor_int(d, "to_row", 0)?;
            let to_col_off: i64 = anchor_int_i64(d, "to_col_off", 0)?;
            let to_row_off: i64 = anchor_int_i64(d, "to_row_off", 0)?;
            let edit_as: String = d
                .get_item("edit_as")?
                .and_then(|v| v.extract().ok())
                .unwrap_or_else(|| "oneCell".to_string());
            Ok(ImageAnchor::TwoCell {
                from_col,
                from_row,
                from_col_off,
                from_row_off,
                to_col,
                to_row,
                to_col_off,
                to_row_off,
                edit_as,
            })
        }
        "absolute" => {
            let x_emu: i64 = anchor_int_i64(d, "x_emu", 0)?;
            let y_emu: i64 = anchor_int_i64(d, "y_emu", 0)?;
            let cx_emu: i64 = anchor_int_i64(d, "cx_emu", 0)?;
            let cy_emu: i64 = anchor_int_i64(d, "cy_emu", 0)?;
            Ok(ImageAnchor::Absolute {
                x_emu,
                y_emu,
                cx_emu,
                cy_emu,
            })
        }
        other => Err(PyValueError::new_err(format!(
            "unknown anchor type: {other:?} (expected one_cell, two_cell, or absolute)"
        ))),
    }
}

fn anchor_int(d: &Bound<'_, PyDict>, key: &str, default: u32) -> PyResult<u32> {
    Ok(d.get_item(key)?
        .and_then(|v| v.extract().ok())
        .unwrap_or(default))
}

fn anchor_int_i64(d: &Bound<'_, PyDict>, key: &str, default: i64) -> PyResult<i64> {
    Ok(d.get_item(key)?
        .and_then(|v| v.extract().ok())
        .unwrap_or(default))
}

// ---------------------------------------------------------------------------
// Sprint Μ Pod-α (RFC-046) — chart dict → typed Chart parsing
// ---------------------------------------------------------------------------

fn parse_chart_dict(d: &Bound<'_, PyDict>, anchor_a1: &str) -> PyResult<Chart> {
    let kind_str: String = d
        .get_item("kind")?
        .ok_or_else(|| PyValueError::new_err("chart dict missing 'kind'"))?
        .extract()?;
    let kind = match kind_str.as_str() {
        "bar" => ChartKind::Bar,
        "line" => ChartKind::Line,
        "pie" => ChartKind::Pie,
        "doughnut" => ChartKind::Doughnut,
        "area" => ChartKind::Area,
        "scatter" => ChartKind::Scatter,
        "bubble" => ChartKind::Bubble,
        "radar" => ChartKind::Radar,
        other => {
            return Err(PyValueError::new_err(format!(
                "unknown chart kind {other:?} (expected bar/line/pie/doughnut/\
                 area/scatter/bubble/radar)"
            )))
        }
    };

    // Anchor: accept (a) explicit dict, (b) A1 string (Pod-β shape),
    // or (c) None / missing — fall back to the call-site `anchor_a1`.
    let anchor = if let Some(v) = d.get_item("anchor")? {
        if v.is_none() {
            a1_to_one_cell_anchor(anchor_a1)?
        } else if let Ok(ad) = v.downcast::<PyDict>() {
            parse_image_anchor(ad)?
        } else if let Ok(s) = v.extract::<String>() {
            a1_to_one_cell_anchor(&s)?
        } else {
            return Err(PyValueError::new_err(
                "chart anchor must be a dict, A1 string, or None",
            ));
        }
    } else {
        a1_to_one_cell_anchor(anchor_a1)?
    };

    let mut chart = Chart::new(kind, anchor);

    if let Some(v) = d.get_item("title")? {
        if !v.is_none() {
            let td = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("chart title must be a dict"))?;
            chart.title = Some(parse_chart_title(td)?);
        }
    }

    if let Some(v) = d.get_item("legend")? {
        if v.is_none() {
            chart.legend = None;
        } else {
            let ld = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("chart legend must be a dict"))?;
            chart.legend = Some(parse_legend(ld)?);
        }
    }

    if let Some(v) = d.get_item("layout")? {
        if !v.is_none() {
            let ld = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("chart layout must be a dict"))?;
            chart.layout = Some(parse_layout(ld)?);
        }
    }

    if let Some(v) = d.get_item("x_axis")? {
        if !v.is_none() {
            let ad = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("x_axis must be a dict"))?;
            chart.x_axis = Some(parse_axis(ad)?);
        }
    }
    if let Some(v) = d.get_item("y_axis")? {
        if !v.is_none() {
            let ad = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("y_axis must be a dict"))?;
            chart.y_axis = Some(parse_axis(ad)?);
        }
    }

    if let Some(v) = d.get_item("series")? {
        let list: Vec<Bound<'_, PyAny>> = v.extract()?;
        for sv in list.iter() {
            let sd = sv
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("each series must be a dict"))?;
            chart.series.push(parse_series(sd)?);
        }
    }

    if let Some(b) = py_opt_bool(d, "plot_visible_only")? {
        chart.plot_visible_only = Some(b);
    }
    if let Some(s) = py_opt_str(d, "display_blanks_as")? {
        chart.display_blanks_as = Some(match s.as_str() {
            "gap" => DisplayBlanksAs::Gap,
            "span" => DisplayBlanksAs::Span,
            "zero" => DisplayBlanksAs::Zero,
            other => {
                return Err(PyValueError::new_err(format!(
                    "unknown display_blanks_as {other:?}"
                )))
            }
        });
    }
    if let Some(b) = py_opt_bool(d, "vary_colors")? {
        chart.vary_colors = Some(b);
    }

    if let Some(s) = py_opt_str(d, "bar_dir")? {
        chart.bar_dir = Some(match s.as_str() {
            "col" => BarDir::Col,
            "bar" => BarDir::Bar,
            other => return Err(PyValueError::new_err(format!("unknown bar_dir {other:?}"))),
        });
    }
    if let Some(s) = py_opt_str(d, "grouping")? {
        chart.grouping = Some(match s.as_str() {
            "clustered" => BarGrouping::Clustered,
            "stacked" => BarGrouping::Stacked,
            "percentStacked" => BarGrouping::PercentStacked,
            "standard" => BarGrouping::Standard,
            other => {
                return Err(PyValueError::new_err(format!(
                    "unknown grouping {other:?}"
                )))
            }
        });
    }
    if let Some(n) = py_opt_u32(d, "gap_width")? {
        chart.gap_width = Some(n);
    }
    if let Some(n) = py_opt_i32(d, "overlap")? {
        chart.overlap = Some(n);
    }
    if let Some(n) = py_opt_u32(d, "hole_size")? {
        chart.hole_size = Some(n);
    }
    if let Some(n) = py_opt_u32(d, "first_slice_ang")? {
        chart.first_slice_ang = Some(n);
    }
    if let Some(s) = py_opt_str(d, "scatter_style")? {
        chart.scatter_style = Some(parse_scatter_style(&s)?);
    }
    if let Some(s) = py_opt_str(d, "radar_style")? {
        chart.radar_style = Some(parse_radar_style(&s)?);
    }
    if let Some(b) = py_opt_bool(d, "bubble3d")? {
        chart.bubble3d = Some(b);
    }
    if let Some(n) = py_opt_u32(d, "bubble_scale")? {
        chart.bubble_scale = Some(n);
    }
    if let Some(b) = py_opt_bool(d, "show_neg_bubbles")? {
        chart.show_neg_bubbles = Some(b);
    }
    if let Some(b) = py_opt_bool(d, "smoothing")? {
        chart.smoothing = Some(b);
    }
    if let Some(n) = py_opt_u32(d, "style")? {
        chart.style = Some(n);
    }

    Ok(chart)
}

fn a1_to_one_cell_anchor(a1: &str) -> PyResult<ImageAnchor> {
    let ((row, col), _) = wolfxl_writer::refs::parse_range(&format!("{a1}:{a1}"))
        .ok_or_else(|| PyValueError::new_err(format!("invalid A1 anchor {a1:?}")))?;
    Ok(ImageAnchor::OneCell {
        from_col: col.saturating_sub(1),
        from_row: row.saturating_sub(1),
        from_col_off: 0,
        from_row_off: 0,
    })
}

fn parse_chart_title(d: &Bound<'_, PyDict>) -> PyResult<ChartTitle> {
    let runs = if let Some(v) = d.get_item("runs")? {
        let list: Vec<Bound<'_, PyAny>> = v.extract()?;
        let mut out = Vec::with_capacity(list.len());
        for rv in list.iter() {
            let rd = rv
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("title run must be a dict"))?;
            let text: String = rd
                .get_item("text")?
                .ok_or_else(|| PyValueError::new_err("title run missing 'text'"))?
                .extract()?;
            out.push(TitleRun {
                text,
                bold: py_opt_bool(rd, "bold")?,
                italic: py_opt_bool(rd, "italic")?,
                underline: py_opt_bool(rd, "underline")?,
                size_pt: py_opt_u32(rd, "size_pt")?,
                color: py_opt_str(rd, "color")?,
                font_name: py_opt_str(rd, "font_name")?,
            });
        }
        out
    } else if let Some(v) = d.get_item("text")? {
        // Convenience: {"text": "Sales"} → single plain run.
        let text: String = v.extract()?;
        vec![TitleRun::plain(text)]
    } else {
        return Err(PyValueError::new_err(
            "chart title must have 'runs' or 'text'",
        ));
    };
    Ok(ChartTitle {
        runs,
        overlay: py_opt_bool(d, "overlay")?,
        layout: parse_optional_layout(d, "layout")?,
    })
}

fn parse_legend(d: &Bound<'_, PyDict>) -> PyResult<Legend> {
    let position = if let Some(s) = py_opt_str(d, "position")? {
        match s.as_str() {
            "r" => LegendPosition::Right,
            "l" => LegendPosition::Left,
            "t" => LegendPosition::Top,
            "b" => LegendPosition::Bottom,
            "tr" => LegendPosition::TopRight,
            other => {
                return Err(PyValueError::new_err(format!(
                    "unknown legend position {other:?}"
                )))
            }
        }
    } else {
        LegendPosition::Right
    };
    Ok(Legend {
        position,
        overlay: py_opt_bool(d, "overlay")?,
        layout: parse_optional_layout(d, "layout")?,
    })
}

fn parse_optional_layout(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<Layout>> {
    if let Some(v) = d.get_item(key)? {
        if !v.is_none() {
            let ld = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err(format!("{key} must be a dict")))?;
            return Ok(Some(parse_layout(ld)?));
        }
    }
    Ok(None)
}

fn parse_layout(d: &Bound<'_, PyDict>) -> PyResult<Layout> {
    let x: f64 = d
        .get_item("x")?
        .ok_or_else(|| PyValueError::new_err("layout missing 'x'"))?
        .extract()?;
    let y: f64 = d
        .get_item("y")?
        .ok_or_else(|| PyValueError::new_err("layout missing 'y'"))?
        .extract()?;
    let w: f64 = d
        .get_item("w")?
        .ok_or_else(|| PyValueError::new_err("layout missing 'w'"))?
        .extract()?;
    let h: f64 = d
        .get_item("h")?
        .ok_or_else(|| PyValueError::new_err("layout missing 'h'"))?
        .extract()?;
    let layout_target = if let Some(s) = py_opt_str(d, "layout_target")? {
        Some(match s.as_str() {
            "inner" => LayoutTarget::Inner,
            "outer" => LayoutTarget::Outer,
            other => {
                return Err(PyValueError::new_err(format!(
                    "unknown layout_target {other:?}"
                )))
            }
        })
    } else {
        None
    };
    Ok(Layout {
        x,
        y,
        w,
        h,
        layout_target,
    })
}

fn parse_axis(d: &Bound<'_, PyDict>) -> PyResult<Axis> {
    let kind: String = d
        .get_item("kind")?
        .ok_or_else(|| PyValueError::new_err("axis dict missing 'kind'"))?
        .extract()?;
    let common = parse_axis_common(d)?;
    match kind.as_str() {
        "category" => Ok(Axis::Category(CategoryAxis {
            common,
            lbl_offset: py_opt_u32(d, "lbl_offset")?,
            lbl_algn: py_opt_str(d, "lbl_algn")?,
        })),
        "value" => Ok(Axis::Value(ValueAxis {
            common,
            min: py_opt_f64(d, "min")?,
            max: py_opt_f64(d, "max")?,
            major_unit: py_opt_f64(d, "major_unit")?,
            minor_unit: py_opt_f64(d, "minor_unit")?,
            crosses: py_opt_str(d, "crosses")?,
        })),
        "date" => Ok(Axis::Date(DateAxis {
            common,
            min: py_opt_f64(d, "min")?,
            max: py_opt_f64(d, "max")?,
            major_unit: py_opt_f64(d, "major_unit")?,
            minor_unit: py_opt_f64(d, "minor_unit")?,
            base_time_unit: py_opt_str(d, "base_time_unit")?,
        })),
        "series" => Ok(Axis::Series(SeriesAxis { common })),
        other => Err(PyValueError::new_err(format!(
            "unknown axis kind {other:?} (expected category|value|date|series)"
        ))),
    }
}

fn parse_axis_common(d: &Bound<'_, PyDict>) -> PyResult<AxisCommon> {
    let ax_id: u32 = d
        .get_item("ax_id")?
        .ok_or_else(|| PyValueError::new_err("axis missing 'ax_id'"))?
        .extract()?;
    let cross_ax: u32 = d
        .get_item("cross_ax")?
        .ok_or_else(|| PyValueError::new_err("axis missing 'cross_ax'"))?
        .extract()?;
    let ax_pos = match py_opt_str(d, "ax_pos")?.as_deref() {
        Some("b") | None => AxisPos::Bottom,
        Some("t") => AxisPos::Top,
        Some("l") => AxisPos::Left,
        Some("r") => AxisPos::Right,
        Some(other) => {
            return Err(PyValueError::new_err(format!(
                "unknown ax_pos {other:?}"
            )))
        }
    };
    let orientation = match py_opt_str(d, "orientation")?.as_deref() {
        Some("minMax") | None => AxisOrientation::MinMax,
        Some("maxMin") => AxisOrientation::MaxMin,
        Some(other) => {
            return Err(PyValueError::new_err(format!(
                "unknown axis orientation {other:?}"
            )))
        }
    };
    let major_tick_mark = if let Some(s) = py_opt_str(d, "major_tick_mark")? {
        Some(parse_tick_mark(&s)?)
    } else {
        None
    };
    let minor_tick_mark = if let Some(s) = py_opt_str(d, "minor_tick_mark")? {
        Some(parse_tick_mark(&s)?)
    } else {
        None
    };
    let title = if let Some(v) = d.get_item("title")? {
        if v.is_none() {
            None
        } else {
            let td = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("axis title must be a dict"))?;
            Some(parse_chart_title(td)?)
        }
    } else {
        None
    };
    Ok(AxisCommon {
        ax_id,
        cross_ax,
        orientation,
        ax_pos,
        delete: py_opt_bool(d, "delete")?,
        major_tick_mark,
        minor_tick_mark,
        title,
        major_gridlines: py_opt_bool(d, "major_gridlines")?.unwrap_or(false),
        minor_gridlines: py_opt_bool(d, "minor_gridlines")?.unwrap_or(false),
        number_format: py_opt_str(d, "number_format")?,
    })
}

fn parse_tick_mark(s: &str) -> PyResult<TickMark> {
    Ok(match s {
        "none" => TickMark::None,
        "in" => TickMark::In,
        "out" => TickMark::Out,
        "cross" => TickMark::Cross,
        other => {
            return Err(PyValueError::new_err(format!(
                "unknown tick mark {other:?}"
            )))
        }
    })
}

fn parse_series(d: &Bound<'_, PyDict>) -> PyResult<Series> {
    let idx: u32 = d
        .get_item("idx")?
        .and_then(|v| v.extract().ok())
        .unwrap_or(0);
    let order: u32 = d
        .get_item("order")?
        .and_then(|v| v.extract().ok())
        .unwrap_or(idx);
    let mut s = Series::new(idx);
    s.order = order;

    if let Some(v) = d.get_item("title")? {
        if !v.is_none() {
            let td = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("series title must be a dict"))?;
            if let Some(rv) = td.get_item("strRef")? {
                let rd = rv
                    .downcast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("strRef must be a dict"))?;
                s.title = Some(SeriesTitle::StrRef(parse_reference(rd)?));
            } else if let Some(lv) = td.get_item("literal")? {
                let s_str: String = lv.extract()?;
                s.title = Some(SeriesTitle::Literal(s_str));
            }
        }
    }

    if let Some(v) = d.get_item("categories")? {
        if !v.is_none() {
            let rd = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("categories must be a dict"))?;
            s.categories = Some(parse_reference(rd)?);
        }
    }
    if let Some(v) = d.get_item("values")? {
        if !v.is_none() {
            let rd = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("values must be a dict"))?;
            s.values = Some(parse_reference(rd)?);
        }
    }
    if let Some(v) = d.get_item("x_values")? {
        if !v.is_none() {
            let rd = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("x_values must be a dict"))?;
            s.x_values = Some(parse_reference(rd)?);
        }
    }
    if let Some(v) = d.get_item("bubble_size")? {
        if !v.is_none() {
            let rd = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("bubble_size must be a dict"))?;
            s.bubble_size = Some(parse_reference(rd)?);
        }
    }

    if let Some(v) = d.get_item("graphical_properties")? {
        if !v.is_none() {
            let gd = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("graphical_properties must be a dict"))?;
            s.graphical_properties = Some(parse_graphical_properties(gd)?);
        }
    }
    if let Some(v) = d.get_item("marker")? {
        if !v.is_none() {
            let md = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("marker must be a dict"))?;
            s.marker = Some(parse_marker(md)?);
        }
    }
    if let Some(v) = d.get_item("data_labels")? {
        if !v.is_none() {
            let dd = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("data_labels must be a dict"))?;
            s.data_labels = Some(parse_data_labels(dd)?);
        }
    }
    if let Some(v) = d.get_item("error_bars")? {
        if !v.is_none() {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            for ev in list.iter() {
                let ed = ev
                    .downcast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("error bar must be a dict"))?;
                s.error_bars.push(parse_error_bars(ed)?);
            }
        }
    }
    if let Some(v) = d.get_item("trendlines")? {
        if !v.is_none() {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            for tv in list.iter() {
                let td = tv
                    .downcast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("trendline must be a dict"))?;
                s.trendlines.push(parse_trendline(td)?);
            }
        }
    }

    s.smooth = py_opt_bool(d, "smooth")?;
    s.invert_if_negative = py_opt_bool(d, "invert_if_negative")?;
    Ok(s)
}

fn parse_reference(d: &Bound<'_, PyDict>) -> PyResult<ChartReference> {
    let sheet: String = d
        .get_item("sheet")?
        .ok_or_else(|| PyValueError::new_err("reference missing 'sheet'"))?
        .extract()?;
    let range: String = d
        .get_item("range")?
        .ok_or_else(|| PyValueError::new_err("reference missing 'range'"))?
        .extract()?;
    Ok(ChartReference::new(sheet, range))
}

fn parse_graphical_properties(d: &Bound<'_, PyDict>) -> PyResult<GraphicalProperties> {
    Ok(GraphicalProperties {
        line_color: py_opt_str(d, "line_color")?,
        line_width_emu: py_opt_u32(d, "line_width_emu")?,
        line_dash: py_opt_str(d, "line_dash")?,
        fill_color: py_opt_str(d, "fill_color")?,
        no_fill: py_opt_bool(d, "no_fill")?.unwrap_or(false),
        no_line: py_opt_bool(d, "no_line")?.unwrap_or(false),
    })
}

fn parse_marker(d: &Bound<'_, PyDict>) -> PyResult<Marker> {
    let symbol = match py_opt_str(d, "symbol")?.as_deref() {
        Some("none") => MarkerSymbol::None,
        Some("circle") | None => MarkerSymbol::Circle,
        Some("square") => MarkerSymbol::Square,
        Some("diamond") => MarkerSymbol::Diamond,
        Some("triangle") => MarkerSymbol::Triangle,
        Some("plus") => MarkerSymbol::Plus,
        Some("x") => MarkerSymbol::X,
        Some("star") => MarkerSymbol::Star,
        Some("dash") => MarkerSymbol::Dash,
        Some("dot") => MarkerSymbol::Dot,
        Some("auto") => MarkerSymbol::Auto,
        Some(other) => {
            return Err(PyValueError::new_err(format!(
                "unknown marker symbol {other:?}"
            )))
        }
    };
    let graphical_properties = if let Some(v) = d.get_item("graphical_properties")? {
        if v.is_none() {
            None
        } else {
            let gd = v
                .downcast::<PyDict>()
                .map_err(|_| PyValueError::new_err("graphical_properties must be a dict"))?;
            Some(parse_graphical_properties(gd)?)
        }
    } else {
        None
    };
    Ok(Marker {
        symbol,
        size: py_opt_u32(d, "size")?,
        graphical_properties,
    })
}

fn parse_data_labels(d: &Bound<'_, PyDict>) -> PyResult<DataLabels> {
    Ok(DataLabels {
        show_val: py_opt_bool(d, "show_val")?,
        show_cat_name: py_opt_bool(d, "show_cat_name")?,
        show_ser_name: py_opt_bool(d, "show_ser_name")?,
        show_percent: py_opt_bool(d, "show_percent")?,
        show_legend_key: py_opt_bool(d, "show_legend_key")?,
        show_bubble_size: py_opt_bool(d, "show_bubble_size")?,
        position: py_opt_str(d, "position")?,
        number_format: py_opt_str(d, "number_format")?,
        separator: py_opt_str(d, "separator")?,
    })
}

fn parse_error_bars(d: &Bound<'_, PyDict>) -> PyResult<ErrorBars> {
    let bar_type = match py_opt_str(d, "bar_type")?.as_deref() {
        Some("plus") => ErrorBarType::Plus,
        Some("minus") => ErrorBarType::Minus,
        Some("both") | None => ErrorBarType::Both,
        Some(other) => {
            return Err(PyValueError::new_err(format!(
                "unknown error bar type {other:?}"
            )))
        }
    };
    let val_type = match py_opt_str(d, "val_type")?.as_deref() {
        Some("cust") => ErrorBarValType::Cust,
        Some("fixedVal") => ErrorBarValType::FixedVal,
        Some("percentage") => ErrorBarValType::Percentage,
        Some("stdDev") => ErrorBarValType::StdDev,
        Some("stdErr") | None => ErrorBarValType::StdErr,
        Some(other) => {
            return Err(PyValueError::new_err(format!(
                "unknown error bar val_type {other:?}"
            )))
        }
    };
    Ok(ErrorBars {
        bar_type,
        val_type,
        value: py_opt_f64(d, "value")?,
        no_end_cap: py_opt_bool(d, "no_end_cap")?,
    })
}

fn parse_trendline(d: &Bound<'_, PyDict>) -> PyResult<Trendline> {
    let kind = match py_opt_str(d, "kind")?.as_deref() {
        Some("linear") | None => TrendlineKind::Linear,
        Some("log") => TrendlineKind::Log,
        Some("power") => TrendlineKind::Power,
        Some("exp") => TrendlineKind::Exp,
        Some("poly") => TrendlineKind::Polynomial,
        Some("movingAvg") => TrendlineKind::MovingAvg,
        Some(other) => {
            return Err(PyValueError::new_err(format!(
                "unknown trendline kind {other:?}"
            )))
        }
    };
    Ok(Trendline {
        kind,
        order: py_opt_u32(d, "order")?,
        period: py_opt_u32(d, "period")?,
        forward: py_opt_f64(d, "forward")?,
        backward: py_opt_f64(d, "backward")?,
        display_equation: py_opt_bool(d, "display_equation")?,
        display_r_squared: py_opt_bool(d, "display_r_squared")?,
        name: py_opt_str(d, "name")?,
    })
}

fn parse_scatter_style(s: &str) -> PyResult<ScatterStyle> {
    Ok(match s {
        "line" => ScatterStyle::Line,
        "lineMarker" => ScatterStyle::LineMarker,
        "marker" => ScatterStyle::Marker,
        "smooth" => ScatterStyle::Smooth,
        "smoothMarker" => ScatterStyle::SmoothMarker,
        "none" => ScatterStyle::None,
        other => {
            return Err(PyValueError::new_err(format!(
                "unknown scatter_style {other:?}"
            )))
        }
    })
}

fn parse_radar_style(s: &str) -> PyResult<RadarStyle> {
    Ok(match s {
        "standard" => RadarStyle::Standard,
        "marker" => RadarStyle::Marker,
        "filled" => RadarStyle::Filled,
        other => {
            return Err(PyValueError::new_err(format!(
                "unknown radar_style {other:?}"
            )))
        }
    })
}

fn py_opt_bool(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<bool>> {
    if let Some(v) = d.get_item(key)? {
        if v.is_none() {
            return Ok(None);
        }
        return Ok(Some(v.extract()?));
    }
    Ok(None)
}

fn py_opt_str(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<String>> {
    if let Some(v) = d.get_item(key)? {
        if v.is_none() {
            return Ok(None);
        }
        return Ok(Some(v.extract()?));
    }
    Ok(None)
}

fn py_opt_u32(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<u32>> {
    if let Some(v) = d.get_item(key)? {
        if v.is_none() {
            return Ok(None);
        }
        return Ok(Some(v.extract()?));
    }
    Ok(None)
}

fn py_opt_i32(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<i32>> {
    if let Some(v) = d.get_item(key)? {
        if v.is_none() {
            return Ok(None);
        }
        return Ok(Some(v.extract()?));
    }
    Ok(None)
}

fn py_opt_f64(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<f64>> {
    if let Some(v) = d.get_item(key)? {
        if v.is_none() {
            return Ok(None);
        }
        return Ok(Some(v.extract()?));
    }
    Ok(None)
}

