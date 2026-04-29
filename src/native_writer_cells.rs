//! Cell-value write helpers for the native writer backend.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use wolfxl_writer::model::date::{date_to_excel_serial, datetime_to_excel_serial};
use wolfxl_writer::model::{FormatSpec, Worksheet, WriteCellValue};
use wolfxl_writer::refs;
use wolfxl_writer::Workbook;

use crate::native_writer_rich_text::py_runs_to_rust_writer;
use crate::util::{parse_iso_date, parse_iso_datetime};

pub(crate) fn parse_a1_to_row_col(a1: &str) -> PyResult<(u32, u32)> {
    let cleaned = a1.replace('$', "");
    refs::parse_a1(&cleaned)
        .ok_or_else(|| PyValueError::new_err(format!("Invalid A1 reference: {a1}")))
}

pub(crate) fn write_cell_payload(
    wb: &mut Workbook,
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
        Some(wb.styles.intern_format(&spec))
    } else {
        None
    };

    let ws = require_sheet(wb, sheet)?;
    ws.write_cell(row, col, value, style_id);
    Ok(())
}

/// Write a rich-text inline-string cell from Python ``(text, font)`` runs.
pub(crate) fn write_rich_text_cell(
    wb: &mut Workbook,
    sheet: &str,
    a1: &str,
    runs: &Bound<'_, PyList>,
) -> PyResult<()> {
    let (row, col) = parse_a1_to_row_col(a1)?;
    let parsed = py_runs_to_rust_writer(runs)?;
    let ws = require_sheet(wb, sheet)?;
    ws.write_cell(row, col, WriteCellValue::InlineRichText(parsed), None);
    Ok(())
}

pub(crate) fn write_array_formula_cell(
    wb: &mut Workbook,
    sheet: &str,
    a1: &str,
    payload: &Bound<'_, PyDict>,
) -> PyResult<()> {
    let (row, col) = parse_a1_to_row_col(a1)?;
    let value = array_formula_payload_to_write_cell_value(payload)?;
    let ws = require_sheet(wb, sheet)?;
    ws.write_cell(row, col, value, None);
    Ok(())
}

/// Bulk-write a rectangular grid of raw Python values starting at `start_a1`.
pub(crate) fn write_value_grid(
    wb: &mut Workbook,
    sheet: &str,
    start_a1: &str,
    values: &Bound<'_, PyAny>,
) -> PyResult<()> {
    let (base_row, base_col) = parse_a1_to_row_col(start_a1)?;
    let ws = require_sheet(wb, sheet)?;
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

/// Convert oracle-shape cell payload dict into a `WriteCellValue`.
fn payload_to_write_cell_value(payload: &Bound<'_, PyAny>) -> PyResult<WriteCellValue> {
    let dict = payload
        .cast::<PyDict>()
        .map_err(|_| PyValueError::new_err("payload must be a dict"))?;

    let type_str: String = dict
        .get_item("type")?
        .ok_or_else(|| PyValueError::new_err("payload missing 'type'"))?
        .extract()?;

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
            Ok(WriteCellValue::Number(require_finite_f64(
                n,
                "number cell",
            )?))
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
            Ok(WriteCellValue::Formula { expr, result: None })
        }
        "error" => Ok(WriteCellValue::String(value_str.unwrap_or_default())),
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

fn require_finite_f64(f: f64, context: &str) -> PyResult<f64> {
    if !f.is_finite() {
        return Err(PyValueError::new_err(format!(
            "{context}: non-finite floats (NaN, Infinity) are not representable in xlsx; got {f}",
        )));
    }
    Ok(f)
}

/// Coerce a raw Python value from `write_sheet_values` to a writer cell value.
fn raw_python_to_write_cell_value(value: &Bound<'_, PyAny>) -> PyResult<Option<WriteCellValue>> {
    if value.is_none() {
        return Ok(None);
    }

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
        return Ok(Some(WriteCellValue::Number(require_finite_f64(
            f,
            "cell value",
        )?)));
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

fn array_formula_payload_to_write_cell_value(
    payload: &Bound<'_, PyDict>,
) -> PyResult<WriteCellValue> {
    let kind: String = payload
        .get_item("kind")?
        .ok_or_else(|| PyValueError::new_err("payload missing 'kind'"))?
        .extract()?;

    let value = match kind.as_str() {
        "array" => {
            let ref_range: String = payload
                .get_item("ref")?
                .ok_or_else(|| PyValueError::new_err("array kind needs 'ref'"))?
                .extract()?;
            let mut text: String = payload
                .get_item("text")?
                .ok_or_else(|| PyValueError::new_err("array kind needs 'text'"))?
                .extract()?;
            if let Some(stripped) = text.strip_prefix('=') {
                text = stripped.to_string();
            }
            WriteCellValue::ArrayFormula { ref_range, text }
        }
        "data_table" => {
            let ref_range: String = payload
                .get_item("ref")?
                .ok_or_else(|| PyValueError::new_err("data_table kind needs 'ref'"))?
                .extract()?;
            let ca: bool = payload
                .get_item("ca")?
                .map(|v| v.extract::<bool>())
                .transpose()?
                .unwrap_or(false);
            let dt2_d: bool = payload
                .get_item("dt2D")?
                .map(|v| v.extract::<bool>())
                .transpose()?
                .unwrap_or(false);
            let dtr: bool = payload
                .get_item("dtr")?
                .map(|v| v.extract::<bool>())
                .transpose()?
                .unwrap_or(false);
            let r1: Option<String> = payload.get_item("r1")?.and_then(|v| v.extract().ok());
            let r2: Option<String> = payload.get_item("r2")?.and_then(|v| v.extract().ok());
            WriteCellValue::DataTableFormula {
                ref_range,
                ca,
                dt2_d,
                dtr,
                r1,
                r2,
            }
        }
        "spill_child" => WriteCellValue::SpillChild,
        other => {
            return Err(PyValueError::new_err(format!(
                "Unknown array-formula kind: '{other}'"
            )))
        }
    };

    Ok(value)
}

fn require_sheet<'wb>(wb: &'wb mut Workbook, name: &str) -> PyResult<&'wb mut Worksheet> {
    wb.sheet_mut_by_name(name)
        .ok_or_else(|| PyValueError::new_err(format!("Unknown sheet: {name}")))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn finite_guard_accepts_normal_numbers() {
        assert_eq!(require_finite_f64(1.25, "cell").unwrap(), 1.25);
    }

    #[test]
    fn finite_guard_rejects_nan() {
        assert!(require_finite_f64(f64::NAN, "cell").is_err());
    }

    #[test]
    fn python_bool_parser_matches_flush_tokens() {
        assert!(parse_python_bool("YES"));
        assert!(parse_python_bool("1"));
        assert!(!parse_python_bool("false"));
    }
}
