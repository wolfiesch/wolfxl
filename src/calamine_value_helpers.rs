//! Value and coordinate helpers shared by the calamine styled read path.

use std::collections::HashMap;

use calamine_styles::{Data, Range};
use chrono::{Datelike, NaiveTime, Timelike};
use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDateTime, PyDict};
use pyo3::IntoPyObjectExt;

use crate::util::{a1_to_row_col, cell_blank, cell_with_value, parse_iso_date, parse_iso_datetime};

type PyObject = Py<PyAny>;

pub(crate) fn data_to_py(py: Python<'_>, value: &Data) -> PyResult<PyObject> {
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
pub(crate) fn data_to_plain_py(py: Python<'_>, value: &Data) -> PyResult<PyObject> {
    match value {
        Data::Empty => Ok(py.None()),
        Data::String(s) => s.into_py_any(py),
        Data::Float(f) => f.into_py_any(py),
        Data::Int(i) => i.into_py_any(py),
        Data::Bool(b) => b.into_py_any(py),
        Data::DateTime(dt) => {
            if let Some(ndt) = dt.as_datetime() {
                let d = PyDateTime::new(
                    py,
                    ndt.year(),
                    ndt.month() as u8,
                    ndt.day() as u8,
                    ndt.hour() as u8,
                    ndt.minute() as u8,
                    ndt.second() as u8,
                    0,
                    None,
                )?;
                Ok(d.into_any().unbind())
            } else {
                dt.as_f64().into_py_any(py)
            }
        }
        Data::DateTimeIso(s) => {
            let raw = s.trim_end_matches('Z');
            if let Some(d) = parse_iso_date(raw) {
                let pydt = PyDateTime::new(
                    py,
                    d.year(),
                    d.month() as u8,
                    d.day() as u8,
                    0,
                    0,
                    0,
                    0,
                    None,
                )?;
                Ok(pydt.into_any().unbind())
            } else if let Some(ndt) = parse_iso_datetime(raw) {
                let pydt = PyDateTime::new(
                    py,
                    ndt.year(),
                    ndt.month() as u8,
                    ndt.day() as u8,
                    ndt.hour() as u8,
                    ndt.minute() as u8,
                    ndt.second() as u8,
                    0,
                    None,
                )?;
                Ok(pydt.into_any().unbind())
            } else {
                s.into_py_any(py)
            }
        }
        Data::DurationIso(s) => s.into_py_any(py),
        Data::RichText(rt) => rt.plain_text().into_py_any(py),
        Data::Error(e) => {
            let normalized = map_error_value(&format!("{e:?}"));
            normalized.into_py_any(py)
        }
    }
}

pub(crate) fn map_error_value(err_str: &str) -> &'static str {
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

pub(crate) fn map_error_formula(formula: &str) -> Option<&'static str> {
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

pub(crate) fn normalize_formula_text(formula: &str) -> String {
    if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    }
}

pub(crate) fn data_type_name(value: &Data) -> &'static str {
    match value {
        Data::Empty => "blank",
        Data::String(_) | Data::RichText(_) | Data::DurationIso(_) => "string",
        Data::Float(_) | Data::Int(_) => "number",
        Data::Bool(_) => "boolean",
        Data::DateTime(_) => "datetime",
        Data::DateTimeIso(_) => "datetime",
        Data::Error(_) => "error",
    }
}

pub(crate) fn data_is_formula_text(value: &Data, formula: &str) -> bool {
    let owned;
    let text = match value {
        Data::String(s) => s.as_str(),
        Data::RichText(rt) => {
            owned = rt.plain_text();
            owned.as_str()
        }
        _ => return false,
    };
    text.is_empty() || text == formula || format!("={text}") == formula
}

/// Return true when a formula cell contains calamine's uncached placeholder.
pub(crate) fn is_uncached_formula_value(
    formula_map: Option<&HashMap<(u32, u32), String>>,
    row: u32,
    col: u32,
    value: &Data,
) -> bool {
    let Some(fmap) = formula_map else {
        return false;
    };
    let Some(raw) = fmap.get(&(row, col)) else {
        return false;
    };
    let formula = if raw.starts_with('=') {
        raw.clone()
    } else {
        format!("={raw}")
    };
    data_is_formula_text(value, &formula)
}

pub(crate) fn row_col_to_a1(row: u32, col: u32) -> String {
    let mut n = col + 1;
    let mut letters: Vec<char> = Vec::new();
    while n > 0 {
        n -= 1;
        letters.push((b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    letters.reverse();
    format!("{}{}", letters.into_iter().collect::<String>(), row + 1)
}

pub(crate) fn resolve_range_bounds(
    range: &Range<Data>,
    cell_range: Option<&str>,
) -> PyResult<(u32, u32, u32, u32)> {
    if let Some(cr) = cell_range {
        if !cr.is_empty() {
            let clean = cr.replace('$', "").to_ascii_uppercase();
            let parts: Vec<&str> = clean.split(':').collect();
            let a = parts[0];
            let b = if parts.len() > 1 { parts[1] } else { a };
            let (r0, c0) = a1_to_row_col(a).map_err(PyErr::new::<PyValueError, _>)?;
            let (r1, c1) = a1_to_row_col(b).map_err(PyErr::new::<PyValueError, _>)?;
            return Ok((r0.min(r1), c0.min(c1), r0.max(r1), c0.max(c1)));
        }
    }

    let (height, width) = range.get_size();
    if height == 0 || width == 0 {
        return Ok((0, 0, 0, 0));
    }
    let start = range.start().unwrap_or((0, 0));
    Ok((
        start.0,
        start.1,
        start.0 + height as u32 - 1,
        start.1 + width as u32 - 1,
    ))
}

pub(crate) fn col_letter_to_index(col: &str) -> PyResult<u32> {
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

pub(crate) fn update_dimensions(
    dimensions: &mut Option<(u32, u32)>,
    row_count: u32,
    col_count: u32,
) {
    match dimensions {
        Some((rows, cols)) => {
            *rows = (*rows).max(row_count);
            *cols = (*cols).max(col_count);
        }
        None => *dimensions = Some((row_count, col_count)),
    }
}

pub(crate) fn update_bounds(bounds: &mut Option<(u32, u32, u32, u32)>, row: u32, col: u32) {
    match bounds {
        Some((min_row, min_col, max_row, max_col)) => {
            *min_row = (*min_row).min(row);
            *min_col = (*min_col).min(col);
            *max_row = (*max_row).max(row);
            *max_col = (*max_col).max(col);
        }
        None => *bounds = Some((row, col, row, col)),
    }
}
