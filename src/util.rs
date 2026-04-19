use pyo3::prelude::*;
use pyo3::types::PyDict;
use pyo3::IntoPyObject;

use chrono::{NaiveDate, NaiveDateTime};

type PyObject = Py<PyAny>;

pub fn a1_to_row_col(a1: &str) -> Result<(u32, u32), String> {
    let mut col: u32 = 0;
    let mut row_digits = String::new();

    for ch in a1.chars() {
        if ch.is_ascii_alphabetic() {
            let uc = ch.to_ascii_uppercase() as u8;
            let val = (uc - b'A' + 1) as u32;
            col = col * 26 + val;
        } else if ch.is_ascii_digit() {
            row_digits.push(ch);
        } else {
            return Err(format!("Invalid cell reference: {a1}"));
        }
    }

    if col == 0 || row_digits.is_empty() {
        return Err(format!("Invalid cell reference: {a1}"));
    }

    let row_1: u32 = row_digits
        .parse()
        .map_err(|_| format!("Invalid cell reference: {a1}"))?;
    if row_1 == 0 {
        return Err(format!("Invalid cell reference: {a1}"));
    }

    Ok((row_1 - 1, col - 1))
}

pub(crate) fn cell_blank(py: Python<'_>) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    // The Python harness treats missing "value" as blank.
    d.set_item("type", "blank")?;
    Ok(d.into())
}

pub(crate) fn cell_with_value<'py, V: IntoPyObject<'py>>(
    py: Python<'py>,
    t: &str,
    value: V,
) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("type", t)?;
    d.set_item("value", value)?;
    Ok(d.into())
}

pub(crate) fn parse_iso_date(s: &str) -> Option<NaiveDate> {
    NaiveDate::parse_from_str(s, "%Y-%m-%d").ok()
}

pub(crate) fn parse_iso_datetime(s: &str) -> Option<NaiveDateTime> {
    let raw = s.trim_end_matches('Z');
    NaiveDateTime::parse_from_str(raw, "%Y-%m-%dT%H:%M:%S")
        .ok()
        .or_else(|| NaiveDateTime::parse_from_str(raw, "%Y-%m-%dT%H:%M:%S%.f").ok())
}
