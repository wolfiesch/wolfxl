//! Named-range reader logic for native XLSX/XLSB books.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use crate::native_reader_backend::{NativeXlsbBook, NativeXlsxBook};

type PyObject = Py<PyAny>;

pub(crate) fn read_named_ranges_xlsx(
    book: &NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    if !book.sheet_names.iter().any(|name| name == sheet) {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Unknown sheet: {sheet}"
        )));
    }
    let result = PyList::empty(py);
    for named_range in book.book.named_ranges() {
        if named_range.scope == "sheet" {
            let refers_to = named_range.refers_to.trim_start_matches('=');
            let Some((sheet_part, _addr)) = refers_to.split_once('!') else {
                continue;
            };
            if sheet_part.trim_matches('\'') != sheet {
                continue;
            }
        }
        let d = PyDict::new(py);
        d.set_item("name", &named_range.name)?;
        d.set_item("scope", &named_range.scope)?;
        d.set_item("refers_to", &named_range.refers_to)?;
        result.append(d)?;
    }
    Ok(result.into())
}

pub(crate) fn read_named_ranges_xlsb(
    book: &NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    if !book.sheet_names.iter().any(|name| name == sheet) {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Unknown sheet: {sheet}"
        )));
    }
    let result = PyList::empty(py);
    for named_range in book.book.named_ranges() {
        if named_range.scope == "sheet" {
            let refers_to = named_range.refers_to.trim_start_matches('=');
            let Some((sheet_part, _addr)) = refers_to.split_once('!') else {
                continue;
            };
            if sheet_part.trim_matches('\'') != sheet {
                continue;
            }
        }
        let d = PyDict::new(py);
        d.set_item("name", &named_range.name)?;
        d.set_item("scope", &named_range.scope)?;
        d.set_item("refers_to", &named_range.refers_to)?;
        result.append(d)?;
    }
    Ok(result.into())
}
