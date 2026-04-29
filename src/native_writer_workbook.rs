//! Workbook-level helpers for the native writer backend.

use std::fs;

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use wolfxl_writer::model::Worksheet;
use wolfxl_writer::Workbook;

pub(crate) fn add_sheet_if_missing(wb: &mut Workbook, name: &str) {
    if wb.sheet_by_name(name).is_some() {
        return;
    }
    wb.add_sheet(Worksheet::new(name));
}

pub(crate) fn rename_sheet(wb: &mut Workbook, old_name: &str, new_name: &str) -> PyResult<()> {
    wb.rename_sheet(old_name, new_name.to_string())
        .map_err(PyValueError::new_err)
}

pub(crate) fn move_sheet(wb: &mut Workbook, name: &str, offset: isize) -> PyResult<()> {
    wb.move_sheet(name, offset).map_err(PyValueError::new_err)
}

pub(crate) fn save_once(wb: &mut Workbook, saved: &mut bool, path: &str) -> PyResult<()> {
    if *saved {
        return Err(PyValueError::new_err(
            "Workbook already saved (NativeWorkbook is consumed-on-save)",
        ));
    }
    // Mark consumed before emit/write so a panic or failed write leaves the
    // workbook un-retryable on potentially mutated state.
    *saved = true;
    let bytes = wolfxl_writer::emit_xlsx(wb);
    fs::write(path, bytes)
        .map_err(|e| PyIOError::new_err(format!("failed to write {path}: {e}")))?;
    Ok(())
}
