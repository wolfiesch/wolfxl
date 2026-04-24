//! `docProps/core.xml` + `docProps/app.xml` emitter. Wave 1C.

use crate::model::workbook::Workbook;

/// `docProps/core.xml` — Dublin Core + basic Office metadata.
pub fn emit_core(_wb: &Workbook) -> Vec<u8> {
    Vec::new()
}

/// `docProps/app.xml` — application-specific metadata (sheet names list,
/// application version, etc.).
pub fn emit_app(_wb: &Workbook) -> Vec<u8> {
    Vec::new()
}
