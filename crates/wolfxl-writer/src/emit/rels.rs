//! `.rels` emitter — root rels, workbook rels, and per-sheet rels.
//! Wave 1C subagent fills this in.
//!
//! OOXML relationships are a layer of indirection between parts: every
//! `r:id="rId5"` attribute somewhere in the workbook is resolved through
//! a `.rels` file to the actual target part path.

use crate::model::workbook::Workbook;

/// `_rels/.rels` — top-level relationships (workbook, core props, app props).
pub fn emit_root(_wb: &Workbook) -> Vec<u8> {
    Vec::new()
}

/// `xl/_rels/workbook.xml.rels` — workbook → sheets, styles, shared strings.
pub fn emit_workbook(_wb: &Workbook) -> Vec<u8> {
    Vec::new()
}

/// `xl/worksheets/_rels/sheet{N}.xml.rels` — sheet → comments, drawings,
/// tables, hyperlinks.
pub fn emit_sheet(_wb: &Workbook, _sheet_idx: usize) -> Vec<u8> {
    Vec::new()
}
