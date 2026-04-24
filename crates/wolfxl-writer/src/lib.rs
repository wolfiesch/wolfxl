//! wolfxl-writer — native xlsx writer for wolfxl.
//!
//! # Design
//!
//! A pure-Rust OOXML emitter that replaces `rust_xlsxwriter`. The writer is
//! split into three layers:
//!
//! 1. [`model`] — pure data (Workbook, Worksheet, Row, WriteCell, format specs).
//!    No I/O, no XML. You build a model, then hand it to the emitter.
//! 2. [`emit`] — one module per OOXML part (styles.xml, sheet1.xml, workbook.xml, ...).
//!    Each module takes the relevant slice of the model and returns UTF-8 bytes.
//! 3. [`zip`] — deterministic ZIP packager that assembles emitted parts into a
//!    valid xlsx container.
//!
//! The top-level [`Workbook`] facade orchestrates all three.
//!
//! # Determinism
//!
//! Byte-identical output is a non-goal for shipping but a gold-star target for
//! the differential test harness. To get close:
//!
//! - `WOLFXL_TEST_EPOCH=0` env var → ZIP entry mtimes are forced to the Unix
//!   epoch (1970-01-01) so two runs produce identical bytes.
//! - `BTreeMap` for row/cell collections → emission order matches the OOXML
//!   spec's "sorted ascending by `r` attribute" rule without an extra sort pass.
//! - `IndexMap` for author lists → preserves insertion order (fixes the
//!   `rust_xlsxwriter` BTreeMap bug that corrupted mixed-author comment files).
//!
//! # Status
//!
//! Skeleton — public surface is stubbed while Wave 1 subagents fill in
//! `refs` and `zip`. Full usage docs arrive at Wave 4 integration.

pub mod emit;
pub mod intern;
pub mod model;
pub mod refs;
pub mod xml_escape;
pub mod zip;

#[cfg(test)]
mod test_utils;

pub use model::workbook::Workbook;
pub use model::worksheet::Worksheet;
