//! `[Content_Types].xml` emitter. Wave 1C subagent fills this in.
//!
//! This is the map of file-extension → MIME-type overrides that the
//! xlsx container needs. Every part inside the ZIP must be accounted
//! for here or Excel flags the file as corrupt.

use crate::model::workbook::Workbook;

/// Emit `[Content_Types].xml` as UTF-8 bytes.
///
/// **Stub** — Wave 1C fills in the real emitter.
pub fn emit(_wb: &Workbook) -> Vec<u8> {
    Vec::new()
}
