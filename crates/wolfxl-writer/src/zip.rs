//! Deterministic xlsx ZIP packager. Wave 1C subagent fills this in.
//!
//! # Behavior (when complete)
//!
//! - Writes xlsx entries in the canonical order Excel produces so diffs
//!   against reference files are shallow:
//!   1. `[Content_Types].xml`
//!   2. `_rels/.rels`
//!   3. `xl/workbook.xml`
//!   4. `xl/_rels/workbook.xml.rels`
//!   5. `xl/worksheets/sheet1.xml`, sheet2, ...
//!   6. `xl/worksheets/_rels/sheet*.xml.rels`
//!   7. `xl/theme/theme1.xml`
//!   8. `xl/styles.xml`
//!   9. `xl/sharedStrings.xml`
//!   10. `xl/tables/table*.xml`
//!   11. `xl/comments/comments*.xml` + `xl/drawings/vmlDrawing*.vml`
//!   12. `docProps/core.xml`, `docProps/app.xml`
//!
//! - Stamps each entry's mtime from `WOLFXL_TEST_EPOCH` if set (for diff
//!   harness byte parity), otherwise from wall clock.
//!
//! - DEFLATE level 6 for XML parts, STORE for small stubs — TBD by 1C.

use std::io::Write;

/// A single (path, bytes) pair awaiting packaging. Construct one per
/// emitted OOXML part; hand a `Vec<ZipEntry>` to the packager.
#[derive(Debug, Clone)]
pub struct ZipEntry {
    /// The full path inside the xlsx, e.g. `"xl/worksheets/sheet1.xml"`.
    pub path: String,
    pub bytes: Vec<u8>,
}

/// Package a sequence of entries into a complete xlsx. Returns the
/// serialized container bytes.
///
/// **Stub** — Wave 1C subagent fills this in.
pub fn package(entries: &[ZipEntry]) -> Result<Vec<u8>, std::io::Error> {
    // Placeholder: minimal non-compressing wrapper so the crate compiles.
    // The real packager uses the `zip` crate with deterministic ordering
    // and mtime handling.
    let _ = entries;
    let mut buf: Vec<u8> = Vec::new();
    buf.write_all(b"PK")?; // ZIP magic, just enough to not be empty
    Ok(buf)
}

/// Read the `WOLFXL_TEST_EPOCH` env var; if set (to any value including
/// "0"), return it as the mtime to stamp on every entry. Otherwise
/// return `None` and the packager uses wall-clock time.
pub fn test_epoch_override() -> Option<i64> {
    std::env::var("WOLFXL_TEST_EPOCH").ok().and_then(|s| s.parse().ok())
}
