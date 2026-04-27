//! Reusable serializers for OOXML fragments shared between the native
//! writer and the patcher. Each module is PyO3-free so the patcher (a
//! separate binary in `src/wolfxl/`) can depend on it without pulling
//! in the Python ABI.
//!
//! | Module | Emits |
//! |--------|-------|
//! | [`workbook_security`] | `<workbookProtection>` + `<fileSharing>` (RFC-058) |
//! | [`sheet_setup`] | `<sheetView>` / `<sheetProtection>` / `<pageMargins>` / `<pageSetup>` / `<headerFooter>` (RFC-055) |

pub mod sheet_setup;
pub mod workbook_security;
