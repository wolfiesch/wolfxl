//! wolfxl-core: pure-Rust spreadsheet primitives for `wolfxl-cli` and other consumers.
//!
//! This crate carries no PyO3 / Python coupling. It opens xlsx workbooks via
//! `calamine-styles`, exposes a sheet/row/cell API tailored for agent-facing
//! tools (peek, map, schema), and renders cell values respecting Excel
//! number formats (currency, percentage, date, etc.).
//!
//! The existing PyO3 layer in `wolfxl` (the cdylib at the workspace root)
//! still owns its own implementation today; unifying the two is tracked as
//! follow-up work.

pub mod cell;
pub mod error;
pub mod format;
pub mod sheet;
pub mod workbook;

pub use cell::{Cell, CellValue};
pub use error::{Error, Result};
pub use format::{FormatCategory, format_cell};
pub use sheet::Sheet;
pub use workbook::Workbook;
