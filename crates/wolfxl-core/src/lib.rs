//! wolfxl-core: pure-Rust xlsx reader with Excel number-format-aware rendering.
//!
//! This crate carries no PyO3 / Python coupling. It opens xlsx workbooks via
//! [`calamine-styles`], exposes a `Workbook → Sheet → Cell` API, and renders
//! cell values with awareness of common Excel number formats (currency,
//! percentage, scientific, date, time).
//!
//! ```no_run
//! use wolfxl_core::Workbook;
//!
//! let mut wb = Workbook::open("examples/sample-financials.xlsx")?;
//! let sheet = wb.first_sheet()?;
//! let (rows, cols) = sheet.dimensions();
//! println!("{} rows × {} columns", rows, cols);
//! # Ok::<_, wolfxl_core::Error>(())
//! ```
//!
//! ## Scope
//!
//! - **In scope today:** read xlsx values + best-effort number-format
//!   strings, classify formats into [`FormatCategory`], render via
//!   [`format_cell`], map workbook structure, infer per-column
//!   schema/cardinality summaries, and walk `xl/styles.xml` cellXfs +
//!   numFmts as a fallback when calamine's fast path returns None
//!   (covers openpyxl-generated fixtures).
//! - **Not yet:** write side.
//!
//! The existing PyO3 layer in the sibling `wolfxl` cdylib still owns its own
//! implementation; unifying the two is follow-up work.

pub mod cell;
pub mod error;
pub mod format;
pub mod map;
pub mod ooxml;
pub mod schema;
pub mod sheet;
pub mod styles;
pub mod workbook;
pub mod worksheet_xml;

pub use cell::{Cell, CellValue};
pub use error::{Error, Result};
pub use format::{format_cell, FormatCategory};
pub use map::{classify_sheet, SheetClass, SheetMap, WorkbookMap};
pub use schema::{infer_sheet_schema, Cardinality, ColumnSchema, InferredType, SheetSchema};
pub use sheet::Sheet;
pub use styles::{builtin_num_fmt, resolve_num_fmt, XfEntry, BUILTIN_NUM_FMTS};
pub use workbook::{Workbook, WorkbookStyles};
