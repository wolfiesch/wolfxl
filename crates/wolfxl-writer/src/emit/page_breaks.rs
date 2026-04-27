//! Sprint Π Pod Π-α (RFC-062) — thin re-exports of the
//! `<rowBreaks>` / `<colBreaks>` emitters from
//! [`crate::parse::page_breaks`]. The native writer's
//! `sheet_xml::emit` calls these at slot 24 (rowBreaks) /
//! slot 25 (colBreaks) per ECMA-376 §18.3.1.99.

pub use crate::parse::page_breaks::{emit_col_breaks, emit_row_breaks, BreakSpec, PageBreakList};
