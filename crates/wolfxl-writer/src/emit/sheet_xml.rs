//! `xl/worksheets/sheet{N}.xml` emitter — rows, cells, merges, freeze,
//! columns, print area, and extension hooks for CF/DV. Wave 2B.
//!
//! # rId convention (must match [`crate::emit::rels::emit_sheet`])
//!
//! Sheet-level relationships are allocated in this order inside
//! `xl/worksheets/_rels/sheet{N}.xml.rels`:
//!
//! 1. **Comments** (if any): `rId1` points at `commentsN.xml`,
//!    `rId2` at `vmlDrawingN.vml`.
//! 2. **Tables**: the next contiguous block. With no comments,
//!    tables start at `rId1`; with comments, at `rId3`.
//! 3. **External hyperlinks** (targets that do not start with `#`):
//!    the tail of the rId range.
//!
//! The emitter MUST walk [`Worksheet::hyperlinks`] with the same
//! filter + iteration order as `rels::emit_sheet` uses when assigning
//! `r:id` attributes in `<hyperlink>` elements, or Excel will follow
//! mismatched rIds and silently drop hyperlink targets.
//!
//! # Extension hooks (Wave 3)
//!
//! The emitter leaves `// EXT-W3A:`, `// EXT-W3B:`, and `// EXT-W3C:`
//! marker comments at the three insertion points where Wave 3 agents
//! plug in comments/VML bridging, tables, conditional formats, and
//! data validations. Keep them even when the related collections
//! are empty — Wave 3 may need to emit structural parents.

use crate::intern::SstBuilder;
use crate::model::format::StylesBuilder;
use crate::model::worksheet::Worksheet;

/// Emit `xl/worksheets/sheet{N}.xml` bytes for one sheet.
///
/// `sheet_idx` is zero-based; the caller converts to 1-based for any
/// user-facing references (`sheet1.xml`, `commentsN.xml`, etc.).
///
/// `sst` is mutable because string cells intern at emit time, not model
/// construction time. `styles` is immutable because all interning already
/// happened during `WriteCell` construction.
pub fn emit(
    _sheet: &Worksheet,
    _sheet_idx: u32,
    _sst: &mut SstBuilder,
    _styles: &StylesBuilder,
) -> Vec<u8> {
    Vec::new()
}
