//! Generic insertion/replacement of sibling blocks inside a `CT_Worksheet`
//! XML stream, preserving ECMA-376 §18.3.1.99 child-element order and passing
//! unknown elements (`extLst`, `x14ac`, future MS extensions, third-party
//! compat tags) through verbatim.
//!
//! See `Plans/rfcs/011-xml-block-merger.md` for the full design rationale.
//! Module-level invariants are pinned in commit 6 of the RFC-011 slice.

// quick_xml types are imported in commit 2 when the streaming algorithm
// lands. The commit-1 stub does not parse XML.

// ---------------------------------------------------------------------------
// SheetBlock — one sibling-block insertion.
// ---------------------------------------------------------------------------

/// One sibling-block insertion. The bytes are pre-serialized including the
/// wrapping element (e.g. `b"<hyperlinks><hyperlink ref=\"A1\" .../></hyperlinks>"`).
/// They MUST be UTF-8 and MUST be a valid XML fragment with one root element
/// at the top — the merger does NOT validate this; malformed input produces
/// malformed output.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SheetBlock {
    /// `<mergeCells count="…">…</mergeCells>` — slot 15 in §18.3.1.99.
    MergeCells(Vec<u8>),
    /// `<conditionalFormatting sqref="…">…</conditionalFormatting>` (one per
    /// range; supply multiple `SheetBlock::ConditionalFormatting` to insert
    /// several). Slot 17.
    ///
    /// **Replace-all semantics:** if any `SheetBlock::ConditionalFormatting`
    /// is present in the `blocks` argument to [`merge_blocks`], every
    /// existing `<conditionalFormatting>` element in the source is removed
    /// before the supplied blocks are inserted. Callers that want to
    /// preserve existing CF rules MUST read them out of the source first
    /// and re-include them in the supplied set. See RFC-011 §5.5.
    ConditionalFormatting(Vec<u8>),
    /// `<dataValidations count="…">…</dataValidations>` — slot 18.
    DataValidations(Vec<u8>),
    /// `<hyperlinks>…</hyperlinks>` — slot 19.
    Hyperlinks(Vec<u8>),
    /// `<legacyDrawing r:id="…"/>` — slot 31. Empty-element form is canonical.
    LegacyDrawing(Vec<u8>),
    /// `<tableParts count="…">…</tableParts>` — slot 37.
    TableParts(Vec<u8>),
}

impl SheetBlock {
    /// The ECMA §18.3.1.99 ordinal (1..=38). Used to decide where to insert
    /// the block when not already present in the source.
    pub fn ecma_position(&self) -> u32 {
        match self {
            SheetBlock::MergeCells(_) => 15,
            SheetBlock::ConditionalFormatting(_) => 17,
            SheetBlock::DataValidations(_) => 18,
            SheetBlock::Hyperlinks(_) => 19,
            SheetBlock::LegacyDrawing(_) => 31,
            SheetBlock::TableParts(_) => 37,
        }
    }

    /// The local-name of the block's root element. Used to detect existing
    /// blocks for replacement (matched against `BytesStart::local_name()`).
    pub fn root_local_name(&self) -> &'static [u8] {
        match self {
            SheetBlock::MergeCells(_) => b"mergeCells",
            SheetBlock::ConditionalFormatting(_) => b"conditionalFormatting",
            SheetBlock::DataValidations(_) => b"dataValidations",
            SheetBlock::Hyperlinks(_) => b"hyperlinks",
            SheetBlock::LegacyDrawing(_) => b"legacyDrawing",
            SheetBlock::TableParts(_) => b"tableParts",
        }
    }

    /// The pre-serialized payload bytes the merger will emit verbatim at the
    /// chosen insertion point.
    pub fn bytes(&self) -> &[u8] {
        match self {
            SheetBlock::MergeCells(b)
            | SheetBlock::ConditionalFormatting(b)
            | SheetBlock::DataValidations(b)
            | SheetBlock::Hyperlinks(b)
            | SheetBlock::LegacyDrawing(b)
            | SheetBlock::TableParts(b) => b,
        }
    }
}

// ---------------------------------------------------------------------------
// ECMA-376 §18.3.1.99 child-element ordering table.
// ---------------------------------------------------------------------------

/// CT_Worksheet child element order per ECMA-376 Part 1 §18.3.1.99. Single
/// source of truth for both the streaming merger ([`merge_blocks`]) and the
/// native writer (`crates/wolfxl-writer/src/emit/sheet_xml.rs`). Adding a
/// slot for a future MS schema extension means updating one constant; both
/// consumers pick it up.
///
/// The table is exhaustive: every `<xsd:element>` child of CT_Worksheet
/// appears here. Unknown elements (extLst-internal extensions, x14ac compat
/// attributes, third-party tags) are NOT in this table and pass through
/// verbatim — they are not part of the §18.3.1.99 sequence.
pub mod ct_worksheet_order {
    /// `(local_name, ordinal)` pairs in ascending order. Lookup is a linear
    /// scan over 38 entries — slower than a `phf::Map` by a constant factor,
    /// but avoids the dependency and is fast enough for once-per-sheet use.
    pub const ECMA_ORDER: &[(&[u8], u32)] = &[
        (b"sheetPr", 1),
        (b"dimension", 2),
        (b"sheetViews", 3),
        (b"sheetFormatPr", 4),
        (b"cols", 5),
        (b"sheetData", 6),
        (b"sheetCalcPr", 7),
        (b"sheetProtection", 8),
        (b"protectedRanges", 9),
        (b"scenarios", 10),
        (b"autoFilter", 11),
        (b"sortState", 12),
        (b"dataConsolidate", 13),
        (b"customSheetViews", 14),
        (b"mergeCells", 15),
        (b"phoneticPr", 16),
        (b"conditionalFormatting", 17),
        (b"dataValidations", 18),
        (b"hyperlinks", 19),
        (b"printOptions", 20),
        (b"pageMargins", 21),
        (b"pageSetup", 22),
        (b"headerFooter", 23),
        (b"rowBreaks", 24),
        (b"colBreaks", 25),
        (b"customProperties", 26),
        (b"cellWatches", 27),
        (b"ignoredErrors", 28),
        (b"smartTags", 29),
        (b"drawing", 30),
        (b"legacyDrawing", 31),
        (b"legacyDrawingHF", 32),
        (b"picture", 33),
        (b"oleObjects", 34),
        (b"controls", 35),
        (b"webPublishItems", 36),
        (b"tableParts", 37),
        (b"extLst", 38),
    ];

    /// Look up the §18.3.1.99 ordinal for a local-name. `None` for unknown
    /// elements (extensions, third-party tags, etc.) — those flow through
    /// the merger verbatim at their source position.
    pub fn ordinal_of(local_name: &[u8]) -> Option<u32> {
        // Linear scan — 38 entries, called at most once per source element.
        // Compared first by length to short-circuit before byte-equality.
        ECMA_ORDER.iter().find_map(|(name, ord)| {
            if *name == local_name {
                Some(*ord)
            } else {
                None
            }
        })
    }
}

// ---------------------------------------------------------------------------
// merge_blocks — public API.
// ---------------------------------------------------------------------------

/// Merge a list of sibling blocks into a worksheet XML stream.
///
/// - If a block of the same root-element-name already exists in `sheet_xml`,
///   the existing block is replaced by the supplied bytes (semantics: "set
///   the entire block to this").
/// - If not present, the block is inserted at the position dictated by
///   ECMA §18.3.1.99 — strictly after every present sibling with a lower
///   ordinal and strictly before every present sibling with a higher ordinal.
/// - Unknown elements (extensions, x14ac, Microsoft-future, etc.) flow
///   through verbatim. The merger never re-serializes attributes or
///   reorders attribute lists on elements it does not own.
/// - `<conditionalFormatting>` is special: multiple
///   `SheetBlock::ConditionalFormatting` in `blocks` produce multiple
///   sibling elements in the output, and existing CF blocks in the source
///   are **all** removed before insertion. See RFC-011 §5.5.
///
/// Errors: `Err(...)` on malformed input XML (delegated from quick_xml).
/// Empty `blocks` returns `sheet_xml` unchanged with no allocation beyond
/// the `Vec::from(...)` for the return type.
pub fn merge_blocks(sheet_xml: &[u8], blocks: Vec<SheetBlock>) -> Result<Vec<u8>, String> {
    if blocks.is_empty() {
        // Empty-blocks fast path: byte-identical no-op. RFC-011 §5.6
        // idempotency contract.
        return Ok(sheet_xml.to_vec());
    }

    // The full streaming algorithm lands in commit 2. Until then, signal
    // that the caller is asking for behavior we don't yet implement so
    // mis-wired callers fail loudly rather than silently dropping blocks.
    Err(
        "wolfxl-merger: merge_blocks streaming algorithm not yet implemented (RFC-011 commit 2)"
            .to_string(),
    )
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::ct_worksheet_order::{ordinal_of, ECMA_ORDER};
    use super::*;

    #[test]
    fn ecma_position_for_each_variant() {
        assert_eq!(SheetBlock::MergeCells(vec![]).ecma_position(), 15);
        assert_eq!(
            SheetBlock::ConditionalFormatting(vec![]).ecma_position(),
            17
        );
        assert_eq!(SheetBlock::DataValidations(vec![]).ecma_position(), 18);
        assert_eq!(SheetBlock::Hyperlinks(vec![]).ecma_position(), 19);
        assert_eq!(SheetBlock::LegacyDrawing(vec![]).ecma_position(), 31);
        assert_eq!(SheetBlock::TableParts(vec![]).ecma_position(), 37);
    }

    #[test]
    fn root_local_name_for_each_variant() {
        assert_eq!(SheetBlock::MergeCells(vec![]).root_local_name(), b"mergeCells");
        assert_eq!(
            SheetBlock::ConditionalFormatting(vec![]).root_local_name(),
            b"conditionalFormatting"
        );
        assert_eq!(
            SheetBlock::DataValidations(vec![]).root_local_name(),
            b"dataValidations"
        );
        assert_eq!(SheetBlock::Hyperlinks(vec![]).root_local_name(), b"hyperlinks");
        assert_eq!(
            SheetBlock::LegacyDrawing(vec![]).root_local_name(),
            b"legacyDrawing"
        );
        assert_eq!(SheetBlock::TableParts(vec![]).root_local_name(), b"tableParts");
    }

    #[test]
    fn ecma_order_table_is_complete_38_slots() {
        // ECMA-376 §18.3.1.99 declares exactly 38 children of CT_Worksheet.
        // If this assertion ever fails, either the spec was extended or we
        // accidentally dropped/duplicated a slot — both are correctness bugs.
        assert_eq!(ECMA_ORDER.len(), 38);

        // Ordinals are strictly increasing 1..=38.
        for (i, (_, ord)) in ECMA_ORDER.iter().enumerate() {
            assert_eq!(*ord, (i as u32) + 1, "slot {} has wrong ordinal", i);
        }

        // No duplicate local-names.
        let mut names: Vec<&[u8]> = ECMA_ORDER.iter().map(|(n, _)| *n).collect();
        names.sort();
        let len = names.len();
        names.dedup();
        assert_eq!(names.len(), len, "ECMA_ORDER has duplicate local-names");
    }

    #[test]
    fn ecma_order_lookup_finds_each_known_local_name() {
        // For each of the 6 SheetBlock variants, the lookup returns the
        // same ordinal as `ecma_position` — i.e. the block helpers and the
        // ordering table cannot drift independently.
        let variants = [
            SheetBlock::MergeCells(vec![]),
            SheetBlock::ConditionalFormatting(vec![]),
            SheetBlock::DataValidations(vec![]),
            SheetBlock::Hyperlinks(vec![]),
            SheetBlock::LegacyDrawing(vec![]),
            SheetBlock::TableParts(vec![]),
        ];
        for v in &variants {
            assert_eq!(
                ordinal_of(v.root_local_name()),
                Some(v.ecma_position()),
                "lookup mismatch for {:?}",
                v.root_local_name()
            );
        }
        // Unknown name returns None.
        assert_eq!(ordinal_of(b"x14ac:unknown"), None);
        assert_eq!(ordinal_of(b""), None);
    }

    #[test]
    fn merge_empty_blocks_is_noop() {
        // The empty-blocks fast path is the idempotency contract from
        // RFC-011 §5.6. Any input bytes — even malformed XML — pass through
        // because we never invoke the parser.
        let xml = b"<worksheet><sheetData/></worksheet>";
        let out = merge_blocks(xml, vec![]).expect("empty blocks must succeed");
        assert_eq!(out, xml);

        let garbage = b"not actually xml at all <<>>";
        let out = merge_blocks(garbage, vec![]).expect("empty blocks bypass parser");
        assert_eq!(out, garbage);
    }
}
