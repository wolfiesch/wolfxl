# RFC-011: Generic XML-Block Insertion Primitive for Sheet XML

Status: Researched
Owner: pod-P1
Phase: 2
Estimate: M
Depends-on: RFC-001
Unblocks: RFC-022, RFC-024, RFC-025, RFC-026

## 1. Problem Statement

`src/wolfxl/sheet_patcher.rs:patch_worksheet` (658 lines) handles **one
thing**: rewriting cell contents inside `<sheetData>`. Every other
modify-mode feature Phase 3 demands needs to insert or replace a *sibling
block* of `<sheetData>`:

| Feature (RFC) | Sibling block to insert |
|---|---|
| Hyperlinks (RFC-022) | `<hyperlinks>…</hyperlinks>` |
| Tables (RFC-024) | `<tableParts>…</tableParts>` |
| Data validations (RFC-025) | `<dataValidations>…</dataValidations>` |
| Conditional formatting (RFC-026) | `<conditionalFormatting>…</conditionalFormatting>` (one per range; multiple allowed) |

ECMA-376 makes the **child element order** of `CT_Worksheet` strictly
mandatory (§18.3.1.99 — full list in §2). Excel rejects out-of-order
sheet XML with a "Repaired" dialog that quarantines the offending
content. So an "insert at the right place" primitive is required, not
optional. Today the patcher would have to either:

- Roll its own streaming insertion logic per RFC (4× duplication, drift
  risk on the ECMA order list), or
- Do a full DOM rebuild (the openpyxl approach — slow, breaks our
  modify-mode "raw-copy unchanged" invariant from `src/wolfxl/mod.rs:1-9`).

Neither is acceptable. RFC-011 is the shared primitive.

User code that depends on this RFC landing:

```python
import wolfxl
wb = wolfxl.load_workbook("report.xlsx", modify=True)
wb["Sheet1"].add_data_validation(dv)        # RFC-025; raises today
wb["Sheet1"].conditional_format(...)         # RFC-026; raises today
wb.save("out.xlsx")
```

After RFC-025/026 land, those calls produce pre-serialized block bytes
that this RFC then merges into the sheet XML at the correct offset,
preserving all other elements (including ones we don't recognize, like
`x14ac:extLst` extensions and `<sortState>`).

## 2. OOXML Spec Surface

ECMA-376 Part 1 §18.3.1.99 — Complex Type `CT_Worksheet`. The schema is
a `<xsd:sequence>` (ordered, not `<xsd:all>`), so element order is
**mandatory**.

The full child element order (verbatim from the spec, in the order they
must appear when present):

1.  `<sheetPr>`
2.  `<dimension>`
3.  `<sheetViews>`
4.  `<sheetFormatPr>`
5.  `<cols>`
6.  `<sheetData>`
7.  `<sheetCalcPr>`
8.  `<sheetProtection>`
9.  `<protectedRanges>`
10. `<scenarios>`
11. `<autoFilter>`
12. `<sortState>`
13. `<dataConsolidate>`
14. `<customSheetViews>`
15. `<mergeCells>`
16. `<phoneticPr>`
17. `<conditionalFormatting>` (0..N — one per CF rule range)
18. `<dataValidations>`
19. `<hyperlinks>`
20. `<printOptions>`
21. `<pageMargins>`
22. `<pageSetup>`
23. `<headerFooter>`
24. `<rowBreaks>`
25. `<colBreaks>`
26. `<customProperties>`
27. `<cellWatches>`
28. `<ignoredErrors>`
29. `<smartTags>`
30. `<drawing>`
31. `<legacyDrawing>`
32. `<legacyDrawingHF>`
33. `<picture>`
34. `<oleObjects>`
35. `<controls>`
36. `<webPublishItems>`
37. `<tableParts>`
38. `<extLst>`

**Namespaces:**

- Default: `http://schemas.openxmlformats.org/spreadsheetml/2006/main`
  (already emitted by `crates/wolfxl-writer/src/emit/sheet_xml.rs:53`).
- `r:` prefix: `http://schemas.openxmlformats.org/officeDocument/2006/relationships`
  (used in `<hyperlink r:id="rId1"/>`, `<tablePart r:id="rId2"/>`,
  `<legacyDrawing r:id="rId3"/>`).
- `mc:`, `x14ac:`, `xr*:` — Microsoft extensions (Mark Compatibility,
  Excel 2010 Authoring Compatibility, etc.). These appear inside `<extLst>`
  and as compatibility attributes on the root. **The merger never inspects
  these; they are passed through verbatim.**

**Relationship-type URIs that this RFC indirectly cares about** (because
the inserted blocks reference rIds the merger doesn't allocate, but must
not corrupt):

- `.../hyperlink` — for `<hyperlink r:id="…"/>` inside `<hyperlinks>`.
- `.../table` — for `<tablePart r:id="…"/>` inside `<tableParts>`.
- `.../vmlDrawing` — for `<legacyDrawing r:id="…"/>` (not a sibling
  block per se, but RFC-023 emits it via this same path).

The merger's contract: rId attributes inside the block payload are the
caller's responsibility. The merger guarantees that whatever rIds the
caller allocated via RFC-010 land in the output verbatim.

**Required vs optional fields**: every element in the list above is
optional (0..1 occurrence except `<conditionalFormatting>` which is 0..N).
A worksheet with only `<dimension>`, `<sheetViews>`, `<sheetFormatPr>`,
and `<sheetData>` is fully valid (and is in fact what
`tests/fixtures/minimal.xlsx` produces).

## 3. openpyxl Reference

File: `.venv/lib/python3.14/site-packages/openpyxl/worksheet/worksheet.py`
(load) and `.../writer/worksheet.py` (save).

Algorithm summary (openpyxl save path):

1. `Worksheet` is a Python object holding fully-deserialized state for
   every CT_Worksheet child element: `merged_cells`, `data_validations`,
   `conditional_formatting`, `auto_filter`, `protection`, etc. All loaded
   from XML via descriptor-driven serialization on workbook open.
2. On save, `openpyxl.writer.worksheet.write_worksheet(ws)` walks those
   attributes in a hardcoded order (matching ECMA §18.3.1.99) and serializes
   each one, then concatenates and writes to the part.
3. Element order is enforced by **the order of the if-blocks in the
   write function**. There is no schema-driven ordering.

This is the wrong shape for us. Three reasons:

1. **Speed.** A worksheet with 100k cells but only 2 hyperlinks should
   take O(2) work to add a hyperlink, not O(100k) to deserialize +
   re-serialize the whole sheet.
2. **Round-trip fidelity.** openpyxl drops anything it doesn't have a
   descriptor for. `x14ac:extLst` cells with extension data
   (`http://schemas.microsoft.com/office/spreadsheetml/2009/9/main`),
   conditional-format extensions for icon sets, sparklines, etc. are
   silently lost. The patcher promises modify mode preserves these
   (CLAUDE.md "Modify mode preserves charts, macros, images, pivots,
   VBA on round-trip").
3. **No re-encoding.** openpyxl's serializer normalizes attribute order,
   indentation, and entity escapes — every save is a large byte diff.
   Modify mode promises minimal diffs.

**Streaming, insertion-only, opposite of openpyxl.** We do exactly one
pass through the source XML, writing every event verbatim except where
we either (a) skip a block we're going to replace, or (b) inject a new
block that wasn't in the source. Unknown elements (extensions, future
ECMA additions, undocumented Microsoft compatibility tags) flow through
untouched.

What we will NOT copy from openpyxl:

- The descriptor-driven model classes for each CT_Worksheet child.
  Block payloads come in as opaque bytes from RFC-022/024/025/026.
- The "rebuild from scratch" save semantics.
- The implicit drop of unknown extension content.

## 4. WolfXL Surface Area

### 4.1 Python coordinator

No new public Python API in this RFC. The merger is plumbing.

The four downstream RFCs add public methods that internally feed
pre-serialized bytes to this merger:

- `Worksheet.add_hyperlink(...)` (RFC-022) — emits `<hyperlinks>` block bytes
- `Worksheet.add_data_validation(...)` (RFC-025) — emits `<dataValidations>`
- `Worksheet.conditional_format(...)` (RFC-026) — emits one or more
  `<conditionalFormatting>` blocks
- `Worksheet.add_table(...)` (RFC-024) — emits `<tableParts>`

### 4.2 Patcher (modify mode) — new module

New file: `src/wolfxl/xml_block_merger.rs` (estimated ~450 LOC).

Public API:

```rust
//! Generic insertion/replacement of sibling blocks inside a CT_Worksheet
//! XML stream, preserving ECMA-376 §18.3.1.99 child-element order and
//! passing unknown elements through verbatim.

/// One sibling-block insertion. The bytes are pre-serialized including
/// the wrapping element (e.g. `b"<hyperlinks><hyperlink ref=\"A1\" .../></hyperlinks>"`).
/// They MUST be UTF-8 and MUST be a valid XML fragment with one root element
/// at the top — the merger does NOT validate this; malformed input produces
/// malformed output.
pub enum SheetBlock {
    /// `<mergeCells count="…">…</mergeCells>` — slot 15 in §18.3.1.99
    MergeCells(Vec<u8>),
    /// `<conditionalFormatting sqref="…">…</conditionalFormatting>` (one per
    /// range; supply multiple `SheetBlock::ConditionalFormatting` to insert
    /// several). Slot 17.
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
    /// The ECMA §18.3.1.99 ordinal (1..38). Used to decide where to insert
    /// the block when not already present in the source.
    pub fn ecma_position(&self) -> u32;
    /// The local-name of the block's root element. Used to detect existing
    /// blocks for replacement.
    pub fn root_local_name(&self) -> &'static [u8];
}

/// Merge a list of sibling blocks into a worksheet XML stream.
///
/// - If a block of the same root-element-name already exists in `sheet_xml`,
///   the existing block is replaced by the supplied bytes (semantics:
///   "set the entire block to this").
/// - If not present, the block is inserted at the position dictated by
///   ECMA §18.3.1.99 — strictly after every present sibling with a lower
///   ordinal and strictly before every present sibling with a higher ordinal.
/// - Unknown elements (extensions, x14ac, Microsoft-future, etc.) flow
///   through verbatim. The merger never re-serializes attributes or
///   reorders attribute lists on elements it does not own.
/// - `<conditionalFormatting>` is special: multiple `SheetBlock::ConditionalFormatting`
///   in `blocks` produce multiple sibling elements in the output. Existing
///   `<conditionalFormatting>` elements in the source are **all** removed
///   before insertion (semantics: "the supplied list IS the new full set").
///   See §5.5 for rationale.
///
/// Errors:
/// - `Err(...)` on malformed input XML (delegated from quick_xml).
/// - Empty `blocks` returns `sheet_xml` unchanged (no allocation).
pub fn merge_blocks(sheet_xml: &[u8], blocks: Vec<SheetBlock>) -> Result<Vec<u8>, String>;
```

ZIP parts read/mutated/emitted: this RFC only operates on
`xl/worksheets/sheet{N}.xml` parts. It is wired into the patcher's
`do_save` (`src/wolfxl/mod.rs:215-362`) at phase 3 (the sheet-XML patch
loop, currently `mod.rs:294-302`) **after** `sheet_patcher::patch_worksheet`
runs. Pseudocode for the new flow:

```rust
for (sheet_path, patches) in &sheet_cell_patches {
    let xml = ooxml_util::zip_read_to_string(&mut zip, sheet_path)?;

    // (existing) cell-level patches
    let after_cells = sheet_patcher::patch_worksheet(&xml, patches)?;

    // (new) sibling-block insertions queued by RFC-022/024/025/026
    let blocks = self.queued_blocks.remove(sheet_path).unwrap_or_default();
    let after_blocks = if blocks.is_empty() {
        after_cells.into_bytes()
    } else {
        xml_block_merger::merge_blocks(after_cells.as_bytes(), blocks)?
    };

    file_patches.insert(sheet_path.clone(), after_blocks);
}
```

`XlsxPatcher` gains a new field:

```rust
queued_blocks: HashMap<String /* sheet_path */, Vec<SheetBlock>>,
```

populated by future-RFC methods like `queue_hyperlink`, `queue_table`,
etc. These methods are out of scope for RFC-011 itself.

### 4.3 Native writer (write mode)

The native writer at `crates/wolfxl-writer/src/emit/sheet_xml.rs`
already builds CT_Worksheet in correct ECMA order from scratch (see
the comments at `sheet_xml.rs:55-99` and the EXT-W3A/W3B/W3C marker
comments at lines 77-95). It does **not** need this merger — it has
full model state and emits everything in one pass.

However, the merger's `SheetBlock::ecma_position` table is the
authoritative source of the ordering, and the writer should consume
it. Refactor:

- Extract a `pub mod ct_worksheet_order` from
  `src/wolfxl/xml_block_merger.rs` containing the 38-element ordinal
  table as `pub const`s.
- `crates/wolfxl-writer/src/emit/sheet_xml.rs::emit` uses those
  constants to label its own emit-section comments (purely
  documentary; no behavior change).

This means future ECMA additions (slot 39 if Microsoft extends the
schema) require updating one table, and both backends pick it up.

## 5. Algorithm

### 5.1 Single-pass streaming with `quick_xml::Reader`

State:

```text
position_cursor: u32              // last-emitted element's ordinal in §18.3.1.99
pending_blocks: BTreeMap<u32, SheetBlock>     // by ecma_position
replace_set:    HashSet<&'static [u8]>        // root names to replace if seen
cf_replace:     bool                          // special case for §5.5
```

Init:

```text
for block in blocks:
    pending_blocks.insert(block.ecma_position(), block);
    replace_set.insert(block.root_local_name());
    if block.is::<ConditionalFormatting>(): cf_replace = true;
```

Main loop (sketch):

```text
loop reader.read_event():
    Event::Start(e) | Event::Empty(e):
        let local = e.local_name();
        let ord = ECMA_ORDER.get(local);   // None if unknown / extension

        if cf_replace and local == b"conditionalFormatting":
            // skip this whole element (and its children)
            consume_until_matching_end(reader, local);
            continue;

        if let Some(ord) = ord:
            // Flush every pending block that comes BEFORE this element.
            while let Some(b) = pending_blocks.first_entry()
                  if b.key() < ord:
                writer.write_raw(pending_blocks.pop_first().bytes());

            if replace_set.contains(local):
                // We have a replacement — drop the source block.
                consume_until_matching_end(reader, local);
                writer.write_raw(pending_blocks.remove(ord).bytes());
                position_cursor = ord;
                continue;

            // Not a replacement — emit the source element verbatim.
            writer.write_event(e);
            position_cursor = ord;
        else:
            // Unknown element (extension or compat). Pass through.
            writer.write_event(e);

    Event::End(e):
        if e.local_name() == b"worksheet":
            // Flush any remaining blocks before closing the root.
            while let Some(b) = pending_blocks.pop_first():
                writer.write_raw(b.bytes());
        writer.write_event(e);

    Event::Eof: break;
    other (Text, CData, Decl, PI, Comment, DocType): writer.write_event;
```

### 5.2 Why streaming and not DOM

A worksheet XML can be 50 MB+ (large datasets). Building a DOM means
allocating 50 MB+ in addition to the input buffer, then walking it.
Streaming is O(input bytes) memory-bounded by the size of the largest
element, which for sheet XML is typically a `<row>` (a few KB).

The existing `sheet_patcher::patch_worksheet` already streams (see
`src/wolfxl/sheet_patcher.rs:75-245`); `xml_block_merger` matches that
pattern and uses the same `quick_xml` types so the two passes can be
fused later if profiling shows the second pass is hot.

### 5.3 Verbatim pass-through of unknown elements

`quick_xml::Reader` emits each element as one or more `Event::Start`,
`Event::Text/CData`, `Event::End` triples. Writing them back via
`Writer::write_event` preserves attribute order, namespace declarations,
self-closing form, and even most whitespace.

**It does not preserve byte-identical attribute *value* encoding.** If
the source has `Target="path?a=1&amp;b=2"` and quick_xml's
`unescape_value` decodes-then-re-encodes it, a `&amp;` could come out as
`&amp;` consistently (good) or could come out as `&#38;amp;` (bad).
quick_xml documents this: the `Reader::trim_text(false)` we use plus
`Writer::write_event` with the same `BytesStart` re-emits attributes
**from the raw byte slice**, not from a decoded form. That preserves
encoding. We must verify this with a test (see §6 test #5).

### 5.4 ECMA position table

```rust
pub const ECMA_ORDER: &[(&[u8], u32)] = &[
    (b"sheetPr",              1),
    (b"dimension",            2),
    (b"sheetViews",           3),
    (b"sheetFormatPr",        4),
    (b"cols",                 5),
    (b"sheetData",            6),
    (b"sheetCalcPr",          7),
    (b"sheetProtection",      8),
    (b"protectedRanges",      9),
    (b"scenarios",           10),
    (b"autoFilter",          11),
    (b"sortState",           12),
    (b"dataConsolidate",     13),
    (b"customSheetViews",    14),
    (b"mergeCells",          15),
    (b"phoneticPr",          16),
    (b"conditionalFormatting", 17),
    (b"dataValidations",     18),
    (b"hyperlinks",          19),
    (b"printOptions",        20),
    (b"pageMargins",         21),
    (b"pageSetup",           22),
    (b"headerFooter",        23),
    (b"rowBreaks",           24),
    (b"colBreaks",           25),
    (b"customProperties",    26),
    (b"cellWatches",         27),
    (b"ignoredErrors",       28),
    (b"smartTags",           29),
    (b"drawing",             30),
    (b"legacyDrawing",       31),
    (b"legacyDrawingHF",     32),
    (b"picture",             33),
    (b"oleObjects",          34),
    (b"controls",            35),
    (b"webPublishItems",     36),
    (b"tableParts",          37),
    (b"extLst",              38),
];
```

Lookup is via a `phf::Map` (compile-time perfect hash) for O(1) match
without a dependency on a runtime hashmap. If we'd rather avoid the
`phf` crate, a linear scan of 38 entries is also fast enough — we'd
match on the first byte before doing the full comparison.

### 5.5 Why `<conditionalFormatting>` is special

Slot 17 allows 0..N elements (one per CF range), unlike every other
slot which is 0..1. Three options for replacement semantics:

- **A. Append:** new CF blocks go after existing ones. Wrong — caller
  cannot remove an existing CF rule via this API.
- **B. Per-range merge:** parse `sqref` of each existing CF block,
  match against caller's blocks, replace per range. Too smart — pulls
  CF parsing into the merger, defeats the "opaque bytes" contract.
- **C. Replace-all:** if any `SheetBlock::ConditionalFormatting` is
  in `blocks`, delete every existing `<conditionalFormatting>` and
  insert all caller blocks. **Chosen.** Forces RFC-026 to read existing
  CF blocks first (if it wants to preserve them), pass the full
  desired set in. Keeps merger's contract clean: opaque bytes, single
  ordinal slot.

This is documented loud and clear on `SheetBlock::ConditionalFormatting`
docs.

### 5.6 Idempotency and determinism

`merge_blocks(merge_blocks(xml, []), []) == merge_blocks(xml, [])` —
empty `blocks` is a no-op, returns input bytes unchanged.

`merge_blocks(xml, blocks)` is deterministic given identical inputs —
`BTreeMap` ordering of `pending_blocks` ensures CF blocks land in
ascending sqref order if the caller sorted, otherwise in insertion
order. We document: "callers SHOULD pre-sort `SheetBlock::ConditionalFormatting`
blocks by sqref for deterministic output across save cycles."

`WOLFXL_TEST_EPOCH=0` is honored (no time-dependent output anywhere
in the merger).

## 6. Test Plan

Standard verification matrix from master plan §Verification, plus:

**Unit tests in `src/wolfxl/xml_block_merger.rs`:**

1. `merge_empty_blocks_is_noop` — `merge_blocks(xml, vec![])` returns
   `xml` byte-identical.
2. `insert_hyperlinks_into_minimal_sheet` — input has only
   `<dimension>`, `<sheetViews>`, `<sheetFormatPr>`, `<sheetData>`,
   `<pageMargins>`. Insert `Hyperlinks(...)`. Output has hyperlinks
   between `<pageMargins>` predecessor (none — but logically between
   `<sheetData>` and `<pageMargins>`). Specifically: hyperlinks lands
   immediately before `<pageMargins>` (slot 21).
3. `replace_existing_hyperlinks` — input already has
   `<hyperlinks><hyperlink ref="A1" .../></hyperlinks>`. Insert new
   `Hyperlinks(...)` block. Output has the new block, not the old.
4. `insert_into_correct_ecma_position` — for each of the 6
   `SheetBlock` variants, build a sheet with `<sheetData>` plus
   one earlier-slot element and one later-slot element, insert the
   block, assert byte-position ordering of opening tags.
5. **`extlst_is_byte_preserved` (RFC §8 risk #1).** Construct a
   sheet with both `<extLst>` extensions on cells (`x14ac` rows) and
   a top-level `<extLst>` block. Insert a `Hyperlinks(...)`. Diff
   the input `<extLst>` bytes against the output `<extLst>` bytes —
   must be byte-identical (no attribute-order shuffling, no entity
   re-encoding, no namespace prefix change). This is the headline
   correctness test.
6. `unknown_element_passthrough` — invent a fake element
   `<wolfxl:custom xmlns:wolfxl="urn:test"/>` between `<sheetData>`
   and `<pageMargins>` in the input. After merging blocks, it appears
   in the same relative position with same attributes.
7. `multiple_conditionalformatting_blocks` — supply 3 CF blocks,
   assert all 3 appear in output in supplied order, all in slot 17
   contiguously.
8. `conditionalformatting_replaces_all_existing` — input has 2 CF
   blocks; supply 1; output has 1 (the supplied one).
9. `block_inserted_when_no_neighbors` — sheet has only `<sheetData>`.
   Insert `TableParts(...)`. Output has `<sheetData>` then
   `<tableParts>` then `</worksheet>`.
10. `tableparts_after_extlst_is_wrong_and_we_fix_it` — input
    pathologically has `<extLst>` before `<tableParts>` (some
    third-party libs do this). Source order is preserved on
    pass-through, BUT inserted blocks land at correct ECMA slot.
    Assert: the source extLst stays where source put it; new
    `<tableParts>` inserted lands before extLst (slot 37 < 38).
11. `large_sheet_streaming_memory_bounded` — 50 MB sheet XML
    (synthetic, 1M `<row>` entries). Merger memory peak < 4 MB.
12. `byte_identical_when_block_already_present_and_unchanged` —
    if the supplied bytes for a block exactly match what's in the
    source, output is byte-identical to input. Lets us cheaply detect
    no-op saves.

**Integration tests:**

13. `tests/integration_merger_round_trip.rs` (new) — for each
    `tests/fixtures/tier2/*.xlsx` worksheet, parse with the patcher,
    merge an empty block list, assert byte-identical to input. Covers
    pathological real-world XML.
14. `tests/integration_hyperlink_end_to_end.rs` (RFC-022 test, listed
    here because the merger is half the integration) — open
    `tier2/13_hyperlinks.xlsx`, add a new hyperlink, save, reopen with
    openpyxl, assert openpyxl can read both old and new hyperlinks.

**Property test:** `proptest` with random sheet XMLs (built from a
small grammar that emits random subsets of CT_Worksheet children in
random order) and random block combinations. Assert: output is valid
XML, output passes the `crates/wolfxl-cli` `schema` command without
"Repaired" warnings, output has every input element preserved or
replaced.

**Cross-surface check:** after RFC-022/024/025/026 land, run
`tests/parity/` against fixture set with hyperlinks/tables/dv/cf.
HARD-tier must pass.

## 7. Migration / Compat Notes

- **No public Python API change.** Pure plumbing.
- **Behavior diff vs openpyxl:** wolfxl preserves unknown extension
  elements; openpyxl drops them. Documented in `KNOWN_GAPS.md` as
  a wolfxl improvement.
- **Feature flag during rollout:** `WOLFXL_DISABLE_BLOCK_MERGER=1`
  env var causes the patcher to skip the `merge_blocks` call entirely,
  reverting to today's "raw-copy non-cell parts" behavior. Useful
  bisecting if RFC-022/024/025/026 land in stages and a regression
  appears. Remove the flag after Phase-3 stabilizes (Phase 4).
- **Backward-compat shim:** none needed. Phase-3 features that depend
  on this RFC raise `NotImplementedError` today (per the
  `_compat._make_stub` framework noted in CLAUDE.md). Once they
  start producing real bytes, the merger handles them.
- **Determinism:** with `WOLFXL_TEST_EPOCH=0`, two consecutive saves
  of the same workbook with the same queued blocks produce
  byte-identical output. Required for golden-file tests.

## 8. Risks & Open Questions

1. **(HIGH) `<extLst>` byte preservation (test #5 above).** If
   quick_xml normalizes any aspect of attribute encoding, namespace
   declaration order, or prefix binding inside extLst children, our
   byte-identity claim breaks. Mitigation: test #5 is the gate. If it
   fails, fall back to a bytewise copy of the source span between
   `<extLst>` open and `</extLst>` close, stitched into the output
   without reparsing. The fallback is straightforward
   (track byte offsets while streaming), just more code.

2. **(MED) Comment and CDATA sections inside CT_Worksheet.** Real
   files occasionally embed XML comments
   (`<!-- generated by foo -->`) at the worksheet level. quick_xml
   delivers these as `Event::Comment`. Our streaming loop forwards
   them. **But:** if a comment sits between two ECMA-ordered elements,
   does it move when we insert a block between them? **Resolution:**
   comments and PIs flow through at the source position; if a block
   is inserted at slot N+1, the comment stays attached to its
   preceding source element. Document this. Add test.

3. **(MED) Self-closing `<worksheet/>` (degenerate input).** No
   fixture has this, but it's syntactically possible. Our state
   machine looks for `Event::End(b"worksheet")` to flush remaining
   blocks; an `Event::Empty(b"worksheet")` would never trigger that.
   **Resolution:** `if let Event::Empty(b"worksheet")` branch flushes
   pending blocks immediately (and necessarily wraps them in an
   explicit `<worksheet>...</worksheet>` instead of self-closing).

4. **(LOW) Namespace prefix mismatch on root.** If `sheet_xml`
   declares `xmlns:r=` on `<worksheet>` but our `Hyperlinks` block
   bytes use `r:id="rId1"` without re-declaring the namespace, the
   output is still valid because the prefix scope inherits down. But
   if the root uses a different prefix
   (`xmlns:rel="...relationships"` instead of `xmlns:r="..."`),
   `r:id` resolves to nothing. **Resolution:** RFC-022/024/025/026
   are responsible for using the prefix bound at the root. RFC-011
   does not rebind. Add a parser-time check: if `<worksheet>` does
   not declare `xmlns:r="...relationships"`, we emit one in the
   output (rewriting the open tag attributes — only place we ever
   modify a non-block element). Test #5 must still pass given this.

5. **(LOW) Performance on 50 MB worksheets.** Streaming should be
   O(input). Profile with test #11. If `Writer::write_event` allocates
   per call (it shouldn't, but `BytesStart::to_owned` does), we
   may need to drop down to direct byte writes for the high-frequency
   `<c>` and `<row>` events that aren't being mutated.

6. **(OPEN) Should the merger validate that block payload bytes
   start with the expected element?** A caller passing
   `SheetBlock::Hyperlinks(b"<wrong/>")` would silently produce
   bad XML. Pro: defense in depth. Con: re-parsing the block bytes
   to check defeats the "opaque bytes" contract speed-wise.
   **Proposed resolution:** debug-assertion only —
   `debug_assert!(block.bytes().starts_with(format!("<{}", root_local_name())))`.
   In release builds, garbage in → garbage out, with the test
   suite as the safety net.

## 9. Effort Breakdown

| Task | LOC est. | Days |
|---|---|---|
| `src/wolfxl/xml_block_merger.rs` — module + state machine + ECMA table | 350 | 1.5 |
| `SheetBlock` enum + variant helpers | 60 | 0.2 |
| Wiring into `XlsxPatcher::do_save` (new field, new merge call) | 40 | 0.3 |
| Unit tests (12 above) | 380 | 0.7 |
| Integration tests (round-trip, hyperlink end-to-end) | 200 | 0.5 |
| Property test (proptest with grammar) | 120 | 0.3 |
| Refactor `wolfxl-writer/src/emit/sheet_xml.rs` to consume `ECMA_ORDER` constants | 50 | 0.2 |
| **Total** | **~1,200** | **~3.7 days** |

Estimate band: M (≤ 3 days) — slightly over; could be downgraded to L
if test #5 fallback (per risk #1) is needed. Realistic call: M with
0.5 day buffer.

## 10. Out of Scope

The following are NOT in this RFC:

- **Parsing the contents of supplied block bytes.** The merger is opaque.
  RFC-022/024/025/026 own block content.
- **Generating the block bytes.** Each downstream RFC builds them.
- **Validating that supplied rIds inside block payloads exist in the
  rels graph.** That's the caller's responsibility. RFC-010 +
  RFC-022/024 cooperate.
- **`<sheetData>` mutation.** Already handled by
  `src/wolfxl/sheet_patcher.rs`. The merger NEVER touches `<sheetData>`
  contents, only its sibling blocks.
- **Pretty-printing or canonicalizing existing source XML.** Pass-through
  preserves whatever the source had.
- **Multiple `<extLst>` elements.** Spec allows at most one (slot 38).
  We assume one or zero. If a malformed source has two, we pass them
  both through untouched (degenerate case).
- **Re-emitting blocks as anything other than UTF-8.** ECMA-376 mandates
  UTF-8 for OOXML; we never see UTF-16 or other encodings.
- **Concurrent `merge_blocks` calls on the same XML.** Single-threaded.
- **Inline mutation of source elements** (e.g., adding `xr:uid="..."` to
  every `<row>`). Not a sibling-block insertion; explicitly out.

## Acceptance: shipped via 6-commit slice on `feat/native-writer` (through f06290f) on 2026-04-25

The `crates/wolfxl-merger/` crate is live with `SheetBlock`, the 38-slot
`ECMA_ORDER` table, and the streaming `merge_blocks` algorithm. 22 tests
green (20 unit including the headline `extlst_is_byte_preserved`, 1
integration round-trip across all `tests/fixtures/` worksheet XMLs, 1
500-iter property test). Wired into `XlsxPatcher` via `queued_blocks` (dead
code at HEAD) and into the writer via `wolfxl_merger::ct_worksheet_order::
ECMA_ORDER` for slot-number references in `emit/sheet_xml.rs` (with a
const-time pin assertion). Risk #1 (extLst byte-preservation) verified
without falling back to byte-span copying — quick-xml's `Reader`+`Writer`
preserves attribute slices verbatim. RFC-022/024/025/026 can now activate
the merger by populating `XlsxPatcher::queued_blocks`.
