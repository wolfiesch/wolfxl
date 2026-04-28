# RFC-013: Patcher Infrastructure Extensions

Status: Researched
Owner: pod-synthesis (cross-cutting from Wave 1+2)
Phase: 2
Estimate: M
Depends-on: RFC-001
Unblocks: RFC-022, RFC-023, RFC-024, RFC-035 (any RFC that adds new ZIP entries or aggregates cross-sheet state)

## 1. Problem Statement

Wave 1 + Wave 2 research surfaced three cross-cutting patcher gaps that no single feature RFC owns. Each is small individually, but every T1.5 RFC in Phase 3 and every structural-op RFC in Phase 4 trips over at least one. Bundling them into a single Phase-2 infrastructure RFC prevents 7 downstream RFCs from each re-litigating the same plumbing.

The three gaps:

1. **ZIP rewriter is ADD-blind**. Today's `XlsxPatcher::do_save` (`src/wolfxl/mod.rs:343-356`) iterates the source ZIP and either applies a patch from `file_patches` or raw-copies. There is no path to ADD a new entry that didn't exist in the source. RFC-022 (hyperlinks needs new `xl/worksheets/_rels/sheetN.xml.rels` if source had none), RFC-023 (comments need NEW `xl/comments<N>.xml` and `xl/drawings/vmlDrawing<N>.vml` parts), RFC-024 (tables need NEW `xl/tables/tableN.xml`), and RFC-035 (copy_worksheet creates an entirely new sheet part) all require ADD.

2. **Patcher reads no ancillary parts**. Today the patcher tracks cell patches and styles only. Tables, comments, conditional formatting, and data validations are read via the `CalamineStyledBook` reader, not loaded into the patcher's mutable state. RFCs 022/023/024/025/026 each need their own "parse existing block at open()" plumbing; RFCs 030/031/034/035 then need that plumbing to be already populated when they query "what tables are on this sheet?". Without RFC-013, every later RFC duplicates the rels-walk + part-load logic.

3. **Cross-sheet aggregation has no two-phase flush**. `[Content_Types].xml` (which gets new entries when a new comments/table part is added) and `xl/styles.xml`'s `<dxfs>` collection (which grows when conditional-formatting rules add formatting) must be modified workbook-wide, not per-sheet. The current patcher flushes per-sheet without an aggregation phase. RFC-024 (multiple sheets adding tables) and RFC-026 (multiple sheets adding CF rules) both noted this gap independently.

## 2. OOXML Spec Surface

**`[Content_Types].xml`**: ECMA-376 Part 2, Open Packaging Conventions, §10.1. Two element types:
- `<Default Extension="..." ContentType="..."/>` — fallback by file extension
- `<Override PartName="/xl/comments1.xml" ContentType="..."/>` — explicit per-part

Namespace: `http://schemas.openxmlformats.org/package/2006/content-types`.

Adding a new part always requires either adding an Override OR ensuring a Default exists for that extension. Comments and tables always need explicit Override entries (they share `.xml` extension with sheet parts but have different content types).

Content types touched by Phase 3 RFCs:
- `application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml`
- `application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml`
- `application/vnd.openxmlformats-officedocument.vmlDrawing` (Default extension `vml`)

## 3. openpyxl Reference

openpyxl rebuilds the entire package on save (`openpyxl/writer/excel.py::ExcelWriter`), so it doesn't have an analogous problem — every part is freshly serialized. We don't copy this approach because modify-mode's whole point is preserving what we don't touch. We need targeted ADD + targeted aggregation.

`openpyxl/packaging/manifest.py` is the closest analog — the `Manifest` class wraps `[Content_Types].xml` with typed `Override` and `Default` records. Mirror its API but keep ours leaner (no descriptor metaclass).

## 4. WolfXL Surface Area

### 4.1 Python coordinator

No new API. RFC-013 is purely below-the-line plumbing — Python-level changes in this RFC limited to:
- `python/wolfxl/_workbook.py:save()` — call `_flush_workbook_aggregations()` BEFORE `_rust_patcher.save()` (the new method threads content-types and styles deltas across all sheets in one pass)

### 4.2 Patcher (modify mode)

Three changes to `src/wolfxl/mod.rs`:

**Change 1 — `file_adds` map alongside `file_patches`**:
```rust
pub struct XlsxPatcher {
    // existing fields ...
    file_patches: HashMap<String, Vec<u8>>,   // existing
    file_adds: HashMap<String, Vec<u8>>,      // NEW: brand-new entries
    file_deletes: HashSet<String>,            // NEW: explicit removals (for reverse direction)
}
```

`do_save` loop additions (`mod.rs:343-356`):
- After processing source-ZIP entries, iterate `file_adds` and emit each as a new ZIP entry with default `SimpleFileOptions` (compressed, current mtime — or epoch if `WOLFXL_TEST_EPOCH=0`).
- Skip source entries listed in `file_deletes`.

**Change 2 — `walk_rels_at_open()` populating ancillary part registry**:
```rust
pub struct AncillaryPartRegistry {
    /// per sheet: paths to comments/drawings/tables/etc. parts referenced from that sheet's rels
    pub per_sheet: HashMap<String, SheetAncillary>,
}
pub struct SheetAncillary {
    pub comments_part: Option<String>,        // "xl/comments3.xml"
    pub vml_drawing_part: Option<String>,
    pub table_parts: Vec<String>,             // ["xl/tables/table1.xml", ...]
    pub hyperlinks_rels: Vec<RelId>,          // pre-existing hyperlink rIds
    pub legacy_drawing_rid: Option<RelId>,
}
```

Population: `XlsxPatcher::open()` loads `xl/_rels/workbook.xml.rels`, walks each sheet, loads that sheet's `_rels/sheetN.xml.rels` via RFC-010's `RelsGraph`, classifies each relationship by type, populates registry. Lazy: only populate when a structural op or T1.5 mutation is queued for that sheet (cheap when modifying just cell values).

**Change 3 — `ContentTypesGraph` + two-phase flush**:
```rust
pub struct ContentTypesGraph {
    defaults: Vec<(String /*ext*/, String /*type*/)>,
    overrides: Vec<(String /*part*/, String /*type*/)>,
}
impl ContentTypesGraph {
    pub fn parse(xml: &[u8]) -> Result<Self>;
    pub fn add_override(&mut self, part: &str, content_type: &str);
    pub fn ensure_default(&mut self, ext: &str, content_type: &str);
    pub fn serialize(&self) -> Vec<u8>;
}
```

Two-phase flush:
1. Per-sheet flush: each sheet's flush returns a list of `ContentTypeOp` (add Override / ensure Default) and a list of `DxfDelta` (new dxf entries to append to styles.xml).
2. Workbook-level aggregation: collect all per-sheet ops, apply once to a single `ContentTypesGraph` (parsed from source), apply all dxfs to a single styles.xml mutation, write back.

### 4.3 Native writer

No changes. The native writer rebuilds `[Content_Types].xml` and `styles.xml` from scratch every save — it's the patcher that needs the aggregation primitives.

## 5. Algorithm

```
# Save flow with RFC-013 in place:
def save(target_path):
    # Phase A: Per-sheet flush (existing + new)
    sheet_results = []
    for sheet in patcher.dirty_sheets:
        result = sheet.flush()   # returns: (sheet_xml_bytes, [content_type_ops], [dxf_deltas], [file_adds], [rels_bytes])
        sheet_results.append(result)
    
    # Phase B: Aggregate cross-sheet state
    content_types = ContentTypesGraph::parse(source_zip["[Content_Types].xml"])
    for r in sheet_results:
        for op in r.content_type_ops:
            content_types.apply(op)
    
    styles_bytes = source_zip["xl/styles.xml"]
    all_dxfs = sheet_results.flat_map(|r| r.dxf_deltas)
    if all_dxfs:
        styles_bytes = styles::append_dxfs(styles_bytes, all_dxfs)
    
    file_patches[CONTENT_TYPES_PATH] = content_types.serialize()
    if all_dxfs: file_patches["xl/styles.xml"] = styles_bytes
    
    # Phase C: ZIP rewrite (existing patches + adds + skip deletes)
    do_save(target_path)
```

**Idempotency**: `ContentTypesGraph::add_override` is idempotent — adding the same (part, type) twice is a no-op. Same for `ensure_default`. Sheet flush ops can be retried safely.

## 6. Test Plan

| Layer | Test |
|---|---|
| Unit (Rust) | `ContentTypesGraph::parse → add → serialize → re-parse` round-trip; `file_adds` flow with a hand-crafted entry |
| Round-trip golden | Modify a fixture, add a comment (forces new comments part + content-type override) — assert the saved ZIP has both new entries and a valid `[Content_Types].xml` |
| openpyxl parity | Same op via openpyxl, semantic XML diff — expect identical content-types and styles.xml deltas |
| LibreOffice | Add 1 table to 3 different sheets in same `save()` — open in LibreOffice, all 3 tables present (catches missing content-type overrides) |
| Cross-mode | Write-mode (rebuilds-from-scratch) vs modify-mode (RFC-013 aggregation) → byte-identical or semantically equal |
| Regression | `tests/fixtures/013-multi-sheet-aggregation.xlsx` with 3 sheets, each getting 1 added feature in a single save |

## 7. Migration / Compat Notes

No user-visible changes. Existing modify-mode tests keep passing because:
- `file_adds` is empty for all current tests (no current test adds a part)
- Two-phase flush is identity when no sheet contributes content-type ops or dxfs
- `walk_rels_at_open()` is lazy — uninvoked unless a queued op needs it

## 8. Risks & Open Questions

1. **(MED)** `walk_rels_at_open()` adds an open-time cost. Mitigation: lazy — only walk on first ancillary query. Benchmark with a 100-sheet fixture.
2. **(LOW)** `file_adds` collision with existing entry — should be a hard error (caller bug). Add an assertion: `file_adds` ∩ source-ZIP names must be empty.
3. **(MED)** ContentTypesGraph order preservation — Excel cares about ordering for some validators. Preserve source order for unchanged entries; append new at end.
4. **(LOW)** Two-phase flush sequencing in `_workbook.py:save()` — must run BEFORE `_rust_patcher.save()`. Document the ordering contract in a comment.

## 9. Effort Breakdown

| Task | LOC est. | Days |
|---|---|---|
| `file_adds` / `file_deletes` plumbing in `mod.rs::do_save` | ~80 | 0.5 |
| `walk_rels_at_open()` + `AncillaryPartRegistry` | ~250 | 1.5 |
| `ContentTypesGraph` parse/mutate/serialize | ~180 | 1 |
| Two-phase flush refactor in `mod.rs` + `_workbook.py:save()` | ~150 | 1 |
| Tests (per layer) | ~400 | 1.5 |
| **Total** | **~1060** | **5.5 days (M)** |

## 10. Out of Scope

- Removing source-ZIP entries (use case: RFC-035 + delete-sheet, not yet on roadmap). `file_deletes` is reserved for future use; the v1 implementation can leave it unused.
- Concurrent modification (two sheets being mutated from different threads). The patcher is single-threaded by design.
- Incremental content-type updates streamed into the ZIP without parsing — the file is small (typically <2KB), full parse+serialize is fine.
- `xl/_rels/workbook.xml.rels` mutation. Touched only by RFC-035 (add new sheet); RFC-035 owns that change.

## Acceptance: shipped via 5-commit slice on feat/native-writer (commits 2f3d5a7..<final>) on 2026-04-25

Three of the four RFC-013 primitives ship live, one ships as scaffolding:

- **`file_adds` / `file_deletes`** (commit 1, `2f3d5a7`): Live. `file_adds` is exercised by RFC-020's optional `docProps/core.xml` add path. `file_deletes` is reserved for RFC-035 — provisioned but unused this slice.
- **`ContentTypesGraph`** (commit 2, `dcdf51e`): Live. Parse + idempotent `add_override` / `ensure_default` + source-order-preserving `serialize`. Wired into `do_save`'s Phase-2.5c via `apply_op` in commit 4.
- **Phase-2.5c content-types aggregation** (commit 4, `3785e4a`): Live. Cross-sheet collection in `sheet_order`, single parse + serialize of `[Content_Types].xml`. No volume producer this slice; first callers will be RFC-022 (Hyperlinks), RFC-023 (Comments), RFC-024 (Tables).
- **`AncillaryPartRegistry`** (commit 3, `2c2ca74`): Scaffolding. Lazy `populate_for_sheet` reads `_rels/sheetN.xml.rels` and classifies entries by `wolfxl_rels::rt::*` URI. No live caller this slice — first caller is RFC-022.

**Test coverage**: 4 inline Rust tests (compile-checked under cdylib link constraint), 6 pytest integration tests in `tests/test_patcher_infra.py` exercising `_test_inject_file_add`, `_test_queue_content_type_op`, `_test_populate_ancillary` hooks. The byte-identical no-op-save guard catches any future regression that would silently rewrite source ZIP entries when no caller has asked for it.
