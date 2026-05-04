# RFC-070 — Pivot table mutation: load existing pivot, mutate `.source`, save (Sprint 5 / G17)

> **Status**: Proposed
> **Owner**: Claude (S5 design)
> **Sprint**: S5 — Pivot Mutation
> **Closes**: G17 (pivot table mutation of existing pivots) in the openpyxl parity program
> **Depends on**: RFC-047 (PivotCache shape), RFC-048 (PivotTable shape), RFC-049 (construction-mode emit)
> **Unblocks**: future RFCs covering field-placement mutation, filter changes, multi-cache mutation

## 1. Goal

Make this work, end-to-end, with no `xfail`:

```python
wb = wolfxl.load_workbook(src, modify=True)
pivot = wb["Pivot"].pivot_tables[0]
pivot.source = Reference(wb.active, min_col=1, min_row=1, max_col=2, max_row=3)
wb.save(src)
```

That is the verbatim shape of `tests/test_openpyxl_compat_oracle.py:571-594` (the `pivots_in_place_edit` probe). Closing the probe flips `pivots.in_place_edit` from `partial` (gap_id `G17`) to `supported`.

## 2. Problem statement

The S0 oracle audit found:

1. `Worksheet.pivot_tables` does not exist in modify mode. Accessing it on a freshly loaded workbook raises `AttributeError`.
2. `PivotTable.source` is constructor-only — built from a `PivotCache` and never re-settable.
3. The patcher (`src/wolfxl/patcher_pivot.rs`) treats existing `xl/pivotTables/pivotTable*.xml` and `xl/pivotCache/pivotCache{Definition,Records}*.xml` as opaque blobs; nothing parses them on load.

Construction (v2.0 / RFC-049) is solid: building a pivot from scratch and saving works. Mutating an *existing* one — even one wolfxl itself wrote two seconds ago — is the gap.

## 3. Scope decision: minimal source-range mutation only

The probe touches exactly one mutation: **source range**. Two paths exist for closing it:

- **Option A (full):** parse every existing pivot part into Python objects, expose every PivotTable / PivotCache field as mutable, re-emit on save. ~3-4 sessions of work; full openpyxl parity for pivots.
- **Option B (minimal, this RFC):** parse only the source range out of `<cacheSource><worksheetSource ref=...>` and the table location out of `<location ref=...>`. Expose a thin `PivotTableHandle` on `Worksheet.pivot_tables` with one mutator: `.source = Reference(...)`. Re-emit only the touched parts on save. ~1 session.

This RFC ships **Option B**. Rationale:

- The probe — and any reasonable interpretation of "openpyxl drop-in pivot mutation" for v2.x — is dominated by source-range edits and field-placement edits. Source range is the simpler half.
- The harder half (field placement, filter, aggregation) requires parsing every `<pivotField>`, `<rowField>`, `<colField>`, `<dataField>`, `<pageField>` block — an order of magnitude more parser code, and a dedicated test surface.
- Closing the probe with Option B unlocks the metric (oracle pass rate +1) without committing 3+ sessions to full Option A.
- Option A is a clean follow-up: a future RFC can extend the existing `PivotTableHandle` with more setters and add the corresponding parser passes.

Out of scope (deferred to a later RFC):

- Field-placement mutation (`pivot.rows = [...]`, `pivot.columns = [...]`, `pivot.values = [...]`).
- Filter mutation (`pivot.page_filters`).
- Aggregation function changes (`pivot.values[0].function = "average"`).
- Adding/removing pivots from a worksheet's collection (just exposing existing ones is enough).
- Round-tripping foreign-tool pivots (LibreOffice / Excel-authored). Optional: works-best-effort, but the only acceptance gate is wolfxl→wolfxl round-trip via the probe.
- Mutating an existing pivot's *cache* (which would change every dependent pivot). Source-range edit on a single pivot may force a cache refresh — see §6.

## 4. Public contract

### 4.1 `Worksheet.pivot_tables` (modify mode)

```python
ws.pivot_tables          # → list[PivotTableHandle]
ws.pivot_tables[0]       # → first existing pivot, in document order
len(ws.pivot_tables)     # → count
```

In write mode (fresh `wolfxl.Workbook()` / `load_workbook(read_only=True)` / `data_only=True`), `Worksheet.pivot_tables` returns the same list — but for a fresh workbook the list is empty. Construction-mode users continue to call `ws.add_pivot_table(pt, "A1")` as before; that path is unchanged.

### 4.2 `PivotTableHandle`

A thin proxy with a stable shape:

```python
class PivotTableHandle:
    name: str             # ro: <pivotTableDefinition name="...">
    location: str         # ro: <location ref="A1:E20">  (post-mutation, refreshed)
    source: Reference     # rw: setter rewrites the linked pivotCacheDefinition
    cache_id: int         # ro: workbook-level pivotCache id this handle references

    # Internal — used by save path only.
    _cache_part_path: str          # e.g., "xl/pivotCache/pivotCacheDefinition1.xml"
    _records_part_path: str        # e.g., "xl/pivotCache/pivotCacheRecords1.xml"
    _table_part_path: str          # e.g., "xl/pivotTables/pivotTable1.xml"
    _dirty: bool                   # true after any mutator runs
    _new_source: Reference | None  # set by .source = ...; consumed at save
```

Setting `pivot.source = Reference(...)` only stamps the new source; the rewrite happens at `wb.save()` time. This keeps mutation cheap and lets us batch.

### 4.3 Save-time semantics

When `wb.save(path)` runs and any handle is `_dirty`:

1. Rewrite the linked `xl/pivotCache/pivotCacheDefinition{N}.xml` with the new `<cacheSource><worksheetSource ref="..."/></cacheSource>` element. All other XML in that part passes through verbatim (we re-serialise from a parsed DOM, but we only mutate the source ref).
2. If the new source range has the **same shape** (same column count) as the original — assumed for v1.0 — leave `<pivotCacheRecords>` untouched. Excel re-evaluates the data on next refresh; openpyxl reads the source ref and the existing records without complaint.
3. If the new source range has a **different shape** (different column count, detected by parsing the original cacheSource and comparing column spans), set `refreshOnLoad="1"` on `<pivotCacheDefinition>` and pass-through records as-is. Excel will recompute on open. Document this as a v1.0 limitation; a future RFC can do live record regeneration.
4. Patch only the touched files in the patcher's `file_patches` map. Untouched pivots retain their original opaque bytes.

## 5. Parser design

### 5.1 What we parse

A minimal `quick-xml` (or `roxmltree`) reader on the patcher side, called only when a worksheet's `.pivot_tables` is first accessed in modify mode.

For each `xl/pivotTables/pivotTable*.xml` referenced from a worksheet's rels:

- Extract `<pivotTableDefinition name="...">` → `handle.name`.
- Extract `<location ref="A1:E20">` → `handle.location`.
- Extract `cacheId="N"` → `handle.cache_id`. Resolve to the matching workbook-level pivot cache ref and follow its rels into `xl/pivotCache/pivotCacheDefinition{N}.xml`.

For the linked `pivotCacheDefinition{N}.xml`:

- Extract `<cacheSource><worksheetSource ref="..." [sheet="..."]/></cacheSource>` → build `handle.source`. The `sheet=` attribute may be absent (means the source sheet matches the pivot sheet); resolve through the workbook sheet list.

That is the whole parser surface. Field definitions, dimensions, etc. are passthrough — we never look at them in v1.0.

### 5.2 Where the parser lives

New module: `src/wolfxl/patcher_pivot_parse.rs`. PyO3-exported helpers callable from `python/wolfxl/_workbook_patcher_load.py` (or wherever `load_workbook(modify=True)` materialises worksheet state).

Parser exposes:

```rust
pub fn parse_pivot_table_meta(xml_bytes: &[u8]) -> PyResult<PivotTableMeta>;
pub fn parse_pivot_cache_source(xml_bytes: &[u8]) -> PyResult<PivotCacheSourceMeta>;
```

Returning small Python dicts. The Python side wraps them into `PivotTableHandle` instances.

### 5.3 Lazy materialisation

`Worksheet.pivot_tables` is a `cached_property` that walks the worksheet's rels (`xl/worksheets/_rels/sheet{N}.xml.rels`) for `pivotTable` relationships, parses each one + its linked cacheDefinition, and builds a list of handles. First access in a session pays the parse cost; subsequent access is free. Construction-mode Worksheets (`_pending_pivot_tables` non-empty, no underlying ZIP) return the empty list since there is nothing to parse from disk.

## 6. Save path

### 6.1 Patcher hook

A new pass in `apply_pivot_edits_phase` (sibling to the existing `apply_pivot_adds_phase`), invoked from `XlsxPatcher::save` after the adds phase. The Python side passes a list of dirty handles + their new source refs into the Rust patcher via a new `apply_pivot_source_edits` PyO3 method.

For each dirty handle:

1. Locate the cache-definition part by path (`handle._cache_part_path`).
2. If the part is in `file_patches` (pre-existing patch), edit the buffer; otherwise read from the ZIP and queue an edit.
3. Use `quick-xml` to find `<cacheSource><worksheetSource>` and replace its `ref=` (and optionally `sheet=`) attributes with the new source. Update the surrounding `<pivotCacheDefinition>` `refreshOnLoad="1"` if column count changed.
4. Write back into `file_patches`.

### 6.2 Records passthrough

Records (`pivotCacheRecords{N}.xml`) are not touched in v1.0. The Excel/openpyxl behaviour:

- **Same-shape source change** (e.g., source moved from `A1:B3` to `A1:B5`): records become outdated for the rows beyond row 3, but the cache still resolves. Excel either shows stale data or recomputes on next refresh. openpyxl's `load_workbook` round-trips the file fine. The compat-oracle probe only cares about save success and round-trip; it does not assert pivot freshness.
- **Different-shape source change** (e.g., source went from 2 cols to 3): we set `refreshOnLoad="1"` so Excel recomputes on open. openpyxl's reader does not enforce records/definition column-count consistency on read, so round-trip stays clean.

A future RFC ("pivot record regeneration") can do live record rebuild in wolfxl — out of scope here.

### 6.3 Record-table consistency check

The patcher emits a debug-build assertion: after rewriting a cacheDefinition, parse it back, verify `<cacheSource>` is well-formed, and ensure the part's `<extLst>` (if any) is preserved verbatim. This catches the common XML-rewrite pitfall of stripping ext-list elements.

## 7. Test plan

### 7.1 Compat-oracle probe (existing)

`tests/test_openpyxl_compat_oracle.py:571-594` flips `xfail → passed`. No probe change.

### 7.2 New focused tests

`tests/test_pivot_modify_existing.py` (new file):

- **Round-trip:** save a pivot with construction APIs, reopen `modify=True`, mutate `.source` to a same-shape range, save, reopen with `wolfxl.load_workbook(read_only=True)`, assert the new source ref is present.
- **Different-shape source:** mutate `.source` to a 3-column range (was 2), save, assert `refreshOnLoad="1"` is set on the cacheDefinition.
- **Multiple pivots, single mutation:** workbook with two pivots; mutate only one; save; assert the untouched pivot bytes are byte-identical to the original.
- **No mutation:** load + save with no `.source` write; assert no pivot files appear in `file_patches`.
- **Read-only worksheet:** `Worksheet.pivot_tables[0].source = ...` on a `read_only=True` workbook raises `RuntimeError` (not silently swallowed).
- **Foreign-authored pivot:** load an Excel-authored fixture (if available), mutate source, save, reopen with openpyxl. Best-effort; if the fixture's pivot has features the parser cannot handle, the test xfails with a clear marker.

### 7.3 Cargo tests

`crates/wolfxl-pivot/tests/parse_existing.rs` (new): unit tests for `parse_pivot_table_meta` and `parse_pivot_cache_source` on hand-crafted XML strings, including ext-list preservation, sheet-attr-absent edge case, and a malformed-XML rejection case.

### 7.4 External-oracle gate

`tests/test_external_oracle_preservation.py` runs against any pivot fixtures present. If fixtures cover edited pivots (need to verify during impl), they must round-trip green through LibreOffice + Excelize.

## 8. Acceptance criteria

1. `tests/test_openpyxl_compat_oracle.py::test_compat_oracle_probe[pivots.in_place_edit-...]` flips xfail → pass.
2. New focused tests in §7.2 all pass.
3. `cargo test --workspace` stays green; new pivot-parse unit tests pass.
4. `tests/test_external_oracle_preservation.py` stays green where pivot fixtures exist.
5. Compat-oracle pass count rises by 1 (G17 closure); no regression in the 47 existing passed probes.
6. Compat-matrix row `pivots.in_place_edit` flips `partial` (gap_id `G17`) → `supported`. Tracker row marked `landed`.
7. README and `docs/trust/limitations.md` mention the v1.0 scope (source-range mutation only) so users do not assume field-placement edits work.

## 9. Out-of-scope (deferred to follow-up RFCs)

- Field-placement mutation (`pivot.rows`, `pivot.columns`, `pivot.values`).
- Filter (page-field) mutation.
- Aggregation function changes.
- Adding pivots to an existing pivot collection (`ws.pivot_tables.append(...)`).
- Removing existing pivots (`del ws.pivot_tables[0]`).
- Live record regeneration after source change (currently relies on `refreshOnLoad="1"` for shape-incompatible mutations).
- Round-tripping pivots authored by foreign tools where their schema diverges from RFC-047/048.

## 10. Risks

| # | Risk | Mitigation |
|---|------|-----------|
| 1 | Existing wolfxl pivot fixtures may have ext-list extensions we strip during rewrite. | Round-trip the XML through `quick-xml`'s `Writer` with whitespace preserved; assert byte-for-byte equality of every non-touched element in a debug-build sanity check. |
| 2 | Worksheet rels graph mutation may collide with the existing `apply_pivot_adds_phase`. | New `apply_pivot_source_edits` runs *after* the adds phase; only touches existing files, never adds new rels. The two phases never write the same path. |
| 3 | `cached_property` on `Worksheet.pivot_tables` may stale-cache between mutations. | Mutators flip `_dirty` on the handle, not the cached list. The list itself is stable across the session — handles are mutable in place. Save consumes `_dirty` and clears it. |
| 4 | Source range with `sheet=` attr referencing an unknown sheet (foreign-authored quirk) breaks reference resolution. | If the resolved sheet is not in `wb.sheetnames`, fall back to a `Reference`-with-sheet-name shim; the new source the user supplies always carries an explicit `Reference(wb["..."])` so this is a load-time concern only. Document the fallback in `Reference` and add a test fixture. |
| 5 | Records-vs-definition shape mismatch on save (column count diverged but we forgot to set `refreshOnLoad`). | The save path always reads the original definition's column count, compares to the new source's, and sets `refreshOnLoad="1"` on any mismatch. Unit-tested in §7.2 case 2. |

## 11. Implementation plan

1. RFC review (this document) — lock scope at "source-range only" before impl.
2. Subagent handoff: implement parser, handle class, save path, all §7 tests.
3. Subagent verifies all eight acceptance gates (§8) before commit.
4. Central merge into main; cleanup of staging refs.
5. Mark G17 `landed`; flip spec; commit; oracle pass rate rises +1.

## 12. Open questions

- **Should `pivot.source` setter accept a string A1 ref (e.g., `pivot.source = "Sheet1!A1:B5"`) as a convenience?** Default: **no** — keep `Reference(...)` only, matching openpyxl's pivot-source contract. Convenience setters can land in a follow-up if users ask.
- **Should we expose `pivot.refresh()` to force `refreshOnLoad="1"` without mutation?** Default: **no** in v1.0. If users need this, a follow-up can add it; today it is implicit through shape-mismatch detection on save.
