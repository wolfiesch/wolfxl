# Parity Closure — Empty `limitations.md`

Date: 2026-05-04
Status: Active. Sequel to `Plans/openpyxl-parity-program.md` (S0-S7 landed; Phases 1-4 and Phase 10 have landed; Phase 9 is implemented locally on `parity-closure-phase9-cf-ratchet`). This plan commits to closing every remaining openpyxl-parity gap so that `docs/trust/limitations.md` becomes empty (deleted) and the `_compat_spec.py` totals reach `supported=100%, partial=0, not_yet=0` for everything openpyxl itself supports.

## Context

The user has demanded zero openpyxl-parity gaps. The current state, per the 2026-05-04 review:

- Phase 1 recategorized wolfxl-extras (`.xls` writes, `.xlsb` writes, `.ods` read+write, `.xls` style reads, and VBA authoring) as `out_of_scope`, because openpyxl exposes none of those surfaces.
- Phase 2 closed `defined_names.edge_cases`; Phase 3 closed `print_settings.depth`; Phase 4 closed `array_formulas.spill`. Phase 10 added compatibility-oracle probes and exposed one intentional `coordinate_to_tuple` strictness delta. Phase 9 is now adding CF secondary probes plus a broader dynamic-surface ratchet.
- `docs/migration/_compat_spec.py`: 66 supported / 2 partial / 1 not_yet / 6 out_of_scope across 75 entries.
- The dynamic-surface ratchet (`tests/test_openpyxl_dynamic_surface.py`) now compares 22 representative objects: Workbook, Worksheet, Cell, Font, Fill, PatternFill, GradientFill, Border, Alignment, Protection, NamedStyle, DefinedName, ExternalLink, BarChart, Image, Comment, Table, DataValidation, ConditionalFormatting, ConditionalFormattingList, AutoFilter, and PageSetup. Pivot and slicer internals remain intentionally behavioral-oracle-first because their openpyxl construction surface is mostly internal state rather than stable end-user API.
- The openpyxl source-distribution test corpus has never been vendored, even though the parity program plan called for it in S1+.

The remaining parity work is now intentionally staged v2 depth: pivot field/filter/aggregation mutation and external-link authoring. Dynamic-array spill metadata, calc-chain edge cases, and coordinate utility strictness have parity coverage; standalone table-driven slicers were reclassified as a possible wolfxl-extra because openpyxl 3.1.5 has no public standalone slicer authoring API. Follow-up acceptance targets live in `Plans/followups/remaining-v2-depth.md`.

## Definition of "no gaps with openpyxl"

Three measurements, each must be green:

1. **Compat oracle**: `tests/test_openpyxl_compat_oracle.py` reports 0 xfails / 0 fails on every probe. Spec entries: 0 `partial`, 0 `not_yet` for any row where openpyxl exposes the surface. `out_of_scope` allowed only for things openpyxl itself doesn't do.
2. **Dynamic-surface ratchet**: `tests/test_openpyxl_dynamic_surface.py` covers Workbook, Worksheet, Cell, Font, Fill, Border, Alignment, Protection, NamedStyle, GradientFill, DefinedName, ExternalLink, PivotTable, PivotCache, Slicer, BarChart (representative chart), Image, Comment, Table, DataValidation, ConditionalFormatting, AutoFilter, PageSetup. ~23 classes. Zero `missing_callables`, `missing_values`, or `type_mismatches`.
3. **Vendored openpyxl test corpus**: `scripts/fetch_openpyxl_corpus.py` (currently absent) fetches openpyxl's source-distribution `tests/` directory under a `sys.modules` shim that swaps `openpyxl` → `wolfxl`. The shimmed run reaches ≥99% pass on the corpus; remaining failures are documented as openpyxl-side bugs or test-infra issues.

When all three are green, `docs/trust/limitations.md` is deleted (or reduced to a single sentence: *"None known. See [compatibility-matrix.md](../migration/compatibility-matrix.md) for the machine-checked support spec."*). README's drop-in claim becomes honest.

## Scope

**In:** every openpyxl-parity gap, the measurement infrastructure, and the cleanup of the limitations doc.

**Out:** `.xls` / `.xlsb` / `.ods` writes, VBA authoring, `.xls` style reads. These are wolfxl-extras (openpyxl doesn't support them either). They get recategorized in `_compat_spec.py` from `not_yet` to `out_of_scope` with a note: *"openpyxl does not support this format; wolfxl tracks it as a separate roadmap item, not an openpyxl-parity gap."* They do not appear in `limitations.md`.

## Phase plan

Each phase is a separate landing on its own branch. Phases are largely independent except the explicit dependencies called out. Per-phase verification gate at end of plan.

### Phase 1 — Recategorize wolfxl-extras (1 session, mechanical) [LANDED 2026-05-04, commit 4825baf]

**Goal:** stop conflating "openpyxl can't do this" with "wolfxl is missing a feature".

**Files:**
- `docs/migration/_compat_spec.py`: change status of `read.xls` (currently partial), `legacy_formats.xlsb_write`, `legacy_formats.xls_write`, `legacy_formats.ods_read_write`, `vba.author` from `partial`/`not_yet` to `out_of_scope` with a `notes` field stating: *"openpyxl does not support this surface. Tracked as a wolfxl-extra roadmap item, not an openpyxl-parity gap."*
- `Plans/openpyxl-parity-program.md`: update G25-G28 status table rows to "out_of_scope (wolfxl-extra)" + drop them from the "≥95% pass-rate" denominator calculation.
- Run `scripts/render_compat_matrix.py` to regenerate `docs/migration/compatibility-matrix.md`.
- `docs/trust/limitations.md`: remove the four wolfxl-extra rows (.ods, xlsb/xls writes, .xls style read, VBA authoring). Keep only the rows that map to genuine openpyxl-parity gaps.

**Verification:** matrix totals shift to "supported / partial / out_of_scope" with no `not_yet`. Compat-oracle passes unchanged.

### Phase 2 — G22 Defined-name edge cases (1-2 sessions) [LANDED 2026-05-04, branch parity-closure-phase2]

**Outcome:**
- All 13 ECMA-376 §18.2.5 attributes (`hidden`, `comment`, plus 11 G22 additions: `customMenu`, `description`, `help`, `statusBar`, `shortcutKey`, `function`, `functionGroupId`, `vbProcedure`, `xlm`, `publishToServer`, `workbookParameter`) round-trip across all three save paths (write mode, modify mode, read).
- Pre-existing reader bug fixed: `comment` and `hidden` previously parsed as 0 attrs, now parse all 13.
- Compat oracle: `defined_names.edge_cases` xfail → pass (53 pass / 1 xfail, was 51 / 2). Cross-tool sub-probe verifies openpyxl reads what wolfxl writes for all 13 attrs.
- Spec totals: 65 supported / 3 partial / 1 not_yet / 6 out_of_scope (was 64 / 4 / 1 / 6).
- Phase 2 suite delta (excluding 13 pre-existing baseline image/chart failures): +1 pass, -1 xfail, otherwise identical.

**Codex review findings (post-implementation):**
- 🚫 MUST-FIX: `hidden=False` in modify mode could not clear an existing `hidden="1"` because the Python patcher payload omitted the key on False, so the Rust upsert saw `None` (preserve source) instead of `Some(false)` (remove). **Fixed**: `python/wolfxl/_workbook_patcher_flush.py:_defined_name_payload` now sends `hidden` as an explicit `bool` always. Regression test `test_g22_modify_mode_clears_hidden_when_set_false` lives in `tests/test_defined_names_modify.py`.
- ⚠️ SHOULD-FIX (deferred): `shortcutKey` length and `functionGroupId` 0–11 range validation missing. ECMA-376 imposes these but Excel itself accepts deviations; defer to a future hardening pass rather than block Phase 2.
- ⚠️ SHOULD-FIX (deferred): Modify-mode upsert preserves source attr order rather than re-canonicalizing. Write/modify byte-identical output is not a Phase 2 contract — round-trip preservation is, and that holds.
- ✅ Confirmed correct: write/modify emit-order parity for newly-emitted elements, reader capture of all 13 attrs, xlsb path defaulting via `..Default::default()`, compat-oracle cross-tool sub-probe coverage.


**Goal:** flip `defined_names.edge_cases` to `supported`. Match openpyxl's full `DefinedName` attribute surface so round-trip preserves every standard ECMA-376 attr.

**Scope refinement after exploration (2026-05-04):** original plan listed 6 attrs; the full openpyxl surface has 11 missing attrs. We're closing all 11 to honor the "no gaps" goal. Additionally the reader has a pre-existing bug — it currently parses only `name`, `localSheetId`, and `refersTo` (the inner formula). The two attrs the Python dataclass already declares (`comment`, `hidden`) are write-only because the reader never extracts them. That bug is in scope here because the existing probe xfails on it.

**Attrs to add (11):**

| openpyxl name | Python field | XML attr | Type |
|---|---|---|---|
| `customMenu` | `custom_menu` | `customMenu` | str |
| `function` | `function` | `function` | bool |
| `functionGroupId` | `function_group_id` | `functionGroupId` | int |
| `shortcutKey` | `shortcut_key` | `shortcutKey` | str (1 char) |
| `statusBar` | `status_bar` | `statusBar` | str |
| `description` | `description` | `description` | str |
| `help` | `help` | `help` | str |
| `vbProcedure` | `vb_procedure` | `vbProcedure` | bool |
| `xlm` | `xlm` | `xlm` | bool |
| `publishToServer` | `publish_to_server` | `publishToServer` | bool |
| `workbookParameter` | `workbook_parameter` | `workbookParameter` | bool |

**Pre-existing reader bug to fix:** `crates/wolfxl-reader/src/lib.rs:1597-1631` parses 3 attrs (`name`, `localSheetId`, inner-text `refersTo`). Extend to also parse `comment`, `hidden`, and the 11 new attrs. Without this, `comment` and `hidden` round-trip silently fails on save (write-only).

**Surgery points:**

1. **`python/wolfxl/workbook/defined_name.py:27-99`** — extend `DefinedName` dataclass:
   - Add 11 new fields to `__slots__` and `__init__` with defaults `None` / `False`.
   - Extend `attr_text` property and dict-serialization helpers (used by `to_dict`/`__repr__`).
   - openpyxl camelCase aliases via `__getattr__` (matches existing `localSheetId` → `local_sheet_id` pattern).
   - **~110 LOC**

2. **`crates/wolfxl-writer/src/model/defined_name.rs:6-37`** — extend `DefinedNameSpec`:
   - Add 11 fields (`Option<String>` / `Option<bool>` / `Option<u32>`).
   - Update `From<&PyDict>` impl to extract them.
   - **~60 LOC**

3. **`crates/wolfxl-writer/src/emit/workbook_xml.rs:48-72`** — emit attrs on `<definedName>`:
   - Add 11 conditional `write_attribute` calls (skip when `None`/default per ECMA-376).
   - bool attrs: emit `="1"` only when `true` (XML default omission).
   - **~50 LOC**

4. **`crates/wolfxl-reader/src/lib.rs:1597-1631`** — fix the parse bug + add new attrs:
   - Extract `comment`, `hidden` (existing dataclass fields, currently write-only).
   - Extract the 11 new attrs.
   - Pass through to the `DefinedName` Python ctor.
   - **~70 LOC**

5. **`src/wolfxl/defined_names.rs:1-723`** — extend `DefinedNameMut` for modify-mode parity:
   - Add the 11 fields + `comment`/`hidden` setters (already present for the latter).
   - The unknown-attr-preservation logic at line 47-50 currently saves us in modify mode; once attrs are explicit, drop unknown-attr passthrough for these 11 (still pass through truly-unknown attrs).
   - **~80 LOC**

6. **`tests/test_openpyxl_compat_oracle.py:1257-1280`** — extend probe:
   - Construct a `DefinedName` exercising all 13 attrs (11 new + `comment` + `hidden`).
   - Round-trip via `wolfxl.save` → `wolfxl.load_workbook` and assert each attr survives.
   - Add a second sub-probe that round-trips through openpyxl as the reader (the cross-tool oracle): write with wolfxl, read with openpyxl, assert openpyxl sees the same attrs.
   - Remove `xfail` if currently marked.
   - **~50 LOC**

7. **`docs/migration/_compat_spec.py`** — flip `defined_names.edge_cases` from `partial` to `supported`; clear `gap_id`.

**Revised estimate:** ~420 LOC + 1-2 sessions (1 for impl, 0.5 for cross-oracle probe + maturin develop loop).

**Verification:**
- `cargo test --workspace` clean.
- `uv run --no-sync maturin develop --release` rebuilds.
- `uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -k defined_names_edge_cases -v` — passes (no xfail).
- `uv run --no-sync pytest tests/test_openpyxl_dynamic_surface.py -v` — does not regress (DefinedName not in current ratchet scope, but checking for crash regressions on the workbook-level diff).
- `python scripts/render_compat_matrix.py` — totals: supported +1, partial -1.

### Phase 3 — G24 Print-settings depth (2 sessions) [LANDED 2026-05-04, branch parity-closure-phase2]

**Outcome:**
- Full `PageSetup` depth now round-trips across write mode, modify mode, and native read, including the G24 attrs `paperHeight`, `paperWidth`, `pageOrder`, and `copies`.
- `<printOptions>` is no longer parse-only. All 5 attrs (`horizontalCentered`, `verticalCentered`, `headings`, `gridLines`, `gridLinesSet`) emit in write mode, patch in modify mode, read back through wolfxl, and are accepted by openpyxl.
- Print titles were confirmed as workbook-scoped `_xlnm.Print_Titles` defined names rather than sheet XML. The save path now emits them through the existing defined-name machinery for both write mode and modify mode.
- Compat oracle: `print_settings.depth` xfail → pass. The expanded probe covers 18 PageSetup attrs, 5 PrintOptions attrs, 6 PageMargins attrs, plus `print_area`, `print_title_rows`, and `print_title_cols`, with an openpyxl reader cross-check.
- Spec totals: 66 supported / 2 partial / 1 not_yet / 6 out_of_scope (was 65 / 3 / 1 / 6).

**Goal:** flip `print_settings.depth` to `supported`. Cover the full openpyxl `PageSetup` + `PrintOptions` + `PageMargins` round-trip surface.

**Scope refinement after exploration (2026-05-04):** original plan said 17 attrs missing. Reality is more nuanced:
- **PageSetup**: 4 attrs missing (`paperHeight`, `paperWidth`, `pageOrder`, `copies`). The other "missing" attrs the original list called out (`gridLines`, `gridLinesSet`, `headings`, `horizontalCentered`, `verticalCentered`, `cellComments`, `errors`, `horizontalDpi`, `verticalDpi`, `useFirstPageNumber`, `draft`) are either parsed-but-not-emitted (PrintOptions group) or already present in PageSetup.
- **PrintOptions**: parse-only. Reader extracts 5 attrs (`horizontalCentered`, `verticalCentered`, `headings`, `gridLines`, `gridLinesSet`); writer has no `<printOptions>` emit path. Round-trip silently drops them on save. **This is the load-bearing gap.**
- **PageMargins**: complete (left/right/top/bottom/header/footer all read+written).

**Surgery points:**

1. **`python/wolfxl/worksheet/page_setup.py`** — extend `PageSetup`:
   - Add 4 fields: `paper_height: str | None`, `paper_width: str | None`, `page_order: str | None` (default `"downThenOver"`), `copies: int | None`.
   - openpyxl camelCase aliases.
   - **~40 LOC**

2. **`crates/wolfxl-writer/src/parse/sheet_setup.rs:337-405`** — extend PageSetupSpec:
   - Add 4 fields to the spec struct.
   - Extend `From<&PyDict>` extraction.
   - Extend `<pageSetup>` emit: 4 conditional `write_attribute` calls.
   - **~50 LOC**

3. **`crates/wolfxl-writer/src/parse/sheet_setup.rs`** — add new `PrintOptionsSpec` with emit path:
   - 5 fields (`horizontal_centered`, `vertical_centered`, `headings`, `grid_lines`, `grid_lines_set`), all `Option<bool>`.
   - `From<&PyDict>` extraction.
   - New emit method: write `<printOptions>` element with conditional attrs (skip element entirely if all `None`/`false`).
   - Wire emit call into the sheet write sequence (between `<pageMargins>` and `<pageSetup>` per ECMA-376 element order — verify exact position against openpyxl-produced fixtures).
   - **~120 LOC**

4. **`crates/wolfxl-reader/src/native_reader_page_setup.rs:127-178`** — extend `page_setup_to_py`:
   - Parse 4 new PageSetup attrs.
   - PrintOptions parse path already exists at lines 168-178 — confirm it's reached and the dict is attached to the Worksheet (currently it may be parse-only on the Rust side without a Python-side surface — worth verifying during impl).
   - **~30 LOC**

5. **`src/wolfxl/patcher_sheet_blocks.rs:637-701`** — extend `apply_sheet_setup_phase`:
   - Route the 4 new PageSetup attrs.
   - Route the new PrintOptionsSpec dict (currently no `<printOptions>` patcher hook exists; add one parallel to the `<pageMargins>` hook).
   - **~70 LOC**

6. **`tests/test_openpyxl_compat_oracle.py:1287-1345`** — rewrite `print_settings_depth` probe:
   - Replace the existing 12-attr probe with a 30-attr probe covering: 18 PageSetup attrs (14 existing + 4 new), 5 PrintOptions attrs, 6 PageMargins attrs, plus the non-attr surface (`print_title_rows`, `print_title_cols`, `print_area`).
   - Cross-tool sub-probe: write with wolfxl → read with openpyxl → assert all 30 attrs survive.
   - **~80 LOC**

7. **`docs/migration/_compat_spec.py`** — flip `print_settings.depth` from `partial` to `supported`; clear `gap_id`.

**Revised estimate:** ~390 LOC + 2 sessions. The PrintOptions writer infrastructure is a non-trivial new emit path (not just attr extension); it dominates Phase 3 LOC.

**Verification:**
- `cargo test --workspace` clean.
- `uv run --no-sync maturin develop --release` rebuilds.
- `uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -k 'print_settings' -v` — `print_settings_basic` + `print_settings_depth` both pass.
- `uv run --no-sync pytest tests/test_external_oracle_preservation.py -v` — 7/7 fixtures still pass (these include real workbooks with print settings; ratchet against silent regressions).
- Round-trip a real openpyxl-produced file with `<printOptions>` set → wolfxl save → diff XML byte-by-byte: `<printOptions>` element survives.
- `python scripts/render_compat_matrix.py` — totals: supported +1, partial -1.

### Phase 4 — G07 Dynamic-array spill metadata (2 sessions)

**Goal:** flip `array_formulas.spill` to `supported`.

**Surface:** openpyxl exposes `Worksheet.formula_attributes` mapping cell refs to spill metadata. wolfxl currently parses dynamic-array formulas (`_xlfn.SUM`, `_xlfn.LAMBDA` etc.) but does not preserve `<extLst><ext><x14:formulaAttribute>` blocks or `cm=` / spill flags.

**Files:**
- `python/wolfxl/_worksheet_records.py`: add `formula_attributes` dict and `SpilledFormula` dataclass.
- `python/wolfxl/_cell.py`: route spill metadata through cell read/write.
- `crates/wolfxl-reader/src/worksheet/formulas.rs`: parse `<extLst>` formula attributes.
- `crates/wolfxl-writer/src/emit/sheet/cells.rs`: emit `cm=` and `<extLst>` round-trip.
- `src/wolfxl/patcher_sheet_blocks.rs`: preserve through modify mode.
- `tests/test_openpyxl_compat_oracle.py`: define probe; round-trip via openpyxl reader.

**Estimate:** 320 LOC + 2 sessions.

### Phase 5 — G21 Standalone (table-driven) slicers (2 sessions)

**Goal:** flip `slicers.standalone` to `supported`.

**Surface:** openpyxl supports slicers tied to a Table (not only PivotCache). wolfxl's `SlicerCache` currently requires a `PivotCache`.

**Files:**
- `python/wolfxl/pivot/_slicer.py`: add `TableSlicerCache` with `cache.source = TableReference(...)`.
- `python/wolfxl/_table/_table.py`: thread slicer registration.
- `crates/wolfxl-writer/src/emit/slicer.rs`: emit table-source slicer cache (tabular `<tableSource>` block instead of `<pivotSource>`).
- `crates/wolfxl-reader/src/slicer.rs`: parse the table-driven variant.
- `src/wolfxl/patcher_workbook.rs`: round-trip new slicer caches in modify mode.
- `tests/test_openpyxl_compat_oracle.py`: define `slicers_standalone` probe.

**Estimate:** 280 LOC + 2 sessions.

### Phase 6 — G23 Calc-chain edge cases (3 sessions)

**Goal:** flip `calc_chain.edge_cases` to `supported`.

**Surface:** cross-sheet ordering, deleted-cell pruning, `calcChainExtLst` ext block.

**Files:**
- `crates/wolfxl-writer/src/emit/calc_chain.rs`: deleted-cell pruning + topological sort across sheets.
- `crates/wolfxl-reader/src/calc_chain.rs`: parse `extLst` ext block.
- `src/wolfxl/patcher_workbook.rs::apply_calc_chain_phase`: extend rebuild logic to handle cross-sheet ordering and prune entries for deleted cells.
- `python/wolfxl/_workbook_calc.py` (47 LOC, currently internal-only): grow to handle pruning + cross-sheet ordering bookkeeping.
- `tests/test_openpyxl_compat_oracle.py`: extend `calc_chain_edge_cases` probe (cross-sheet + deleted cell + ext block).

**Estimate:** 420 LOC + 3 sessions.

### Phase 7 — G17 follow-up: pivot field/filter/aggregation mutation (4 sessions)

**Goal:** flip `pivots.in_place_edit` from "supported (source-range only)" to "supported (full)". This is the largest single phase.

**Surface (from RFC-070 deferred Option-A list):**
- Field placement: mutate `<rowFields>`, `<colFields>`, `<dataFields>`, `<pageFields>` blocks on existing pivot tables.
- Filter mutation: page-field filters and field-level filter overrides.
- Aggregation function changes: `<dataField subtotal="sum">` → `"average"` etc.
- Live record regeneration: when source data changes, recompute aggregates without requiring `refreshOnLoad="1"`.

**Files:**
- `crates/wolfxl-pivot/src/mutate.rs`: add field-block parser+rewriter (currently only `worksheetSource`); add data-field subtotal mutator; add page-field filter mutator.
- `src/wolfxl/patcher_pivot_parse.rs`: expand `parse_pivot_table_meta` to surface field arrays.
- `src/wolfxl/patcher_pivot_edit.rs`: add per-attribute drains (Phase 2.5m-edit fanout).
- `python/wolfxl/pivot/_handle.py`: add `PivotTableHandle.row_fields`, `column_fields`, `value_fields`, `page_fields`, `set_aggregation(field, fn)`, `set_filter(field, criteria)` setters.
- `crates/wolfxl-pivot/src/records.rs`: optional pre-aggregation refresh when wolfxl can compute aggregates Rust-side (dodges `refreshOnLoad`).
- `tests/test_pivot_modify_existing.py`: expand from 6 tests to ~25 covering each mutation type.
- `tests/test_openpyxl_compat_oracle.py`: extend `pivots_in_place_edit` probe + add `pivots_field_mutation`, `pivots_filter_mutation`, `pivots_aggregation_mutation` probes.

**RFC:** RFC-070 Option-A appendix; supersede the v1.0 Option-B notes.

**Estimate:** 2500 LOC + 4 sessions.

### Phase 8 — G18 follow-up: external-link authoring (2 sessions)

**Goal:** flip `external_links.workbook_collection` from "read-only inspection + opaque preservation" to full authoring (append/remove/edit).

**Surface (from RFC-071 deferred list):** `_external_links.append(ExternalLink(...))`, `.remove(idx)`, `.update_target(idx, new_path)`, refreshing `<sheetDataSet>` cached values when wolfxl can dereference the target.

**Files:**
- `crates/wolfxl-writer/src/emit/external_links.rs` (new): emit `xl/externalLinks/externalLink{N}.xml` and `xl/externalLinks/_rels/externalLink{N}.xml.rels`.
- `python/wolfxl/_external_links.py`: convert `_external_links` from a tuple to a mutable collection with `append`/`remove`; add `ExternalLink.update_target`.
- `src/wolfxl/external_links.rs`: extend PyO3 helpers with serialize-side functions.
- `src/wolfxl/patcher_workbook.rs`: hook new external-link parts into the rels graph + content-types in modify mode.
- `tests/test_external_links.py`: expand from inspection-only tests to authoring round-trip tests.
- `tests/test_openpyxl_compat_oracle.py`: extend `external_links_collection` probe to cover authoring.

**RFC:** RFC-071 v2.0 appendix.

**Estimate:** 1300 LOC + 2 sessions.

### Phase 9 — CF rare combos audit + dynamic-surface ratchet expansion (2 sessions)

**Goal:** retire the `limitations.md` clause "rare builder combinations may need a manual probe" and expand the dynamic-surface test from 3 classes to ~23. [IMPLEMENTED LOCALLY 2026-05-04, branch `parity-closure-phase9-cf-ratchet`]

**Subtask 9a — CF audit:**
- Inventory openpyxl's CF rule constructors and operator combinations.
- Round-trip every combo through wolfxl + openpyxl reader. Fix any deltas.
- Add probes for any new rule combos discovered.

**Subtask 9b — Dynamic-surface expansion:**
- `tests/test_openpyxl_dynamic_surface.py`: extend to cover Font, Fill, Border, Alignment, Protection, NamedStyle, GradientFill, DefinedName, ExternalLink, PivotTable, PivotCache, Slicer, BarChart, Image, Comment, Table, DataValidation, ConditionalFormatting, AutoFilter, PageSetup.
- Fix any `missing_callables` / `missing_values` / `type_mismatches` surfaced.

**Estimate:** 500 LOC + 2 sessions.

### Phase 10 — Probe coverage debt (1 session)

**Goal:** every spec entry marked `supported` has a passing probe. Currently 13 supported entries are unprobed (hidden coverage debt).

**Files:**
- `tests/test_openpyxl_compat_oracle.py`: add 13 probes (one per unprobed `supported` entry in `_compat_spec.py`).
- Any entry that fails its new probe gets demoted to `partial` and routes back into Phase 2-9 queues.

**Estimate:** 250 LOC + 1 session.

### Phase 11 — Vendor openpyxl source test corpus (3 sessions)

**Goal:** the ultimate "no gaps" gate — run openpyxl's own tests against wolfxl with a sys.modules shim.

**Files:**
- `scripts/fetch_openpyxl_corpus.py` (new): pin openpyxl version, fetch `openpyxl/tests/`, write to `tests/vendored_openpyxl/` (gitignored or vendored under license-compatible terms).
- `tests/conftest.py`: install a sys.modules shim mapping `openpyxl` → `wolfxl` for the vendored slice only.
- `tests/test_vendored_openpyxl.py`: pytest collector that picks up vendored tests and runs them under the shim.
- Triage failing vendored tests:
  - real wolfxl bugs → fix.
  - openpyxl-internal API uses (`openpyxl.utils.bound_dict`, etc.) that wolfxl deliberately doesn't replicate → mark xfail with reason.
  - openpyxl-side test-infra dependencies (lxml comparisons, fixture paths) → adapt or skip with reason.

**Pass target:** ≥99% on the vendored corpus, with documented xfails for the remaining <1%.

**Estimate:** unbounded; budget 3 sessions for first pass + 2-4 follow-up sessions to fix surfaced bugs.

### Phase 12 — Empty `limitations.md` + README claim flip (1 session)

**Goal:** ship the user's stated outcome.

**Files:**
- `docs/trust/limitations.md`: delete the file, OR reduce to a single sentence: *"None known. See [compatibility-matrix.md](../migration/compatibility-matrix.md)."*
- `README.md`: change "openpyxl-compatible" → "drop-in replacement for openpyxl" with a one-line cite of the three-measurement gate (compat oracle pass-%, dynamic-surface 23 classes, vendored corpus pass-%).
- `docs/index.md` and any other public-facing claim docs: update language.
- `Plans/openpyxl-parity-program.md`: mark all gaps (incl. G07/G21-G24 and the G17/G18 follow-ups) as landed; flip program status to "complete".

**Verification:** current user-facing docs (`README.md`, `docs/index.md`, `docs/migration/`, `docs/trust/`) no longer describe openpyxl-supported surfaces as `partial`, `not_yet`, `deferred`, or known limitations, except where historical release notes or the parity-program plan intentionally preserve project history.

## Critical files (cross-reference)

Frequently-modified shared files across phases:

- `docs/migration/_compat_spec.py` — every phase touches.
- `tests/test_openpyxl_compat_oracle.py` — every phase touches (adds probes).
- `crates/wolfxl-writer/src/emit/` — phases 2, 3, 4, 5, 8.
- `crates/wolfxl-reader/src/` — phases 2, 3, 4, 5, 6, 7, 8.
- `src/wolfxl/patcher_workbook.rs` — phases 2, 5, 6, 8.
- `src/wolfxl/patcher_sheet_blocks.rs` — phases 3, 4.
- `src/wolfxl/patcher_pivot_edit.rs` — phase 7.
- `python/wolfxl/pivot/_handle.py` — phase 7.
- `python/wolfxl/_external_links.py` — phase 8.

## Verification

**Per-phase gate** (green before next phase starts):

1. `cargo test --workspace` clean.
2. `uv run --no-sync maturin develop --release` rebuilds.
3. `uv run --no-sync pytest -q` — full Python suite no regressions (current baseline 2534/0/41/3).
4. `uv run --no-sync pytest tests/test_external_oracle_preservation.py -v` — 7/7 fixtures.
5. `uv run --no-sync pytest tests/test_openpyxl_compat_oracle.py -v` — pass count up by the phase's probe count, no new xfails.
6. `uv run --no-sync pytest tests/test_openpyxl_dynamic_surface.py -v` — green (expanding scope phase 9).
7. New phase-specific tests added pass.

**Program-completion gate** (green before Phase 12 ships):

1. Compat oracle: 0 xfail / 0 fail. Spec: 0 partial, 0 not_yet (out_of_scope rows clearly tagged "wolfxl-extra, openpyxl doesn't support").
2. Dynamic-surface ratchet: 23 classes, zero missing/mismatch.
3. Vendored corpus: ≥99% pass with documented xfails for the rest.
4. `docs/trust/limitations.md` empty or deleted.
5. README claim says "drop-in replacement for openpyxl" honestly.

## Calendar

| Phase | Sessions | Owner mode | Dependencies | Status |
|---|---|---|---|---|
| 1 — Recategorize wolfxl-extras | 1 | Claude (mechanical) | none | LANDED 2026-05-04 |
| 2 — G22 defined-name attrs | 1 | Codex 🤖 | Phase 1 | LANDED 2026-05-04 |
| 3 — G24 print settings depth | 2 | Codex 🤖 | Phase 1 | LANDED 2026-05-04 |
| 4 — G07 spill metadata | 2 | Claude 🧠 | Phase 1 | |
| 5 — G21 standalone slicers | 2 | Claude 🧠 | Phase 1 | |
| 6 — G23 calc-chain edges | 3 | Claude 🧠 | Phase 1 | |
| 7 — G17 pivot full mutation | 4 | Claude 🧠 (RFC-070 v2) | Phase 1 | |
| 8 — G18 external-link authoring | 2 | Claude 🧠 (RFC-071 v2) | Phase 1 | |
| 9 — CF audit + ratchet expansion | 2 | Claude 🧠 | none | |
| 10 — Probe coverage debt | 1 | Codex 🤖 | Phase 1 | |
| 11 — Vendor openpyxl corpus | 3-5 | Claude 🧠 + bug-fix follow-ups | Phases 2-10 (so corpus runs against the closed-gap tree) | |
| 12 — Empty limitations + claim flip | 1 | Claude (mechanical) | Phases 1-11 | |

**Total:** 24-26 sessions. Phases 2/3 and 5/9/10 can run in parallel worktrees if pod-dispatched.

## Out of scope

- `.xls` / `.xlsb` / `.ods` writes, VBA authoring, `.xls` style reads. These are wolfxl-extras (openpyxl matches wolfxl on lacking these) and are tracked in `Plans/openpyxl-parity-program.md` G25-G28 as decision-gated wolfxl-roadmap items, not openpyxl-parity gaps.
- Refactoring beyond what each phase requires (the `pre-release-expanded-oracle-and-cleanup.md` track owns refactors).
- The pre-release release-freeze flag — separate concern; this plan's completion supersedes it because the freeze was waiting on parity oracle work.

## Risks

| # | Risk | Mitigation |
|---|---|---|
| 1 | Phase 7 (pivot mutation) is the single largest LOC item; budget overrun likely. | Stage in 4 sub-PRs (field placement, filter, aggregation, regeneration); each closeable independently with its own probe. |
| 2 | Phase 11 vendored corpus surfaces 100+ failures; triage takes longer than budgeted. | Cap fix-effort at 5 follow-up sessions; remaining failures get documented xfails with openpyxl-version pin so future sessions can target them. |
| 3 | Phase 9 dynamic-surface expansion exposes ~50 missing methods/attrs across 20 classes. | Treat the surface expansion as a discovery phase; each `missing_callable` either fixed inline or routed back to a phase-2-style follow-up. |
| 4 | "limitations.md empty" temptation to redefine "limitation" rather than fix the gap. | Spec entries must be `supported` (with passing probe) to count toward empty; renaming a `partial` to `out_of_scope` requires a new RFC documenting why openpyxl-side equivalence holds. |
| 5 | Codex pods on Phases 2/3/10 mass-merge into shared files (`_compat_spec.py`, `oracle.py`). | Use the existing `serial-merge-after-N-parallel` pattern from S3; one pod merges, sibling pods rebase before merging. |

## Changelog

- 2026-05-04: Plan created. Phase 1 landed (commit 4825baf).
- 2026-05-04: Phase 2 landed on branch `parity-closure-phase2`; Codex review tightened the hidden-clear regression and updated stale limitations/tracker text.
- 2026-05-04: Phase 3 landed on branch `parity-closure-phase2`; G24 print settings depth moved to supported and the limitations row was removed.
- 2026-05-04: Phase 4 landed via PR #27; array formula and data-table formula surface moved to supported, with openpyxl 3.1.5 `Worksheet.array_formulae` treated as the reference surface.
- 2026-05-04: Phase 10 landed via PR #28; compatibility-oracle coverage expanded, leaving `coordinate_to_tuple` strictness as the remaining probe-discovered partial.
- 2026-05-05: Remaining-limitations closure slice tightened `coordinate_to_tuple`, added calc-chain edge-case coverage (`extLst` preservation, stale ref pruning, cross-sheet formulas), and reclassified standalone slicers as out-of-scope parity based on openpyxl 3.1.5 source evidence.
- 2026-05-04: Phase 9 implemented locally on `parity-closure-phase9-cf-ratchet`; CF rare-combo probes, modify-mode CF attr tests, and a 22-object dynamic-surface ratchet are green locally.
