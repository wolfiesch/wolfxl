# Pre-Release Expanded Oracle and Cleanup Sprint

Date: 2026-04-28
Status: Proposed pre-release hardening track.

## Release freeze

Do not publish, tag, release, or post publicly from this track until the
expanded oracle pass and cleanup pass have been audited. The current local
openpyxl parity result is strong enough to continue hardening, but it should
not be treated as sufficient public-release evidence by itself.

## Goal

Close the risk that openpyxl is too narrow as the only behavioral oracle, then
reduce implementation risk in the largest WolfXL modules before v2.0 is pushed
or announced.

This is deliberately a multi-sprint plan. The objective is production-grade
confidence, not a fast release.

## Research snapshot

ExcelBench already covers a broad Python/Rust spreadsheet ecosystem:
openpyxl, xlsxwriter, python-calamine/calamine, rust_xlsxwriter, pylightxl,
pyexcel, pandas, polars/fastexcel, tablib, xlrd/xlwt, umya/pyumya, and
WolfXL.

The highest-signal additions are not more Python value-only libraries. They
are cross-runtime writers/readers and renderer-oracles that can produce or
validate OOXML structures openpyxl does not construct deeply.

| Priority | Tool | Runtime | Why it matters | Initial scope |
|---|---|---:|---|---|
| P0 | Excelize | Go | Full read/write library with active support for charts, slicers, pivot tables, rich formatting, images, and streaming. It can create fixtures openpyxl does not construct. | External subprocess adapter that writes/reads JSON-described `.xlsx` fixtures. |
| P0 | LibreOffice Calc headless | CLI / UNO | Not a library replacement, but the best open-source "will this workbook open/render/round-trip" oracle. | Open/save smoke, PDF/export smoke for visual corruption, repair-dialog detection where possible. |
| P1 | Apache POI | Java | Mature OOXML/XLS implementation with usermodel, streaming XSSF/SXSSF, pivot-table examples, conditional formatting, pictures, charts, and explicit documented limitations. | External Java adapter for fixture generation plus readback metadata. |
| P1 | ClosedXML | .NET | High-level workbook API with table, pivot cache/table, conditional formatting, and rich cell model. Useful as a developer-ergonomics and OOXML semantics cross-check. | External .NET adapter for pivots, CF, tables, rich text, comments, and protections. |
| P2 | NPOI | .NET | POI-derived .NET surface. Potentially useful if .NET users expect POI-like behavior, but likely overlaps with Apache POI and ClosedXML. | Investigate after POI/ClosedXML are wired. |
| P2 | SheetJS CE | JavaScript | Strong broad-format data toolkit, but advanced styling/images/charts/pivots are positioned as Pro features rather than CE. | Value/formula/import-export sanity only unless CE evidence shows richer support. |
| P2 | libxlsxwriter / rust_xlsxwriter | C / Rust | Excellent writer reference for charts, conditional formatting, tables, images, and byte-level Excel fixture discipline. ExcelBench already has rust_xlsxwriter; keep as a targeted fixture source. | Use for focused low-level writer comparisons, not as a new external adapter. |
| P3 | pyxlsb / ODS tools | Python | Useful for non-xlsx read/format scope, but not a v2.0 xlsx publish blocker. | Later format-expansion track. |

License caution:

- Treat EPPlus as out of scope for an "open-source oracle" until we explicitly
  decide whether its current licensing is acceptable for this project.
- Treat Aspose.Cells and Syncfusion as proprietary/commercial references, not
  open-source benchmark targets.

## Features openpyxl may not cover enough

The expanded oracle fixtures should target capabilities where openpyxl either
round-trips more than it constructs, validates only shallowly, or normalizes
away structural detail:

- Pivot caches with saved records, calculated fields/items, filters, grouping,
  styles, and refresh-on-open flags.
- Pivot-linked slicers and slicer-cache OOXML.
- Pivot charts and chart objects linked to pivot ranges.
- Combination charts, per-point overrides, data labels, display units, trend
  lines, and chart sheets.
- Rich text runs in shared strings, inline strings, comments, headers, footers,
  and chart labels.
- Conditional formatting beyond basic rules: icon sets, data bars, color
  scales, stop-if-true, priority ordering, differential styles, and pivot-scoped
  conditional formats.
- Tables with totals formulas, structured references, style info, filters, and
  table preservation during modify-mode saves.
- Drawings and images with one-cell/two-cell/absolute anchors, alt text,
  names, hyperlinks, and EMU conversions.
- Workbook and sheet protection, workbook security metadata, macro preservation
  for `.xlsm`, external links, defined-name edge cases, data tables, array
  formulas, calc-chain behavior, print settings, and page setup.

## Harness plan

Add an "external oracle" layer in ExcelBench rather than baking Go/Java/.NET
dependencies into WolfXL itself.

1. Define a subprocess adapter contract:
   - Input: JSON fixture manifest with sheets, cells, styles, formulas, tables,
     pivots, charts, drawings, validations, and expected feature probes.
   - Output: created workbook path plus JSON readback metadata and diagnostics.
   - Failure mode: structured skip if the runtime/tool is unavailable.
2. Start with write fixtures, because competing tools can create complex OOXML
   that WolfXL can inspect, modify, and re-save.
3. Add readback probes only after write fixtures are stable.
4. Keep external-oracle outputs out of public benchmark claims until manually
   audited. Use a local results directory such as `results_dev_external/`.
5. Promote only stable, deterministic cases into checked-in ExcelBench fixtures.

Initial adapters:

- `excelize_external`: Go helper binary.
- `libreoffice_external`: CLI validator, not a normal benchmark competitor.
- `poi_external`: Maven/Gradle helper.
- `closedxml_external`: `dotnet` helper.

Implementation checkpoint, 2026-04-28:

- ExcelBench commit `91097bf` adds the initial Excelize external-oracle helper
  under `tools/external-oracles/excelize`.
- ExcelBench commit `e988eed` adds the initial LibreOffice external-oracle
  helper under `tools/external-oracles/libreoffice`.
- The first generated Excelize smoke workbook includes table, pivot cache,
  pivot table, slicer, slicer cache, chart, drawing, and picture parts.
- Local truth pass: openpyxl opened the workbook with the expected unsupported
  extension warning for slicer metadata, WolfXL read the expected cell values,
  and LibreOffice headless exported the workbook to PDF without stderr.
- Follow-up mining found a real WolfXL bug: modify-mode saves back to the
  original source path called the normal patcher `save(path)` path and could
  trip a ZIP checksum read error. The fix routes same-path modify saves through
  the patcher's atomic `save_in_place()` path and adds a regression test.
- ExcelBench commit `2de89d2` adds a repeatable local fixture-pack generator:
  `uv run python scripts/generate_external_oracle_fixtures.py`.
- ExcelBench commit `133d3ab` adds a WolfXL preservation validator for those
  generated external fixtures.
- The first ClosedXML truth pass produced a second real WolfXL bug: prefixed
  worksheet XML such as `<x:c>` received unprefixed inserted cells during
  modify-mode saves. Openpyxl ignored the un-namespaced marker even though
  WolfXL's permissive reader saw it. The fix teaches the stream patcher to
  preserve the worksheet element prefix for inserted/replaced cells and adds a
  regression test with a prefixed namespace worksheet.
- ExcelBench commit `d2c21ba` adds the NPOI fixture generator, commit
  `623de0e` adds the ExcelJS fixture generator, and commit `dee7cc0` adds the
  Apache POI fixture generator with pinned Maven Central dependency downloads.
- Current external fixture pack: seven workbooks from Excelize, ClosedXML,
  NPOI, ExcelJS, and Apache POI. The 2026-04-28 local truth pass validated the
  full pack with LibreOffice headless open/save, LibreOffice PDF export, WolfXL
  read, and WolfXL in-place modify-save preservation.
- WolfXL now has a local pre-release gate in
  `tests/test_external_oracle_preservation.py`. It runs against
  `WOLFXL_EXTERNAL_FIXTURES_DIR` or the sibling ExcelBench generated fixture
  directory when present, and skips cleanly in isolated checkouts.

## Cleanup plan

Refactors should be characterization-test driven. Avoid broad behavior-free
rewrites; each split needs proof that public Python APIs, PyO3 bindings, and
OOXML output stay stable.

Current largest WolfXL hotspots:

| Module | Current LOC | Cleanup direction |
|---|---:|---|
| `src/wolfxl/mod.rs` | 2503 | Continue splitting patcher phases and save-path orchestration behind the same PyO3 surface. |
| `src/calamine_styled_backend.rs` | 4967 | Split reader extraction into styles, hyperlinks, comments, drawings, tables, conditional formatting, and validations modules. |
| `src/native_writer_backend.rs` | 529 | Continue splitting the remaining Python-to-writer bridge into focused helper modules while keeping the PyO3 surface stable. |
| `python/wolfxl/_worksheet.py` | 1513 | Continue extracting pending-flush helpers and feature-specific collections while preserving openpyxl-shaped imports. |
| `crates/wolfxl-writer/src/emit/sheet_xml.rs` | 386 | Keep as the CT_Worksheet coordinator with minimal full-sheet ordering and well-formedness coverage. |
| `python/wolfxl/_workbook.py` | 1321 | Continue separating workbook orchestration from calculation, lifecycle, feature registration, and save pipeline helpers. |

Suggested sprint sequence:

1. **Oracle Sprint A: research-to-fixtures.**
   Add Excelize and LibreOffice external harness scaffolding, with 5-8 fixtures
   that stress pivots, slicers, charts, CF, tables, drawings, and protection.
2. **Cleanup Sprint B: Python API docstrings and public docs.**
   Add Google-style docstrings to public Python APIs and fill obvious gaps in
   examples. Keep private helper docstrings sparse unless they clarify behavior
   that tests depend on.
3. **Cleanup Sprint C: `_worksheet.py` and `_workbook.py` split.**
   Extract feature collections and pending-flush helpers with focused unit
   tests and import-compatibility checks.
4. **Oracle Sprint D: POI and ClosedXML adapters.**
   Add Java/.NET adapters only after the subprocess contract has settled.
5. **Cleanup Sprint E: Rust writer/parser split.**
   Split `native_writer_backend.rs` and `sheet_xml.rs`, then rerun native
   writer, diffwriter, and ExcelBench parity gates.
6. **Cleanup Sprint F: calamine reader and patcher split.**
   Split the reader and `src/wolfxl/mod.rs` after the external oracle corpus is
   broad enough to catch drift.

### Rust `src/wolfxl/mod.rs` split entry plan

Treat `src/wolfxl/mod.rs` as the highest-risk cleanup target because it owns
the PyO3 entrypoints, patcher queues, and workbook mutation ordering in one
file. The first split should reduce file size without changing exported names
or OOXML write order.

Phase 0: inventory and characterization.

- Map every `#[pyclass]`, `#[pymethods]`, and `#[pyfunction]` item, plus each
  queue type consumed by `XlsxPatcher.save`, before moving code.
- Pin a focused behavior baseline with `cargo test`, `uv run --no-sync maturin
  develop`, modify-mode Python tests, and the external oracle preservation
  gate.
- Current known-good evidence before this split track: WolfXL full suite
  `2285 passed, 29 skipped`, and the ExcelBench external fixture validator
  passed the seven-workbook pack with `55` readback probes and no failures.

Phase 1: extract pure models first.

- Move queue payload structs and small helper enums into a private Rust module
  such as `src/wolfxl/patcher_models.rs` or `src/wolfxl/patcher/queues.rs`.
- Keep PyO3 class/function definitions and `XlsxPatcher` method signatures in
  place until the compiler proves the model split is mechanical.
- Avoid path/order churn in emitted ZIP parts unless byte-level regression tests
  prove the change is intentional.

Phase 2: split patcher flush phases one group at a time.

- Start with isolated queue drainers: cells/formats, relationships/drawings,
  worksheet setup, workbook metadata, and then pivots/slicers.
- Each extracted phase should expose a small Rust function that receives the
  existing patcher state rather than inventing new ownership boundaries.
- After each phase, rerun the focused tests for the touched feature plus
  `tests/test_external_oracle_preservation.py` when the ExcelBench fixture pack
  is present.

Phase 3: only then reduce the PyO3 surface file.

- Once phase helpers are stable, move internal implementation blocks out of
  `mod.rs` while keeping public Python import behavior unchanged.
- Stop immediately if a move forces a PyO3 signature change, changes save-path
  ordering, or requires broad fixture rewrites to explain output drift.

Phase 0 inventory snapshot, 2026-04-28:

- PyO3 class surface in this file is currently a single `#[pyclass]`:
  `XlsxPatcher`.
- Public constructor/save surface: `open`, `sheet_names`, `save`, and
  `save_in_place`.
- Public queue surface: `queue_value`, `queue_rich_text_value`,
  `queue_array_formula`, `queue_format`, `queue_border`,
  `queue_data_validation`, `queue_conditional_formatting`,
  `queue_hyperlink`, `queue_hyperlink_delete`, `queue_comment`,
  `queue_comment_delete`, `queue_table`, `queue_image_add`,
  `queue_chart_add`, `queue_pivot_cache_add`, `queue_pivot_table_add`,
  `queue_slicer_add`, `queue_autofilter`, `queue_sheet_setup_update`,
  `queue_page_breaks_update`, `queue_workbook_security`,
  `queue_defined_name`, `queue_sheet_move`, `queue_axis_shift`,
  `queue_range_move`, `queue_sheet_copy`, and `queue_properties`.
- Test-only PyO3 helpers still live on the public impl block:
  `_test_inject_file_add`, `_test_queue_content_type_op`,
  `_test_populate_ancillary`, `_test_ancillary_is_populated`,
  `_test_ancillary_comments_part`, `_test_ancillary_vml_drawing_part`,
  `_test_ancillary_table_parts`, `_test_ancillary_hyperlink_rids`,
  `_test_inject_hyperlink`, `_test_inject_hyperlink_delete`, and
  `_test_get_extracted_hyperlinks`.
- Queue/model structs still local to `mod.rs`: `QueuedChartAdd`,
  `QueuedImageAdd`, `QueuedImageAnchor`, `SheetCopyOp`, `AxisShift`, and
  `RangeMove`. These are the safest first extraction candidates because they
  are pure Rust data carriers with no PyO3 annotations.
- Save-phase helpers already have natural seams inside the private impl:
  `apply_image_adds_phase`, `apply_chart_adds_phase`,
  `apply_pivot_adds_phase`, `apply_slicer_adds_phase`,
  `apply_axis_shifts_phase`, `apply_sheet_copies_phase`,
  `apply_range_moves_phase`, `rebuild_calc_chain_phase`, and
  `ensure_calc_chain_metadata`.
- Bottom-of-file parser/render helpers were the second split track after pure
  queue models: rich-text, format, border, conditional-format,
  workbook-security, image-anchor, and chart drawing helpers now live in
  focused patcher modules.

First no-behavior split target, completed 2026-04-28:

1. Extract the six local queue/model types into `src/wolfxl/patcher_models.rs`
   and re-export them privately into `mod.rs`. Commit target: pure Rust data
   carriers only, with no PyO3 annotation or save-order changes.
2. Verification completed: `cargo test`, `uv run --no-sync maturin develop`,
   focused modify/oracle tests, and the full WolfXL suite.
3. Images/charts parser and drawing-render helpers moved into
   `src/wolfxl/patcher_drawing.rs` on 2026-04-28. Verification completed:
   `cargo test`, `uv run --no-sync maturin develop`, and focused
   image/chart/external-oracle Python tests.
4. Rich-text, format, border, conditional-format, workbook-security, and
   generic dict payload parsers moved into `src/wolfxl/patcher_payload.rs` on
   2026-04-28. Verification completed: `cargo test`, `uv run --no-sync
   maturin develop`, and focused formatting/CF/security/rich-text/external
   oracle Python tests.
5. Workbook relationship helpers, current-part byte loaders, sheet-copy
   workbook splicing, deterministic ZIP timestamps, and the minimal styles
   fallback moved into `src/wolfxl/patcher_workbook.rs` on 2026-04-28.
   Verification completed: `cargo test`, `uv run --no-sync maturin develop`,
   and focused copy-worksheet/move-range/properties/patcher-infra/external
   oracle Python tests.
6. The sheet-copy save phase moved into `src/wolfxl/patcher_sheet_copy.rs` on
   2026-04-28 while keeping the `mod.rs` call site and phase ordering intact.
   Verification completed: `cargo test`, `uv run --no-sync maturin develop`,
   and focused copy-worksheet/external-oracle Python tests.
7. Axis-shift and range-move save phases moved into
   `src/wolfxl/patcher_structural.rs` on 2026-04-28. Verification completed:
   `cargo test`, `uv run --no-sync maturin develop`, and focused
   axis-shift/move-range/external-oracle Python tests.
8. Pivot and slicer save phases moved into `src/wolfxl/patcher_pivot.rs` on
   2026-04-28 while preserving the `mod.rs` wrapper methods and the existing
   Phase 2.5m / 2.5p ordering.
9. Image and chart save phases moved into `src/wolfxl/patcher_drawing.rs` on
   2026-04-28, co-located with the existing drawing XML helper functions while
   keeping the `mod.rs` wrapper methods and Phase 2.5k / 2.5l ordering intact.
10. Calc-chain rebuild and metadata save phases moved into
   `src/wolfxl/patcher_workbook.rs` on 2026-04-28 while keeping the `mod.rs`
   rebuild wrapper and Phase 2.8 ordering intact.
11. Content-types aggregation and document-properties save phases moved into
   `src/wolfxl/patcher_workbook.rs` on 2026-04-28 while keeping the `mod.rs`
   wrapper methods and Phase 2.5c / 2.5d ordering intact.
12. Sheet-setup and page-break block phases moved into
   `src/wolfxl/patcher_sheet_blocks.rs` on 2026-04-28 while keeping the
   `mod.rs` call sites and Phase 2.5n / 2.5r ordering intact.
13. AutoFilter block and hidden-row computation moved into
   `src/wolfxl/patcher_sheet_blocks.rs` on 2026-04-28 while keeping the
   `mod.rs` call site and Phase 2.5o ordering intact.
14. Workbook security, sheet-reorder, and defined-name workbook XML phases
   moved into `src/wolfxl/patcher_workbook.rs` on 2026-04-28 while preserving
   the shared `workbook_xml_in_progress` handoff and Phase 2.5q / 2.5h /
   2.5f ordering.
15. Worksheet XML patching pass moved into
   `src/wolfxl/patcher_sheet_blocks.rs` on 2026-04-28 while preserving the
   cell-patch, sibling-block, autoFilter hidden-row, and cloned-sheet
   `file_adds` routing sequence.
16. Cell style preparation and sheet cell-patch construction moved into
   `src/wolfxl/patcher_cells.rs` on 2026-04-28 while preserving the shared
   `styles_xml` handoff to conditional-formatting dxf aggregation.
17. Data-validation and conditional-formatting sheet-block phases moved into
   `src/wolfxl/patcher_sheet_blocks.rs` on 2026-04-28 while preserving
   existing-block extraction, deterministic CF sheet ordering, and the shared
   `styles_xml` dxf handoff.
18. Hyperlink sheet-block and sheet-rels phase moved into
   `src/wolfxl/patcher_sheet_blocks.rs` on 2026-04-28 while preserving
   ancillary population, existing hyperlink extraction, and empty-block
   deletion semantics.
19. Table sheet-block, table-part, rels, and content-type phase moved into
   `src/wolfxl/patcher_sheet_blocks.rs` on 2026-04-28 while preserving cloned
   table collision checks, shared part-id allocation, table inventory updates,
   and cloned-sheet rels resolution.
20. Comments and VML sheet-block phase moved into
   `src/wolfxl/patcher_sheet_blocks.rs` on 2026-04-28 while keeping final
   comment/VML part routing in `mod.rs` and preserving author dedup,
   existing-part reuse, content-type ops, and legacyDrawing deletion semantics.
21. Relationship serialization, generic part-write/delete routing, and final
   ZIP rewrite moved into `src/wolfxl/patcher_workbook.rs` on 2026-04-28 while
   preserving in-place-vs-add routing, source-entry collision checks, source
   entry metadata, deterministic added-entry ordering, and
   `WOLFXL_TEST_EPOCH` timestamps.
22. No-op save detection, source-file copy, and permissive seed draining moved
   into `src/wolfxl/patcher_workbook.rs` on 2026-04-28 while preserving
   byte-identical no-op saves and one-shot permissive workbook XML rewrites.
23. Public API docstring inventory for `python/wolfxl/_workbook.py`,
   `python/wolfxl/_worksheet.py`, `python/wolfxl/_cell.py`, and
   `python/wolfxl/_streaming.py` completed on 2026-04-28 by adding concise
   Google-style docstrings to remaining public property setters, core cell
   accessors, streaming-cell accessors, and debug representations.
24. Simple modify-mode feature drains for hyperlinks, tables, images,
   comments, data validations, and conditional formatting moved into
   `python/wolfxl/_workbook_patcher_flush.py` on 2026-04-28 while preserving
   the existing `_workbook.py` wrapper methods and save-order call sites.
25. Structural move/copy, defined-name, workbook-security, and document-property
   modify-mode drains moved into `python/wolfxl/_workbook_patcher_flush.py` on
   2026-04-28 while preserving the existing `_workbook.py` wrapper methods,
   save-order call sites, dirty-flag resets, and one-shot queue drains.
26. Sheet-setup, page-break, sheet-format, and autofilter modify-mode drains
   moved into `python/wolfxl/_workbook_patcher_flush.py` on 2026-04-28 while
   preserving wrapper methods, defensive malformed-wrapper skips, and
   no-op-save byte stability probes.
27. Chart, slicer, and pivot modify-mode drains moved into
   `python/wolfxl/_workbook_patcher_flush.py` on 2026-04-28 while preserving
   wrapper methods, serializer import timing, cache id sanity checks, and
   one-shot chart/pivot/slicer queue drains.
28. Native-writer workbook metadata, defined-name, and security flush moved into
   `python/wolfxl/_workbook_writer_flush.py` on 2026-04-28, and eager
   modify-mode sheet-move queueing moved into `_workbook_patcher_flush.py`,
   while preserving writer payload shapes and wrapper methods.
29. Worksheet append and bulk-write buffer materialization helpers moved into
   `python/wolfxl/_worksheet_write_buffers.py` on 2026-04-28 while preserving
   lazy append materialization, patcher-path bulk materialization, and
   non-batchable per-cell extraction semantics. The format/border bounding-box
   dict batch helper now lives in the same module.
30. Worksheet modify-mode dirty cell, format, rich-text, array-formula, and
   spill-child placeholder flushing moved into
   `python/wolfxl/_worksheet_patcher_flush.py` on 2026-04-28 while preserving
   the existing `_worksheet.py` wrapper and patcher queue payload shapes.
31. Worksheet write-mode append buffer, bulk write, dirty value, rich-text,
   array-formula, spill-child placeholder, and format/border flushing moved into
   `python/wolfxl/_worksheet_writer_flush.py` on 2026-04-28 while preserving
   the existing `_worksheet.py` wrapper and native writer payload shapes.
32. Worksheet record iteration, pending-overlay patching, cached formula values,
   sheet visibility caching, and schema inference moved into
   `python/wolfxl/_worksheet_records.py` on 2026-04-28 while preserving public
   method docstrings and the legacy `_worksheet._canonical_data_type` import
   path used by existing callers/tests.
33. Worksheet openpyxl-style cell/range access helpers moved into
   `python/wolfxl/_worksheet_access.py` on 2026-04-28 while preserving
   `Worksheet.__getitem__`, `_resolve_string_key`, and row/column rectangle
   wrapper methods.
34. Worksheet row/column iteration, read-mode values-only bulk reads, and
   streaming dispatch moved into `python/wolfxl/_worksheet_iteration.py` on
   2026-04-28 while preserving `Worksheet.iter_rows`, `iter_cols`,
   `_iter_rows_bulk`, and `_iter_cols_bulk` wrapper methods.
35. Worksheet structural operation validation and queueing for row/column
   shifts plus `move_range` moved into `python/wolfxl/_worksheet_structural.py`
   on 2026-04-28 while preserving the public worksheet method docstrings and
   pending patcher queue payload shapes.
36. Worksheet chart, pivot table, slicer, A1 anchor validation, chart
   removal/replacement, and image queueing moved into
   `python/wolfxl/_worksheet_media.py` on 2026-04-28 while preserving the
   public worksheet method docstrings and pending writer/patcher payload
   queues.
37. Worksheet freeze panes, page setup/margins, header/footer, sheet view,
   protection, page breaks, sheet format, print titles, dimension holder, and
   setup/page-break Rust payload helpers moved into
   `python/wolfxl/_worksheet_setup.py` on 2026-04-29 while preserving the
   public worksheet property/method docstrings and serializer payload shapes.
38. Worksheet used-range, dimension bounds, read-mode dimension discovery, and
   `max_row` / `max_column` helpers moved into
   `python/wolfxl/_worksheet_bounds.py` on 2026-04-29 while preserving the
   public worksheet wrapper methods and pending-write bounds semantics.
39. Worksheet table and data-validation write-side collection helpers moved
   into `python/wolfxl/_worksheet_features.py` on 2026-04-29, placing them
   alongside the existing lazy comments, hyperlinks, tables, validations, and
   conditional-formatting loaders while preserving wrapper method behavior.
40. Workbook sheet create/remove/copy/move helpers moved into
   `python/wolfxl/_workbook_sheets.py` on 2026-04-29 while preserving public
   method signatures, error messages, copy-option capture, print-title cloning,
   and write/modify-mode routing.
41. Shared xlsx read/modify/bytes constructor bookkeeping moved into
   `build_xlsx_wb()` and shared worksheet-proxy/pending-queue initializers in
   `python/wolfxl/_workbook_state.py` on 2026-04-29 while preserving source
   path, read-only, data-only, patcher, and pending mutation semantics.
42. Workbook document-properties, defined-name, protection, and file-sharing
   backing helpers moved into `python/wolfxl/_workbook_metadata.py` on
   2026-04-29 while preserving lazy cache hydration, dirty flags, user-write
   queueing, and openpyxl-shaped type errors.
43. Workbook-level chart, pivot-cache, and slicer-cache registration helpers
   moved into `python/wolfxl/_workbook_features.py` on 2026-04-29 while
   preserving modify-mode guards, id allocation, cache materialization, and
   pending queue semantics.
44. Writer worksheet data-validation XML emission moved into
   `crates/wolfxl-writer/src/emit/data_validations.rs` on 2026-04-29 while
   preserving CT_Worksheet slot ordering between conditional formatting and
   hyperlinks plus formula/operator/error-style serialization.
45. Writer worksheet hyperlink XML emission moved into
   `crates/wolfxl-writer/src/emit/hyperlinks.rs` on 2026-04-29 while
   preserving internal/external target routing, optional display/tooltip
   attributes, and the comments/tables/external-hyperlink rId convention.
46. Writer worksheet merged-cell XML emission moved into
   `crates/wolfxl-writer/src/emit/merges.rs` on 2026-04-29 while preserving
   deterministic top-left range sorting and the empty-merge omission path.
47. Writer worksheet table-part XML emission moved into
   `crates/wolfxl-writer/src/emit/table_parts.rs` on 2026-04-29 while
   preserving comments/VML rId offsets and sheet-local table ordering.
48. Writer worksheet drawing and legacyDrawing relationship refs moved into
   `crates/wolfxl-writer/src/emit/drawing_refs.rs` on 2026-04-29 while
   preserving comments/VML, table, and external-hyperlink rId allocation.
49. Writer worksheet dimension XML emission moved into
   `crates/wolfxl-writer/src/emit/dimension.rs` on 2026-04-29 while preserving
   blank-cell exclusion and styled-blank inclusion in the emitted bounds.
50. Writer worksheet legacy sheet-view XML emission moved into
   `crates/wolfxl-writer/src/emit/sheet_views.rs` on 2026-04-29 while
   preserving selected-first-sheet, freeze-pane, and split-pane behavior.
51. Writer worksheet column XML emission moved into
   `crates/wolfxl-writer/src/emit/columns.rs` on 2026-04-29 while preserving
   width, hidden, outline, style, and default-column omission behavior.
52. Writer worksheet row and cell XML emission moved into
   `crates/wolfxl-writer/src/emit/sheet_data.rs` on 2026-04-29 while
   preserving sheetData, row-attribute, SST, formula, rich-text, and spill-cell
   serialization.
53. Writer worksheet conditional-formatting XML emission moved into
   `crates/wolfxl-writer/src/emit/conditional_formats.rs` on 2026-04-29 while
   preserving supported rule serialization and all-stub wrapper omission.
54. Native writer Python cell-value coercion moved into
   `src/native_writer_cells.rs` on 2026-04-29 while preserving payload dict
   conversion, raw `write_sheet_values` coercion, boolean precedence,
   date/datetime serials, formula stripping, and non-finite float rejection.
55. Native writer format, border, and color-normalization payload parsing moved
   into `src/native_writer_formats.rs` on 2026-04-29 while preserving writer
   format/style interning, border-only style interning, and conditional-format
   color normalization.
56. Native writer rich-text run payload parsing moved into
   `src/native_writer_rich_text.rs` on 2026-04-29 while preserving inline
   rich-text tuple coercion, font-property parsing, and non-finite rich-text
   integer-field rejection.
57. Native writer sheet-feature payload parsing for hyperlinks, comments,
   conditional formats, data validations, and tables moved into
   `src/native_writer_sheet_features.rs` on 2026-04-29 while preserving
   optional wrapper unwrapping, silent malformed-payload no-ops, dxf
   interning, author interning, and table column/style payload conversion.
58. Native writer workbook metadata and security payload parsing moved into
   `src/native_writer_workbook_metadata.rs` on 2026-04-29 while preserving
   defined-name scoping, document-property field mapping, created-date parsing,
   and workbookProtection/fileSharing dict conversion.
59. Native writer array/data-table formula payload parsing moved into
   `src/native_writer_cells.rs` on 2026-04-29 while preserving payload
   validation errors, leading-equals stripping, data-table flags, references,
   and spill-child placeholder handling.
60. Shared native writer drawing-anchor payload parsing moved into
   `src/native_writer_anchors.rs` on 2026-04-29 while preserving image and
   chart one-cell, two-cell, and absolute anchor dict semantics.
61. Native writer image payload parsing moved into
   `src/native_writer_images.rs` on 2026-04-29 while preserving data/ext/size
   extraction, lowercase extension normalization, anchor dict validation, and
   sheet image queueing.
62. Native writer autofilter installation moved into
   `src/native_writer_autofilter.rs` on 2026-04-29 while preserving
   dict-value parsing, XML emission, row-hidden reset scope, evaluator grid
   conversion, and hidden-row stamping.
63. Native writer freeze/split pane settings moved into
   `src/native_writer_sheet_state.rs` on 2026-04-29 while preserving wrapper
   handling, default freeze mode, A1 validation, and split coordinate defaults.
64. Native writer chart dict parsing and the module-level
   `serialize_chart_dict` PyO3 export moved into `src/native_writer_charts.rs`
   on 2026-04-29 while preserving add-chart, modify-mode chart XML, pivot
   source, anchor, axis, series, data-label, error-bar, trendline, and 3D chart
   payload semantics.
65. Native writer worksheet state application in
   `src/native_writer_sheet_state.rs` expanded on 2026-04-29 to cover row and
   column dimensions, merged ranges, print areas, sheet setup blocks, page
   breaks, and sheet format blocks while preserving existing PyO3 method
   names.
66. Native writer style application moved into `src/native_writer_formats.rs`
   on 2026-04-29 while preserving cell and grid format/border style-id
   interning, blank-cell creation for styled empty cells, and save-path style
   output.
67. Native writer cell write routing moved into `src/native_writer_cells.rs`
   on 2026-04-29 while preserving typed payload conversion, date/datetime
   default number formats, rich-text inline strings, array/data-table/spill
   formulas, raw grid writes, non-finite float rejection, and A1 validation.
68. The `serialize_workbook_security_dict` PyO3 export moved into
   `src/native_writer_workbook_metadata.rs` on 2026-04-29 so the backend no
   longer owns workbook-level metadata serialization.
69. Workbook save orchestration moved into `python/wolfxl/_workbook_save.py`
   on 2026-04-29 while preserving the public `Workbook.save` wrapper,
   encrypted-save behavior, modify-mode flush order, same-path in-place saves,
   and write-mode workbook/sheet flush ordering.
70. Workbook formula calculation helpers moved into
   `python/wolfxl/_workbook_calc.py` on 2026-04-29 while preserving
   `calculate`, `recalculate`, evaluator caching, and workbook-level cached
   formula value aggregation.
71. Native writer workbook lifecycle helpers moved into
   `src/native_writer_workbook.rs` on 2026-04-29 while preserving idempotent
   sheet creation, rename/move error translation, consumed-on-save semantics,
   and filesystem write errors.
72. Workbook lifecycle helpers moved into
   `python/wolfxl/_workbook_lifecycle.py` on 2026-04-29 while preserving
   `close`, context-manager entry/exit, tempfile cleanup, and debug
   representation behavior.
73. Writer sheet setup and page-break slot adapters moved out of
   `crates/wolfxl-writer/src/emit/sheet_xml.rs` on 2026-04-29 into
   `emit/sheet_setup.rs` and `emit/page_breaks.rs` while preserving typed
   sheet views, sheet format defaults, sheet protection, page margins, page
   setup, header/footer, and row/column page-break slot ordering.
74. Worksheet rich-text run payload conversion moved into
   `python/wolfxl/_worksheet_rich_text.py` on 2026-04-29 while preserving the
   writer and patcher flush payload shape for `CellRichText` runs.
75. Worksheet debug representation moved into
   `python/wolfxl/_worksheet_repr.py` on 2026-04-29 while preserving the
   openpyxl-shaped `repr(worksheet)` string.
76. Sheet data unit coverage moved from
   `crates/wolfxl-writer/src/emit/sheet_xml.rs` into
   `crates/wolfxl-writer/src/emit/sheet_data.rs` on 2026-04-29, keeping
   `sheet_xml.rs` focused on integration ordering while preserving coverage for
   blank cells, number formatting, SST insertion, booleans, formula result
   variants, date serials, and style attributes.
77. Merge, sheet-view, and column unit coverage moved from
   `crates/wolfxl-writer/src/emit/sheet_xml.rs` into their owning
   `emit/merges.rs`, `emit/sheet_views.rs`, and `emit/columns.rs` modules on
   2026-04-29, keeping `sheet_xml.rs` focused on full worksheet integration and
   rId/order interactions.
78. Hyperlink, dimension, legacy drawing, and table-part unit coverage moved
   from `crates/wolfxl-writer/src/emit/sheet_xml.rs` into the owning
   `emit/hyperlinks.rs`, `emit/dimension.rs`, `emit/drawing_refs.rs`, and
   `emit/table_parts.rs` modules on 2026-04-29 while preserving rId offset and
   dimension edge-case coverage.
79. Conditional-formatting and data-validation unit coverage moved from
   `crates/wolfxl-writer/src/emit/sheet_xml.rs` into
   `emit/conditional_formats.rs` and `emit/data_validations.rs` on 2026-04-29,
   leaving `sheet_xml.rs` with only empty-sheet, kitchen-sink well-formedness,
   and CF/DV/hyperlink ordering checks.
80. Next helper candidate: continue with another narrow Rust save phase only if
   the state boundary is clean, or switch to Python public API docstrings and
   `_worksheet.py` / `_workbook.py` cleanup if the remaining phases look too
   coupled for another safe extraction.

## Verification gates

For every cleanup branch:

- Run focused Python tests for the touched feature.
- Run relevant Rust crate tests.
- Rebuild the extension with `uv run --no-sync maturin develop` when PyO3/Rust
  changes affect Python behavior.
- When the generated ExcelBench fixture pack is available, run
  `uv run --no-sync pytest tests/test_external_oracle_preservation.py -q` to
  catch modify-mode drift against external writer outputs.
- Run a representative ExcelBench local comparison before merging broad
  refactors.
- Preserve the current no-release freeze until the expanded oracle matrix has
  been truth-passed.

For every external oracle branch:

- The tool must be optional: missing Go/Java/.NET/LibreOffice should skip, not
  fail the core suite.
- Fixture outputs must be deterministic enough for CI or clearly marked
  local-only.
- Any newly found WolfXL gaps must be triaged as: fix now, defer with explicit
  docs, or out of scope.

## Source notes

- Excelize: https://github.com/qax-os/excelize and
  https://xuri.me/excelize/en/releases/v2.10.1.html
- Apache POI spreadsheet docs:
  https://poi.apache.org/components/spreadsheet/ and
  https://poi.apache.org/components/spreadsheet/quick-guide.html
- Apache POI limitations:
  https://poi.apache.org/components/spreadsheet/limitations.html
- ClosedXML pivot-table docs:
  https://docs.closedxml.io/en/latest/features/pivot-tables.html
- ClosedXML conditional-formatting examples:
  https://github.com/closedxml/closedxml/wiki/Conditional-Formatting
- SheetJS CE docs:
  https://docs.sheetjs.com/docs/
- libxlsxwriter conditional-formatting docs:
  https://libxlsxwriter.github.io/working_with_conditional_formatting.html
- libxlsxwriter chart docs:
  https://libxlsxwriter.github.io/working_with_charts.html
