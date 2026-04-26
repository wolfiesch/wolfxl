# Changelog

## Unreleased — RFC-035 (in progress)

`Workbook.copy_worksheet(source, name=None)` (modify mode) — being
shipped in parallel by Sprint Ε Pods α/β/γ/δ. Will be folded into
the `1.1.0` section below at tag time.

## wolfxl 1.1.0 (TBD) — Full structural parity

User-facing release notes: `docs/release-notes-1.1.md`.

### Added

- **RFC-034** — `Worksheet.move_range(cell_range, rows=0, cols=0,
  translate=False)` (modify mode). Paste-style relocation of a
  rectangular block of cells. Formulas inside the moved block are
  paste-translated (`respect_dollar=true` per RFC-012 §5.5):
  relative refs shift by `(rows, cols)`; `$`-marked refs do NOT
  shift. With `translate=True`, formulas in cells outside the
  moved block that reference cells inside the source rectangle
  are also re-anchored. MergeCells / hyperlinks / DV-CF sqref
  pieces fully inside the source rectangle shift with the block;
  pieces straddling the boundary or outside are left in place.
  New crate module: `crates/wolfxl-structural/src/range_move.rs`.
  Patcher Phase 2.5j drain in `src/wolfxl/mod.rs`. Tests:
  `tests/test_move_range_modify.py` (15 cases).
- **RFC-001** — Removed leftover `rust_xlsxwriter` workspace dependency. Native writer (RFC W5 replacement) has been the sole xlsx-write path since Phase 2; the dep was unused.
- **RFC-036** — `Workbook.move_sheet(sheet, offset)` (modify mode).
  Reorders sheets in-place; updates `<sheets>` order in workbook.xml
  and any internal references that depend on tab index.
- **RFC-030** — `Worksheet.insert_rows(idx, amount=1)` /
  `delete_rows(idx, amount=1)` (modify mode). Pure-Rust XML rewrite
  via the new `crates/wolfxl-structural` workspace crate. Shifts
  every cell-coord, dimension, merge, hyperlink, table, DV, CF
  anchor, defined-name, and formula reference touched by the row
  band. Delete tombstones (`#REF!`) are emitted per OOXML semantics.
- **RFC-031** — `Worksheet.insert_cols(idx, amount=1)` /
  `delete_cols(idx, amount=1)` (modify mode). Symmetric to RFC-030
  on the column axis; shares the `wolfxl-structural` crate. Adds the
  col-only `<col>` span splitter (`crates/wolfxl-structural/src/cols.rs`)
  so per-column width / style metadata is preserved across inserts
  and deletes. `idx` accepts either a 1-based int or an Excel
  column letter.

### Fixed

- **RFC-031 round 2** — `<tableColumns>` on `tableN.xml` now
  correctly grows / shrinks when `insert_cols` / `delete_cols`
  overlap the table's column band. Previously the `count="N"`
  attribute and `<tableColumn>` element list stayed at the
  pre-shift size, producing an xlsx that Excel and openpyxl
  refused to load. Fixed by `crates/wolfxl-structural/src/shift_workbook.rs`
  (`extract_table_col_band` + `rewrite_table_columns_block`).
  Regression: `tests/test_col_shift_modify.py::test_rfc031_round2_*`
  (4 cases). Also closes the corresponding action items in
  `Plans/followups/rfc-030-031-api-coordination.md`.

### Notes

- Phase-4a (RFC-030, RFC-031, RFC-036) was dispatched as three
  parallel pods. RFC-031 reconciliation required hand-porting the
  `<col>` splitter onto RFC-030's crate layout — sprint retro now
  documents the rule "if two pods touch a shared crate, sequence
  them, don't parallelize."
- Fuzz / property test for `apply_workbook_shift` added under
  `crates/wolfxl-structural/tests/prop_apply_workbook_shift.rs`
  (5000 deterministic iterations on a small fixture; asserts
  no panic + well-formed XML output).

## wolfxl 0.5.0 (2026-04-20) - PyPI cdylib parity release

### Added

- **`Worksheet.schema()`** and **`Worksheet.classify_format(fmt)`**
  Python methods, plus a module-level `wolfxl.classify_format(fmt)`.
  Both delegate to the bridge added in the previous entry so Python
  callers get byte-compatible answers with `wolfxl schema --format
  json` for the structural fields (name, row count, column names,
  null counts, unique counts, cardinality, samples). See
  "Known divergences" in `tests/test_classifier_parity.py` for the
  two fields that don't yet match (numeric `int` vs `float` and the
  openpyxl-styled format-category gap — both close when the sprint-3
  "Option A" engine-collapse work lands).
- **`infer_sheet_schema(rows, name, number_formats=None)`** — the
  bridge now accepts an optional parallel `List[List[Optional[str]]]`
  of per-cell `number_format` strings. Without it, Python's inferred
  `format_category` would silently drift to `"general"` for every
  column because `py_to_cell` couldn't see format metadata. The
  Python `Worksheet.schema()` passes formats from
  `iter_cell_records(include_format=True)` so both surfaces see the
  same format context going into `wolfxl_core::infer_sheet_schema`.
- **`tests/test_classifier_parity.py`** (~240 LOC) — cross-surface
  drift test. Runs `cargo run --quiet --release -p wolfxl-cli --
  schema <fixture> --format json` as a subprocess and compares the
  result to `Worksheet.schema()` on the same workbook. Four cases:
  structural parity (row/column counts + names + null/unique/cardinality),
  sample-list parity (as multisets), direct `classify_format`
  round-trip over every `FormatCategory` variant, and
  `Worksheet.classify_format` → module-level identity.

### Notes

- **Task #22b ships net-new Python surface, not replacements.** The
  sprint-2 plan described #22b as "replace duplicate classifiers in
  `calamine_styled_backend.rs`" — but inspection showed the cdylib
  doesn't actually duplicate any classifier logic; it just returns
  raw `number_format` strings. The "authoritative classifier" work
  was already fully in `wolfxl-core`. So #22b collapsed to exposing
  the bridge methods on the Python surface and adding the parity
  test, both of which were also in the sprint plan's scope.
- **Parity test is the drift detector going forward.** Any future
  change to either the CLI's schema output or Python's
  `worksheet.schema()` that breaks structural parity (null counts,
  unique counts, column names, etc.) fails CI immediately. The
  narrower format-category and int/float parity will tighten when
  Option-A collapses the reader paths.

### Core Bridge Groundwork

#### Added

- **`wolfxl_core_bridge` PyO3 module** (new `src/wolfxl_core_bridge.rs`,
  ~260 LOC). Exposes three `wolfxl-core` classifiers on the `_rust`
  extension module:
  - `classify_format(fmt: str) -> str` — thin wrapper on
    `wolfxl_core::classify_format`, returns the category string
    (`"general"`, `"currency"`, `"date"`, ...) that `wolfxl schema
    --format json` emits in the `format` field.
  - `classify_sheet(rows: List[List[Any]], name: str = "Sheet1") -> str`
    — returns the sheet-class string (`"empty"`, `"readme"`,
    `"summary"`, `"data"`) that `wolfxl map --format json` emits in
    the `class` field.
  - `infer_sheet_schema(rows, name = "Sheet1") -> dict` — returns the
    per-column schema dict in the same shape as `wolfxl schema
    --format json`, minus the outer `"sheets"` wrapper.
- **Native Python input coercion** in the bridge: `None` / `bool` /
  `int` / `float` / `datetime.datetime` / `datetime.date` /
  `datetime.time` / `str` all map to their `CellValue` counterparts.
  Unknown types fall back to `str()` so the bridge never raises on a
  novel type.
- **`wolfxl-core` dep added to the cdylib's `Cargo.toml`** (version
  `0.8`, path `crates/wolfxl-core`). First time the PyO3 surface has
  taken a direct dep on the core crate — prerequisite for the
  classifier-collapse work in the follow-up PR.
- **`Sheet::from_rows` promoted to `pub`** in `wolfxl-core`. The CSV
  backend already used it crate-internally; making it public lets the
  bridge feed externally-sourced Python lists through
  `infer_sheet_schema` / `classify_sheet` without round-tripping
  through a file.

#### Notes

- **Purely additive surface.** This PR does not replace the duplicate
  per-cell classification calls that already live inside
  `calamine_styled_backend.rs` — that wiring is the follow-up PR
  (sprint-2 task #22b). All 617 existing pytest cases still pass
  unchanged; the bridge is extra surface, not a rewrite.
- **Single source of truth for future consumers.** Python callers that
  want a classifier answer can now go through the bridge and get
  byte-identical results to `wolfxl <subcommand> --format json`. The
  cross-surface parity test lands with task #22b.

## wolfxl-core 0.8.0 / wolfxl-cli 0.8.0 (2026-04-20)

### Added

- **Multi-format `Workbook::open`**: `.xls`, `.xlsb`, `.ods`, and `.csv`
  paths now open through the same API as `.xlsx`. Dispatch lives in
  `Workbook::open`, so `wolfxl peek`, `wolfxl map`, `wolfxl agent`, and
  `wolfxl schema` all gain the new format coverage for free. This
  closes the breadth regression relative to `xleak` (the pre-2.0
  predecessor), which handled the same four formats.
- **CSV backend** (`wolfxl_core::csv_reader`, crate-private): reads a
  CSV into a single synthetic `Sheet` named after the filename stem.
  RFC-4180-ish parser handles quoted fields with embedded commas,
  doubled quotes (`""` → `"`), and `\r\n` / `\n` line endings; ragged
  rows are padded to the max column width so downstream
  `dimensions()` / `headers()` consumers see a rectangular shape.
  Cells land as `CellValue::String` — per invariant B4, schema
  inference is the single source of truth for per-column types.
- **Schema inference parses numeric-looking strings**: a CSV column of
  `"100","200",...` now classifies as `Int` instead of `String`.
  `CellValue::String` cells that parse cleanly as `i64` / `f64` are
  counted as the parsed type in `TypeCounts::observe`; strings with
  currency / thousand-separator / percent markers stay as `Other` and
  classify as `String` (the number-format string, when present, still
  drives the separate `format_category`).
- **`SourceFormat` enum** and `Workbook::format()` accessor expose
  which backend the dispatch routed to — `Xlsx`, `Xls`, `Xlsb`, `Ods`,
  or `Csv` — for callers that need to condition on it.
- **CLI multi-format smoke tests**: `tests/cli.rs` now drives `peek`
  against `.csv`, `.xls`, and `.ods` fixtures and drives `schema`
  against `.csv`, asserting the CSV's numeric columns classify as
  `int`. No goldens locked for non-xlsx renders since calamine's
  xls/xlsb/ods readers return empty styles (R1 risk from the sprint plan)
  and the boxed renderer's column widths can drift without the
  styled fast path.
- **Expanded CLI confidence matrix**: committed a tiny `.xlsb` fixture
  sourced from calamine's MIT-licensed test corpus and now smoke-tests
  `peek`, `map`, `schema`, and `agent` across `.csv`, `.xls`, `.xlsb`,
  and `.ods`. CLI help and README text now describe the broad input
  surface instead of implying `.xlsx` only.
- **Number-format-aware CLI previews**: added a formatted workbook
  fixture and assertions that human-facing `peek` box/text/CSV renders
  currency symbols and percentage formats while JSON preserves raw
  machine values. `agent` keeps compact raw numerics intentionally to
  protect token budgets.

### Changed

- **`Workbook::styles()` errors for non-xlsx formats** with a clear
  "`WorkbookStyles` only supports xlsx" message. xls/xlsb/ods carry no
  style information in calamine's public API, and CSV has no
  concept of styles. Callers that want styled rendering should
  branch on `Workbook::format()` before reaching for `styles()`.
- **`WorkbookMap` on CSV** reports a single sheet entry classified
  via the same heuristics as any other sheet; `named_ranges()`
  returns an empty slice on CSV (no workbook-level metadata
  exists).
- **`wolfxl-cli` depends on `wolfxl-core 0.8`** (was 0.7). CLI
  version bumps to 0.8.0 alongside core — shipping the two in
  lockstep keeps the version math honest for users installing via
  `cargo install wolfxl-cli`.

### Notes

- **xls / xlsb / ods are value-only today.** calamine-styles leaves
  `worksheet_style()` empty for these formats, so
  `Cell::number_format` is always `None`; schema inference still
  classifies numeric columns correctly because it reads values, not
  styles. This is the documented R1 mitigation from the sprint-2 plan.
- **CSV parsing intentionally minimal.** UTF-8 only, no custom
  delimiters, no BOM detection. If users hit workbooks that need more,
  the backend can swap to the `csv` crate later — the public API
  (`Workbook::open`, single synthetic sheet, string-valued cells)
  stays the same.

## wolfxl-core 0.7.0 (superseded by 0.8.0 above)

### Added

- **`xl/styles.xml` cellXfs walker** in `wolfxl-core`: new `ooxml`,
  `styles`, and `worksheet_xml` modules plus a `WorkbookStyles` bundle
  that parses cellXfs + numFmts and per-sheet `(row, col) → styleId`
  maps on demand. `Sheet::load` now resolves `number_format` via a
  two-step chain — calamine-styles' fast path first, then the walker
  fallback — so workbook shapes that leave `Style::get_number_format()`
  returning `None` (openpyxl-emitted styles with unpaired cellStyleXfs,
  and similar edge cases) still surface the author-intended currency /
  percentage / date codes. Public re-exports: `WorkbookStyles`,
  `XfEntry`, `BUILTIN_NUM_FMTS`, `builtin_num_fmt`, `resolve_num_fmt`.
- **Integration test**: `tests/styles_walker.rs` covers the combined
  fast-path + fallback end-to-end on a styled fixture, plus a direct
  `parse_cellxfs` + `parse_num_fmts` + `resolve_num_fmt` drive-through
  on synthetic OOXML.

### Notes

- The scope-docs "Not yet" bullet on the styles walker is now
  resolved; the `schema` format-detection note about openpyxl
  workbooks falling back to `general` no longer applies when the
  workbook actually carries `cellXfs` + `numFmts` (even if calamine
  can't see them). Workbooks that emit no styled cells at all still
  fall back to general because there is nothing to resolve.

## wolfxl-cli 0.7.0 / wolfxl-core 0.6.0 (2026-04-19)

### Added

- **`wolfxl schema <file>` subcommand**: per-column type, cardinality,
  null count, format category, and up to three sample values. Defaults
  to JSON for agent consumption; `--format text` produces a tabular
  terminal view. Pass `--sheet NAME` to scope to one sheet, omit to
  schema every sheet in the workbook.
- **`wolfxl-core::schema` module**: `InferredType`, `Cardinality`,
  `ColumnSchema`, `SheetSchema`, and the `infer_sheet_schema` entry
  point — callable from third-party Rust consumers, identical answers
  to the CLI.

### Notes

- **Cardinality buckets** are: `unique` (every non-null cell distinct),
  `categorical` (≤20 distinct AND distinct × 2 ≤ non-null — the
  "lookup-friendly dimension" bucket an agent needs to plan a `WHERE`
  clause), `high-cardinality` (everything else above the cap or with
  many distincts), and `empty`.
- **Type inference** collapses `Int + Float` in the same column to
  `Float` (numeric supertype). Any other multi-type column resolves to
  `Mixed` so an agent doesn't pick a dominant type from a noisy mix.
- **Unique-count tracking is capped at 10 000** distinct rendered
  values per column; columns past the cap report
  `unique_capped: true` and class as `high-cardinality` (the safer
  bucket — caller won't wrongly treat an unverified column as a
  categorical lookup). Picked so a million-row sheet doesn't blow
  memory on the per-column HashSet.
- **Format category is locked from the first non-empty cell** of each
  column. Mixed-format columns are rare in practice; if a user wanted
  per-cell formatting they would be looking at a CSV. Note: openpyxl-
  generated fixtures often emit no `cellXfs` styles, so format
  detection on those workbooks falls back to `general`. Real Excel-
  authored workbooks carry the styles correctly. The full styles.xml
  walker that lifts this limitation is tracked separately.

## wolfxl-cli 0.6.0 (2026-04-19)

### Added

- **`wolfxl agent <file> --max-tokens N` subcommand**: composes a
  token-budgeted workbook briefing for an LLM context window. Emits a
  workbook overview (every sheet with dims/class/first-column header),
  picks the largest `data`-class sheet (or `--sheet` override), then
  greedily fills the remaining budget with header row, head 3 rows, tail
  2 rows, and up to 8 stratified middle samples. Token counts use
  `tiktoken-rs::cl100k_base` to match the GPT-4 family tokenizer (and
  `spreadsheet-peek/benchmarks/measure_tokens.py`); verified at 0-token
  drift against Python `tiktoken`. Falls back to orientation-only output
  if the budget is too tight (and reports the overage in the footer
  rather than silently truncating).
- **Stratified row sampling**: head + tail + uniform-stride middle
  samples instead of head-only. An LLM seeing rows 1, 2, 3 of a 50-row
  P&L can't tell totals from line items; rows 1-3, 25-26, 49-50 plus
  middle samples surface the shape of the data.
- **Token budget tracker**: `Budget::used_with(buf, section)` re-encodes
  the full concatenation rather than summing per-section counts, because
  cl100k_base BPE merges across boundaries (additive checks would
  over-count and reject sections that actually fit). After PR review,
  the budget reserves a worst-case footer cost up-front so the printed
  `--max-tokens N` is honored end-to-end (body + footer).
- Best-effort `NAMED_RANGES` block (capped at 8 entries with overflow
  marker) gated through `try_append`, so a workbook with hundreds of
  named ranges cannot single-handedly drain the agent's budget.

### Notes

- `--agent` deliberately does NOT thousand-group integers (`1234567`,
  not `1,234,567`). Every comma is a token boundary in cl100k_base, so
  ungrouped costs ~2 tokens vs grouped ~5 for the same number. Pretty
  output costs the agent context.
- The orientation core (workbook overview + sheet header + columns) is
  emitted even when it overflows the budget. We'd rather report the
  overage in the footer than hide workbook structure from the agent.

## wolfxl-core 0.5.0 / wolfxl-cli 0.5.0 (2026-04-19)

### Added

- **`wolfxl map <file>` subcommand**: one-page workbook overview for agents
  that need to orient before fetching cell ranges. Emits per-sheet name,
  dimensions, headers, anchored tables, and a coarse classification
  (`empty` / `readme` / `summary` / `data`) plus workbook-level defined
  names. Two output formats: `--format json` (default, machine-parseable)
  and `--format text` (terminal-friendly, sectioned per sheet, header
  preview capped at 8 columns with overflow count).
- **`wolfxl-core::map` module**: `WorkbookMap`, `SheetMap`, `SheetClass`,
  and the `classify_sheet` heuristic, callable from third-party Rust
  consumers via the new `Workbook::map()` method.
- `Workbook::named_ranges()` and `Workbook::table_names_in_sheet(name)`
  pass-throughs to calamine's metadata accessors. Tables are now eagerly
  loaded at `Workbook::open` so these accessors stay infallible (calamine
  panics on `table_names*` without a prior `load_tables`).
- Test-only `Sheet::from_rows_for_test` constructor (gated behind
  `#[cfg(test)]`) so the classifier can exercise `Empty` / `Readme` /
  sparse `Summary` branches that the committed xlsx fixtures don't hit.

### Notes

- The classifier intentionally does not look at merged cells (the upstream
  PyO3 layer does, but `wolfxl-core` doesn't expose merge metadata yet).
  A merged-title-row sheet today classifies as `summary` via the size +
  density rule, which is the right answer for a typical dashboard.
- Pivot detection is still out of scope — calamine doesn't surface pivot
  parts directly, and the agent value of "this sheet is a pivot" is
  marginal next to dimensions + headers.

## 0.4.0 (2026-04-19)

### Added

- **`wolfxl-core` crate** (crates.io): pure-Rust xlsx reader with Excel
  number-format-aware cell rendering. Exposes `Workbook`, `Sheet`, `Cell`,
  `CellValue`, `FormatCategory`, and `format_cell` for third-party Rust
  consumers. No PyO3 coupling.
- **`wolfxl-cli` crate** (crates.io): installs the `wolfxl` binary with a
  `peek` subcommand. `wolfxl peek <file> [-n N] [-s SHEET] [-w WIDTH]
  [-e {box,text,csv,json}]` produces a styled box preview by default and
  text/csv/json exports tuned for piping into agent or shell pipelines.
  Install via `cargo install wolfxl-cli`.

### Changed

- **PyO3 0.24 → 0.28**: required for Python 3.14 support. No public Python
  API changes; all 611 pytest tests pass on 3.12 and 3.14.
- Repository converted to a Cargo workspace with the existing PyO3 cdylib
  at the root and the new `crates/wolfxl-core` + `crates/wolfxl-cli`
  members.

### Fixed

- `wolfxl-core` currency rendering: `format_currency(1.995, 2)` now returns
  `"$2.00"` (was `"$1.100"` due to splitting `trunc()`/`fract()` separately
  before rounding).

## 0.3.2 (2026-04-16)

### Added

- **Bulk styled cell records**: `Worksheet.iter_cell_records()` and `Worksheet.cell_records()` return populated cells as dictionaries with values, formulas, coordinates, and compact formatting metadata.
- **Record-shape controls**: `include_empty`, `include_format`, `include_formula_blanks`, `include_coordinate`, and per-call `data_only` options support ingestion, dataframe, and sparse-workbook workloads.
- **Robust dimensions**: `Worksheet.calculate_dimension()` now merges stale worksheet dimension tags with parsed value/formula storage and preserves offset used ranges such as `C4:C4`.

### Changed

- `max_row` / `max_column` now benefit from the same stale-dimension hardening while preserving their openpyxl-style bottom/right edge semantics.
- `calculate_dimension()` includes buffered `append()` / `write_rows()` data before save, making write-mode dimension reporting more useful for standalone callers.

## 0.3.1 (2026-02-20)

### Added

- **TIME functions**: `NOW()`, `HOUR()`, `MINUTE()`, `SECOND()` with `_serial_to_time` helper for fractional day extraction
- **OFFSET promoted to builtins**: OFFSET now registered in `_BUILTINS` via `_raw_args` protocol, making it visible in `supported_functions` (was previously a hidden evaluator special case)
- **Print area roundtrip**: `ws.print_area = "A1:D10"` now writes through to the xlsx file via the Rust backend (previously stored in Python but never flushed to the writer)

### Changed

- Builtins: 62 -> 67 (OFFSET + NOW + HOUR + MINUTE + SECOND)
- Whitelist: 63 -> 67 (now fully synced with builtins)
- Evaluator function dispatch refactored to use `_raw_args` attribute protocol instead of string-equality special case

## 0.3.0 (2026-02-19)

### Added

- **Formula engine self-sufficiency**: 62 builtin functions covering math, logic, text, lookup, date, financial, and conditional aggregation
- **openpyxl compat expansion**: freeze/split panes, unmerge_cells, print_area property, conditional formatting, data validation, named ranges, tables
- **VLOOKUP/HLOOKUP builtins**: native lookup functions without `formulas` library dependency
- **Conditional aggregation**: AVERAGEIF, AVERAGEIFS, MINIFS, MAXIFS
- **Text functions**: UPPER, LOWER, TRIM, SUBSTITUTE, TEXT, REPT, EXACT, FIND

## 0.1.1 (2026-02-16)

### Fixed

- Build full wheel matrix for macOS and Windows (Python 3.9-3.13)
- Use macos-14 (Apple Silicon) with cross-compilation for x86_64 macOS wheels (macos-13 Intel runners unavailable)
- Fix Windows build failure caused by PyO3 discovering Python 3.14 pre-release

## 0.1.0 (2026-02-15)

Initial release. Extracted from [ExcelBench](https://github.com/SynthGL/ExcelBench).

### Features

- **Read mode**: Full-fidelity xlsx reading via calamine-styles (Font, Fill, Border, Alignment, NumberFormat)
- **Write mode**: Full-fidelity xlsx writing via rust_xlsxwriter
- **Modify mode**: Surgical ZIP patching for fast read-modify-write workflows (10-14x vs openpyxl)
- **openpyxl-compatible API**: `load_workbook()`, `Workbook()`, Cell/Worksheet/Font/PatternFill/Border
- **Bulk operations**: `read_sheet_values()` / `write_sheet_values()` for batch cell I/O
- **Performance**: 3-5x faster than openpyxl for per-cell operations, up to 5x for bulk writes
