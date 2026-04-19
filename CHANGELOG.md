# Changelog

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

## 0.5.0 (2026-04-19)

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
