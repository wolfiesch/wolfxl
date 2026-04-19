# Changelog

## 0.4.0 (2026-04-19)

### Added

- **`wolfxl-core` crate** (crates.io): pure-Rust xlsx reader with Excel
  number-format-aware cell rendering. Exposes `Workbook`, `Sheet`, `Cell`,
  `CellValue`, `FormatCategory`, and `format_cell` for third-party Rust
  consumers. No PyO3 coupling.
- **`wolfxl-cli` crate** (crates.io): installs the `wolfxl` binary with a
  `peek` subcommand. `wolfxl peek <file> [-n N] [-s SHEET] [-w WIDTH]
  [-e {box,text,csv,json}]` produces previews at xleak byte-parity for
  text/csv/json (77/81 in-repo fixtures match) and a styled box view by
  default. Install via `cargo install wolfxl-cli`.

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
