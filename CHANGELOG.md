# Changelog

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
