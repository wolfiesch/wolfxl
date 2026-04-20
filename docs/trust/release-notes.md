# WolfXL Changelog (Docs View)

This page is a release-summary companion to repository release notes.

## v0.5.0 - 2026-04-20

- Added Python `Worksheet.schema()` and `classify_format()` surfaces backed by
  `wolfxl-core`, plus cross-surface parity tests against the `wolfxl schema`
  CLI output.
- Added the PyO3 `wolfxl_core_bridge` groundwork so Python callers can share
  classifier and schema inference logic with the Rust CLI/core path.
- Added `wolfxl-core 0.8.0` / `wolfxl-cli 0.8.0` multi-format dispatch for
  `.xlsx`, `.xls`, `.xlsb`, `.ods`, and `.csv`.
- Fixed explicit bottom-alignment reads so Excel-authored `bottom` alignment
  round-trips instead of collapsing to the default.

## Release template

Use this template for each release:

```text
## vX.Y.Z - YYYY-MM-DD

- Added:
- Changed:
- Fixed:
- Performance:
- Compatibility:
- Known regressions:
```
