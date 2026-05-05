# Known Limitations

This page lists concrete fidelity limits in WolfXL so you can evaluate fit
before migrating. Surfaces that openpyxl itself does not expose
(`.xlsb`/`.xls` writes, `.ods`, VBA authoring, standalone table-driven
slicer authoring) are wolfxl-extras tracked in
[`Plans/openpyxl-parity-program.md`](../../Plans/openpyxl-parity-program.md);
they are not openpyxl parity gaps. The machine-checked source of truth is
[`compatibility-matrix.md`](../migration/compatibility-matrix.md).

## Current hard limits

| Feature | Read | Write | Modify | Notes |
|---------|:----:|:-----:|:------:|-------|
| Pivot cache record regeneration after layout edits | Yes | No | Refresh | WolfXL can mutate existing pivot source ranges, row/column/page field placement, page-field selection, and data-field aggregation. Layout edits stamp `refreshOnLoad="1"` and let Excel regenerate cache records on open rather than recalculating `pivotCacheRecords` inside WolfXL. |

"Refresh" means the edited workbook opens cleanly and Excel refreshes derived
cache data when needed.

## API surface gaps vs openpyxl

These openpyxl APIs are intentionally narrower:

- Existing pivot-table layout edits prioritize OOXML correctness and refresh-on-open over WolfXL-side cache-record regeneration.

## Performance claim guardrails

- Speedup numbers (3-5x) are measured on Apple Silicon with ExcelBench. Your results will vary with workload shape, file complexity, and hardware.
- Modify mode (10-14x) is measured on files where only a small fraction of cells change. The advantage shrinks as the edit ratio approaches 100%.
- Always validate on your own files before committing to a migration.

## Integrity guidance

- Use reproducible fixtures and benchmarks for acceptance testing.
- Before public releases, run the local external-oracle fixture pack generated
  by ExcelBench. It currently exercises Excelize, ClosedXML, NPOI, ExcelJS, and
  Apache POI outputs through LibreOffice and WolfXL preservation checks.
- Review output workbooks in Excel for business-critical templates.
- WolfXL's fidelity is tracked by [ExcelBench](https://excelbench.vercel.app) — check the dashboard for current scores.
- When sharing numbers externally, check [Public Evidence Status](public-evidence.md) first so historical snapshots are not presented as current release evidence.
