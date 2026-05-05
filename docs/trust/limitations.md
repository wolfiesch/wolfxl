# Openpyxl Parity Limitations

This page lists the current openpyxl-parity truth in WolfXL so you can
evaluate fit before migrating. There are no known `partial` or `not_yet`
rows for the tracked openpyxl-supported API surface in the generated
compatibility matrix. Surfaces that openpyxl itself does not expose
(`.xlsb`/`.xls` writes, `.ods`, VBA authoring, standalone table-driven
slicer authoring) are wolfxl-extras tracked in
[`Plans/openpyxl-parity-program.md`](../../Plans/openpyxl-parity-program.md);
they are not openpyxl parity gaps. The machine-checked source of truth is
[`compatibility-matrix.md`](../migration/compatibility-matrix.md).

## Current caveat

| Feature | Read | Write | Modify | Notes |
|---------|:----:|:-----:|:------:|-------|
| Pivot cache record regeneration after layout edits | Yes | Excel refresh | Refresh | WolfXL can mutate existing pivot source ranges, row/column/page field placement, page-field selection, and data-field aggregation. Layout edits stamp `refreshOnLoad="1"` and let Excel regenerate derived cache records on open rather than recalculating `pivotCacheRecords` inside WolfXL. This is not an openpyxl advantage: openpyxl preserves existing pivot parts but does not provide a public cache-record regeneration engine either. |

"Refresh" means the edited workbook opens cleanly and Excel refreshes derived
cache data when needed.

## API surface gaps vs openpyxl

None known in the tracked matrix/oracle. Any future claim of an
openpyxl-supported gap should be added to
[`_compat_spec.py`](../migration/_compat_spec.py) as `partial` or `not_yet`
with a failing oracle probe before being documented here.

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
