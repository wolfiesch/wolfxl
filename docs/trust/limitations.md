# Known Limitations

This page lists concrete openpyxl-parity gaps in WolfXL so you can evaluate
fit before migrating. Surfaces that openpyxl itself does not expose
(`.xlsb`/`.xls` writes, `.ods`, VBA authoring, standalone table-driven
slicer authoring) are wolfxl-extras tracked in
[`Plans/openpyxl-parity-program.md`](../../Plans/openpyxl-parity-program.md);
they are not parity gaps and are not listed here. The
machine-checked source of truth is
[`compatibility-matrix.md`](../migration/compatibility-matrix.md).

## Current hard limits

| Feature | Read | Write | Modify | Notes |
|---------|:----:|:-----:|:------:|-------|
| In-place pivot-table field/filter/aggregation edits | Partial | Partial | Partial | WolfXL constructs pivot caches/tables/charts, copies pivot-bearing sheets, and mutates the source range of existing pivots via `ws.pivot_tables[i].source = ...`; field placement, filter, aggregation-function changes, and live aggregate regeneration remain open. |
| External workbook link authoring (`wb._external_links`) | Yes | No | Preserve | WolfXL exposes a read-only `ExternalLink` collection (target, sheet names, cached values) and round-trips `xl/externalLinks/` parts byte-for-byte on modify-save; append/remove/edit authoring is not implemented yet. |

"Preserve" means the feature survives a load-modify-save cycle untouched, even
though WolfXL does not expose a full authoring API for that surface.

## API surface gaps vs openpyxl

These openpyxl APIs are still incomplete or intentionally narrower:

- Existing pivot-table mutation in v1.0 covers only source-range edits (`ws.pivot_tables[i].source = "Sheet!A1:E100"`); field placement, filter, and aggregation mutations are deferred to v2.
- External workbook links are exposed for inspection and opaque modify-mode preservation; appending, removing, and editing external links are deferred to v2.

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
