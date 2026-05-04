# Known Limitations

This page lists concrete openpyxl-parity gaps in WolfXL so you can evaluate
fit before migrating. Surfaces that openpyxl itself does not expose
(`.xlsb`/`.xls` writes, `.ods`, VBA authoring) are wolfxl-extras tracked in
[`Plans/openpyxl-parity-program.md`](../../Plans/openpyxl-parity-program.md)
under G25-G28; they are not parity gaps and are not listed here. The
machine-checked source of truth is
[`compatibility-matrix.md`](../migration/compatibility-matrix.md).

## Current hard limits

| Feature | Read | Write | Modify | Notes |
|---------|:----:|:-----:|:------:|-------|
| In-place pivot-table field/filter/aggregation edits | Partial | Partial | Partial | WolfXL constructs pivot caches/tables/charts, copies pivot-bearing sheets, and mutates the source range of existing pivots via `ws.pivot_tables[i].source = ...`; field placement, filter, aggregation-function changes, and live aggregate regeneration remain open. |
| External workbook link authoring (`wb._external_links`) | Yes | No | Preserve | WolfXL exposes a read-only `ExternalLink` collection (target, sheet names, cached values) and round-trips `xl/externalLinks/` parts byte-for-byte on modify-save; append/remove/edit authoring is not implemented yet. |
| Print settings depth | Partial | Partial | Partial | Basic `page_setup`, `page_margins`, print titles, and print areas work. The remaining gap is full PageSetup/PrintOptions depth, especially `<printOptions>` emission and a few less-common PageSetup attributes. |
| Dynamic-array spill metadata | Partial | Partial | Partial | Array formulas and data-table formulas round-trip; dynamic-array spill metadata is not fully preserved yet. |
| Calc-chain edge cases | Partial | Partial | Partial | Basic calc-chain rebuild is supported; cross-sheet ordering, deleted-cell pruning, and `calcChainExtLst` edge cases remain open. |
| Standalone table-driven slicers | No | No | Preserve | Pivot-backed slicers are supported. Slicers tied directly to tables, outside a pivot context, are still unimplemented. |

"Preserve" means the feature survives a load-modify-save cycle untouched, even
though WolfXL does not expose a full authoring API for that surface.

## API surface gaps vs openpyxl

These openpyxl APIs are still incomplete or intentionally narrower:

- `ws.conditional_formatting` supports common rules; some complex builder combinations remain lower-priority.
- Existing pivot-table mutation in v1.0 covers only source-range edits (`ws.pivot_tables[i].source = "Sheet!A1:E100"`); field placement, filter, and aggregation mutations are deferred to v2.

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
