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
| In-place pivot-table edits | Partial | Partial | Partial | WolfXL constructs pivot caches/tables/charts, copies pivot-bearing sheets, and (v1.0, RFC-070 Option B) mutates the source range of existing pivots via `ws.pivot_tables[i].source = ...`; field placement, filter, and aggregation mutations are deferred. |
| Image replacement/deletion | Partial | Yes | Partial | `Image(...)` and `ws.add_image(...)` are supported; replacing or deleting existing image media is not a public API yet. |
| Combination / multi-plot charts | Partial | No | Partial | Single-family chart construction is covered; combination charts are deferred. |
| External workbook links (`wb._external_links`) | Partial | No | Preserve | v1.0 exposes a read-only `ExternalLink` collection (target, sheet names, cached values) and round-trips `xl/externalLinks/` parts byte-for-byte on modify-save; authoring new external links is deferred. |
| Diagonal border fidelity | Partial | Yes | Partial | Read recognizes diagonal metadata, but full diagonal style parity is still lower-confidence than top/bottom/left/right borders. |

"Preserve" means the feature survives a load-modify-save cycle untouched, even
though WolfXL does not expose a full authoring API for that surface.

## API surface gaps vs openpyxl

These openpyxl APIs are still incomplete or intentionally narrower:

- `ws.add_image()` and `ws.add_chart()` support construction, but not public replace/delete operations.
- `ws.conditional_formatting` supports common rules; some complex builder combinations remain lower-priority.
- Existing pivot-table mutation in v1.0 covers only source-range edits (`ws.pivot_tables[i].source = "Sheet!A1:E100"`); field placement, filter, and aggregation mutations are deferred to v2.
- `Workbook(write_only=True)` streams rows to a per-sheet temp file with bounded peak RSS (SST + styles only). It is append-only — `ws.append(row)` is the single row API; random access (`ws["A1"]`, `ws.cell(...)`, `iter_rows`), `merge_cells`, `add_chart`, `add_image`, `add_table`, `add_data_validation`, `add_pivot_table`, and `conditional_formatting` raise `AttributeError` on a write-only worksheet. Re-saving raises `WorkbookAlreadySaved`. Matches openpyxl's `_write_only.py` contract.

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
