# Known Limitations

WolfXL targets high-impact openpyxl-style workflows, not complete openpyxl API parity. This page lists concrete gaps so you can evaluate fit before migrating.

## Current hard limits

| Feature | Read | Write | Modify | Notes |
|---------|:----:|:-----:|:------:|-------|
| `.ods` workbooks | No | No | No | OpenDocument is out of scope. |
| `.xlsb` / `.xls` writes | No | No | No | Binary and legacy formats are read-only; transcribe to `.xlsx` with a new workbook. |
| Styles on `.xls` reads | No | No | No | Legacy `.xls` is value-only; style accessors raise by design. Native `.xlsb` reads do expose cell styles. |
| VBA macros | Inspect (modify-mode) | No | Preserve | `.xlsm` parts round-trip on modify-save; `Workbook.vba_archive` exposes raw `xl/vbaProject.bin` bytes for read-only inspection. WolfXL does not generate or edit VBA. |
| In-place pivot-table field edits | n/a | n/a | No | Pivot construction (cache + table + chart) is supported, copy_worksheet of pivot-bearing sheets deep-clones, and source-range mutation works (`ws.pivot_tables[i].source = "Sheet!A1:E100"`). Field placement, filter, and aggregation mutations on existing pivots are deferred. |
| External workbook links (`wb._external_links`) | Yes (read-only) | No | Preserve | Read-only `ExternalLink` collection (target, sheet names, cached values); modify-mode round-trips `xl/externalLinks/` parts byte-for-byte. Authoring new external links is deferred. |
| Standalone slicers (no pivot) | No | No | No | Slicers wired to a `PivotCache` are supported; table-driven and other non-pivot slicers are deferred. |
| Source-workbook chart deletion/replacement | Yes | n/a | Partial | `remove_chart` / `replace_chart` cover charts added in the current session. Deleting or replacing charts that were already present in the source workbook is deferred. |
| Image append into existing drawing parts | Yes | Yes | Partial | Image writes and common modify-mode adds are supported. Appending a new image to a sheet whose source workbook already has a drawing relationship is still narrower than Excel/openpyxl template-editing expectations. |

"Preserve" means the feature survives a load-modify-save cycle untouched, even
though WolfXL does not expose a full authoring API for that surface.

## API surface gaps vs openpyxl

These openpyxl APIs are still incomplete or intentionally narrower:

- `ws.conditional_formatting` covers cellIs / containsText / expression / colorScale / iconSet / dataBar with the common cfvo ladders; rare builder combinations may need a manual probe before relying on them.
- `.xls` workbooks remain value-only and style accessors raise by design; `.xlsb` reads do expose cell styles.
- Existing pivot-table mutation covers source-range edits (`ws.pivot_tables[i].source = "Sheet!A1:E100"`); field placement, filter, and aggregation mutations on already-authored pivots are deferred.
- `print_options` / `page_setup` / `page_margins` cover the common attributes; the full ~30-attribute openpyxl surface is partially covered (depth audit pending).
- `DefinedName` covers the common authoring path; openpyxl's edge-case attrs (`hidden`, `comment`, `custom_menu`, `function`, `function_group_id`, `shortcut_key`) are partially covered.
- `Workbook(write_only=True)` streams rows to a per-sheet temp file with bounded peak RSS (SST + styles only). It is append-only — `ws.append(row)` is the single row API; random access (`ws["A1"]`, `ws.cell(...)`, `iter_rows`), `merge_cells`, `add_chart`, `add_image`, `add_table`, `add_data_validation`, `add_pivot_table`, and `conditional_formatting` raise `AttributeError` on a write-only worksheet. Re-saving raises `WorkbookAlreadySaved`. Matches openpyxl's `_write_only.py` contract.
- WolfXL enforces bounded OOXML ZIP package reads and atomic save paths. Files that exceed configured entry/count/ratio limits fail closed rather than being partially parsed or partially written.

For the canonical, machine-checked support matrix, see [`docs/migration/compatibility-matrix.md`](../migration/compatibility-matrix.md), generated from `docs/migration/_compat_spec.py` and validated by `tests/test_openpyxl_compat_oracle.py`.

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
