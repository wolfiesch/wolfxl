# Known Limitations

WolfXL targets high-impact openpyxl-style workflows, not complete openpyxl API parity. This page lists concrete gaps so you can evaluate fit before migrating.

## Current hard limits

| Feature | Read | Write | Modify | Notes |
|---------|:----:|:-----:|:------:|-------|
| `.ods` workbooks | No | No | No | OpenDocument is out of scope. |
| `.xlsb` / `.xls` writes | No | No | No | Binary and legacy formats are read-only; transcribe to `.xlsx` with a new workbook. |
| Styles on `.xlsb` / `.xls` reads | Partial | No | No | Native `.xlsb` reads expose cell styles; legacy `.xls` remains value-only. |
| VBA macros | Preserve | No | Preserve | `.xlsm` parts survive modify-mode saves, but WolfXL does not inspect or generate VBA. |
| In-place pivot-table edits | Partial | Partial | Partial | WolfXL can construct pivot caches/tables/charts and copy pivot-bearing sheets; editing arbitrary existing pivot definitions remains limited. |
| Image replacement/deletion | Partial | Yes | Partial | `Image(...)` and `ws.add_image(...)` are supported; replacing or deleting existing image media is not a public API yet. |
| Combination / multi-plot charts | Partial | No | Partial | Single-family chart construction is covered; combination charts are deferred. |
| Diagonal border fidelity | Partial | Yes | Partial | Read recognizes diagonal metadata, but full diagonal style parity is still lower-confidence than top/bottom/left/right borders. |

"Preserve" means the feature survives a load-modify-save cycle untouched, even
though WolfXL does not expose a full authoring API for that surface.

## API surface gaps vs openpyxl

These openpyxl APIs are still incomplete or intentionally narrower:

- `ws.add_image()` and `ws.add_chart()` support construction, but not public replace/delete operations.
- `ws.conditional_formatting` supports common rules; some complex builder combinations remain lower-priority.
- `.xlsb` workbooks expose read-side style metadata; `.xls` workbooks remain value-only and style accessors raise by design.
- Existing pivot-table mutation is narrower than construction and copy support.

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
