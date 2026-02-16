# Known Limitations

WolfXL targets high-impact openpyxl-style workflows, not complete openpyxl API parity. This page lists concrete gaps so you can evaluate fit before migrating.

## Not yet supported

| Feature | Read | Write | Modify | Notes |
|---------|:----:|:-----:|:------:|-------|
| Images (pictures/drawings) | No | No | — | Preserved in modify mode but not accessible via API |
| Diagonal borders | Partial | Yes | — | Read extracts top/bottom/left/right; diagonal is recognized but fidelity score is 1 |
| Charts | No | No | — | Preserved untouched in modify mode |
| Pivot tables | No | No | — | Preserved in modify mode; creation requires Excel |
| VBA macros | No | No | — | Preserved in modify mode (`.xlsm` files) |
| DefinedName manipulation | No | No | No | Named ranges can be read; programmatic creation not yet exposed |
| Print settings | No | No | — | Page setup, margins, headers/footers |
| Protection (sheet/workbook) | No | No | — | Password-protected files cannot be opened |
| Rich text (mixed fonts in one cell) | No | No | No | Font applies per-cell, not per-run |

"Preserved in modify mode" means the feature survives a load-modify-save cycle untouched, even though WolfXL cannot read or create it programmatically.

## API surface gaps vs openpyxl

These openpyxl APIs do not have WolfXL equivalents yet:

- `ws.add_image()`, `ws.add_chart()`
- `ws.protection`, `wb.security`
- `ws.page_setup`, `ws.print_options`
- `ws.auto_filter` (preserved in modify, not creatable)
- `ws.conditional_formatting` complex rule builders (basic rules work)
- `copy_worksheet()`

## Performance claim guardrails

- Speedup numbers (3-5x) are measured on Apple Silicon with ExcelBench. Your results will vary with workload shape, file complexity, and hardware.
- Modify mode (10-14x) is measured on files where only a small fraction of cells change. The advantage shrinks as the edit ratio approaches 100%.
- Always validate on your own files before committing to a migration.

## Integrity guidance

- Use reproducible fixtures and benchmarks for acceptance testing.
- Review output workbooks in Excel for business-critical templates.
- WolfXL's fidelity is tracked by [ExcelBench](https://excelbench.vercel.app) — check the dashboard for current scores.
