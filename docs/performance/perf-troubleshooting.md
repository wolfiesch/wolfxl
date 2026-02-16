# Performance Troubleshooting

## Symptoms and checks

### Slower-than-expected read

- Confirm you are importing WolfXL (`from wolfxl import load_workbook`).
- Ensure local build is release-optimized when benchmarking native code.
- Re-run benchmark after warm-up to avoid one-time startup effects.

### Slower-than-expected modify/save

- Confirm `modify=True` is used when editing existing files.
- Check whether many style updates are being applied per cell.
- Compare with same workbook and same changed-cell count.

### Inconsistent results

- Pin Python and dependency versions.
- Run on idle machine and avoid thermal throttling.
- Report medians, not single-run numbers.

## Integrity first

Never optimize by skipping correctness checks. Pair every performance run with spot validation in Excel and automated tests.
