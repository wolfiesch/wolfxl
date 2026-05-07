# Benchmark Results

> **Reference**: WolfXL **v1.7.0** historical baseline.
> **Status as of**: 2026-04-28.

This page is historical context, not the final WolfXL 2.0 release proof.
Use [Public Evidence Status](../trust/public-evidence.md) for the current
claim policy and release-readiness checklist.

The benchmark numbers below summarise the v1.7 posture vs openpyxl 3.1.x.
For the current harness and timestamped snapshots see **ExcelBench**.

- ExcelBench dashboard: [excelbench.vercel.app](https://excelbench.vercel.app)
- ExcelBench repo: [github.com/SynthGL/ExcelBench](https://github.com/SynthGL/ExcelBench)

## Hardware context

Numbers below were collected on:

- Apple MacBook Pro M4 Pro
- 24 GB RAM
- Python 3.13
- openpyxl 3.1.5 (comparison baseline)
- WolfXL 1.7.0

Always attach runtime details (OS, Python, package versions) when
sharing new runs.

## Headline numbers (v1.7 baseline)

Median of 5 runs per cell. Lower is better. Wall-clock seconds.

### Read

| Workload | openpyxl 3.1.5 | WolfXL 1.7 | Speedup |
|---|---:|---:|---:|
| 1k rows × 10 cols, plain values | 0.18 | 0.04 | **4.5×** |
| 10k rows × 10 cols, plain values | 1.12 | 0.18 | **6.2×** |
| 100k rows × 10 cols, plain values | 11.4 | 1.05 | **10.9×** |
| 100k rows, `read_only=True` (streaming) | 9.8 | 0.55 | **17.8×** |
| 10k rows × 20 cols with styled cells | 5.2 | 0.62 | **8.4×** |
| `.xlsb`, 50k rows × 8 cols | n/a (no openpyxl support) | 0.41 | — |

### Write

| Workload | openpyxl 3.1.5 | WolfXL 1.7 | Speedup |
|---|---:|---:|---:|
| 1k rows × 10 cols, fresh workbook | 0.21 | 0.06 | **3.5×** |
| 10k rows × 10 cols, fresh workbook | 1.85 | 0.22 | **8.4×** |
| 100k rows × 10 cols, fresh workbook | 18.1 | 1.78 | **10.2×** |
| 10k rows × 10 cols, with styles | 3.6 | 0.45 | **8.0×** |
| 1 chart × 4 series × 1k points | 0.31 | 0.09 | **3.4×** |

### Modify mode (the v1.0 / v1.1 differentiator)

For modify-mode workloads WolfXL surgically rewrites the changed
parts of the ZIP and copies everything else verbatim — in contrast
to openpyxl which loads the entire workbook into a Python DOM,
mutates, and re-serialises. The speedup on small-edit workloads is
where WolfXL is most differentiated.

| Workload | openpyxl 3.1.5 | WolfXL 1.7 | Speedup |
|---|---:|---:|---:|
| Touch 1 cell, save (1k-row workbook) | 0.42 | 0.03 | **14×** |
| Touch 1 cell, save (100k-row workbook) | 22.4 | 0.18 | **124×** |
| Add 1 hyperlink, save (10k-row workbook) | 4.1 | 0.09 | **45×** |
| `copy_worksheet` (10k-row, with table + DV + CF) | 3.9 | 0.21 | **18×** |
| `insert_rows(idx=2, amount=100)` (10k-row sheet with formulas) | 2.8 | 0.28 | **10×** |

## Methodology

See [methodology.md](methodology.md) for the full reproduction recipe.
Quick summary:

1. Each test case is a Python script that uses `time.perf_counter`
   to bracket the operation under measurement.
2. Numbers are the **median of 5 runs**, with one warm-up run
   discarded.
3. Fixtures live under `fixtures/excel/` in the
   [ExcelBench](https://github.com/SynthGL/ExcelBench) repo and are
   regenerated deterministically.
4. Memory is sampled via `tracemalloc` and reported separately
   (omitted from the headline tables to keep them readable).

## Regenerate the reports

From the [ExcelBench](https://github.com/SynthGL/ExcelBench) repo:

```bash
uv run excelbench benchmark --tests fixtures/excel --output results
uv run excelbench perf --tests fixtures/excel --output results
uv run excelbench report --input results/xlsx/results.json --output results/xlsx
uv run excelbench heatmap
uv run excelbench html
```

The HTML output ships to the dashboard.

## What to publish with each release

1. Fidelity summary table — every fixture round-trips bit-identical
   or with documented ratchet entries.
2. Performance summary table (this page).
3. Hardware / runtime metadata.
4. Link to raw JSON artifacts in ExcelBench.
5. Notable regressions and fixes since the prior release.

## Where WolfXL is not yet faster

- **Tiny-workbook overhead**. For a workbook with < 100 cells the
  Python ↔ Rust FFI cost dominates; openpyxl can be on par or
  marginally faster. Most relevant for unit-test fixtures, not
  production workloads.
- **Pure-Python computation paths** (e.g. iterating in Python over
  every cell). The Rust backend speeds up the I/O and serialisation
  layers; once you're in a Python loop, you pay Python loop costs.
  Use `iter_rows(values_only=True)` (which streams in Rust) over
  per-cell Python iteration when the workload allows.
- **Formula recompute**. WolfXL's formula engine handles 67
  functions; for workbooks that depend on the long tail of
  spreadsheet formulas, defer recompute to Excel-on-open by
  preserving formulas verbatim (`load_workbook(...,
  data_only=False)`).

## Run benchmarks on your own files

See [run-on-your-files.md](run-on-your-files.md) for a copy-pasteable
harness covering read / write / modify mode.

## Track perf in CI

WolfXL's own CI runs the `tests/test_streaming_perf.py` slow-marked
test (100k-row streaming benchmark). Add `WOLFXL_RUN_PERF=1` to your
CI environment to enable it; the test asserts a wall-clock budget so
regressions surface immediately.
