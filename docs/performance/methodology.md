# Benchmark Methodology

> **Reference**: WolfXL **v1.7.0** (Sprint Ξ).
> **Status as of**: 2026-04-27.

WolfXL performance claims should be reproducible. This page is the
contract: every number we publish (in release notes, the docs, or
external blog posts) traces back to a recipe on this page.

## Default benchmark hardware context

Unless a result explicitly says otherwise, benchmark numbers in WolfXL
docs were collected on:

- Apple MacBook Pro M4 Pro
- 24 GB RAM
- Python 3.13
- WolfXL **1.7.0**
- openpyxl **3.1.5** (comparison baseline)

## Principles

1. **Publish exact commands.** Every number has a runnable script.
2. **Publish hardware + runtime context.** CPU model, RAM, OS,
   Python version, package versions of every library compared.
3. **Include raw outputs.** The HTML dashboard exports the raw JSON
   alongside.
4. **Pair speed with fidelity checks.** A workbook that's faster but
   wrong is a regression, not a win. Every benchmark fixture
   round-trips through the parity ratchet
   (`tests/parity/openpyxl_surface.py`).
5. **Avoid cherry-picked scenarios.** The headline tables average
   over a representative spread (1k / 10k / 100k rows × plain /
   styled / chart-bearing fixtures).

## Standard context to include

When sharing a new run, attach:

- CPU model and RAM.
- OS and Python version.
- WolfXL and comparison library versions.
- Dataset / fixture description (row × col counts, style density,
  chart count, formula density).
- Number of runs and aggregation method (median recommended).

## Suggested benchmark commands

From the [ExcelBench](https://github.com/SynthGL/ExcelBench) repo:

```bash
uv run excelbench benchmark --tests fixtures/excel --output results
uv run excelbench perf --tests fixtures/excel --output results
uv run excelbench report --input results/xlsx/results.json --output results/xlsx
```

For one-off ad-hoc benchmarks on your own files see
[run-on-your-files.md](run-on-your-files.md).

## Reporting guidelines

- Show both **absolute times** and **throughput** (rows/sec for
  reads/writes, cells/sec for full-cell-iteration paths).
- **Show where WolfXL is slower** if observed. The README and docs
  call out the tiny-workbook FFI overhead explicitly; new
  regressions surface as the same kind of honest disclosure.
- Keep benchmark scripts and fixture definitions versioned in git.
- Tag benchmark runs with the wolfxl + openpyxl versions; never
  compare across versions without a re-run.

## Construction-side benchmarks (NEW in v1.7)

v1.6 added chart construction; v1.7 adds remove/replace + adds the
chart-construction harness to the perf suite. The methodology above
applies but with two extra pieces of context:

1. **Chart cache rebuild semantics.** WolfXL emits chart XML with
   the cell-range references but skips the `<c:strCache>` /
   `<c:numCache>` cached-values block. Excel rebuilds these on
   first open. openpyxl emits the cache. Skipping the cache is
   ~30 % faster; the user sees no difference unless they're
   programmatically reading the cached values from a closed
   workbook.
2. **Image media reuse on copy_worksheet.** WolfXL aliases image
   media (RFC-035 §5.3): a copied sheet's drawing rels point at
   the same `xl/media/imageN.png` as the source. openpyxl
   re-encodes via Pillow on the copy. Aliasing avoids 50× workbook
   bloat on logo-heavy templates and is faster, but means a future
   "modify a copy's image" idiom would deep-clone (tracked as a
   v1.8+ follow-up).

When publishing chart or copy_worksheet benchmarks, mention these
contracts so the reader understands what's being compared.
