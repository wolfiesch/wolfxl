# Benchmark Results

This page is intentionally lightweight and should link to generated artifacts.

## Current source of truth

- ExcelBench dashboard: [excelbench.vercel.app](https://excelbench.vercel.app)
- ExcelBench repo: [github.com/SynthGL/ExcelBench](https://github.com/SynthGL/ExcelBench)

## Hardware context (current docs baseline)

- Apple MacBook Pro M4 Pro
- 24 GB RAM

Always attach runtime details (OS, Python, package versions) when sharing new runs.

## Regenerate reports

From the [ExcelBench](https://github.com/SynthGL/ExcelBench) repository:

```bash
uv run excelbench report --input results/xlsx/results.json --output results/xlsx
uv run excelbench heatmap
uv run excelbench html
```

## What to publish with each release

1. Fidelity summary table
2. Performance summary table
3. Hardware/runtime metadata
4. Link to raw JSON artifacts
5. Notable regressions and fixes
