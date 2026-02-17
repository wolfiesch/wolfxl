# Benchmark Methodology

WolfXL performance claims should be reproducible.

## Default benchmark hardware context

Unless a result explicitly says otherwise, benchmark numbers in WolfXL docs were collected on:

- Apple MacBook Pro M4 Pro
- 24 GB RAM

## Principles

1. Publish exact commands
2. Publish hardware/runtime context
3. Include raw outputs
4. Pair speed with fidelity checks
5. Avoid cherry-picked scenarios

## Standard context to include

- CPU model and RAM
- OS and Python version
- WolfXL and comparison library versions
- Dataset/fixture description
- Number of runs and aggregation method (median recommended)

## Suggested benchmark commands

From the [ExcelBench](https://github.com/SynthGL/ExcelBench) repository:

```bash
uv run excelbench benchmark --tests fixtures/excel --output results
uv run excelbench perf --tests fixtures/excel --output results
uv run excelbench report --input results/xlsx/results.json --output results/xlsx
```

## Reporting guidelines

- Show both absolute times and throughput
- Show where WolfXL is slower if observed
- Keep benchmark scripts and fixture definitions versioned
