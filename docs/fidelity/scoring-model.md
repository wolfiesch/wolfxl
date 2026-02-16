# Fidelity Scoring Model

WolfXL uses [ExcelBench](https://excelbench.vercel.app) scoring conventions for feature fidelity.

## Score bands

- `3` = all tests pass (complete fidelity)
- `2` = at least 80% pass
- `1` = at least 50% pass
- `0` = below 50% pass

## Weighting

- `basic` tests are must-pass functionality
- `edge` tests capture bonus and long-tail behavior

## Why this model

Fidelity is not binary in real-world spreadsheets. The score bands make progress and regressions visible while still rewarding full correctness.
