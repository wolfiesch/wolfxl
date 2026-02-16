# WolfXL Docs

## Speed Up Excel Pipelines Without Rewriting Your Workflow

WolfXL is a Rust-backed Excel engine for Python teams using openpyxl-style workflows.
It is built around three priorities:

1. Lower migration friction
2. Reproducible performance results
3. Explicit fidelity tracking

## Start Here

| Goal | Link |
|---|---|
| Get running quickly | [Quickstart](getting-started/quickstart.md) |
| Migrate existing code | [Openpyxl Migration Guide](migration/openpyxl-migration.md) |
| Check support coverage | [Compatibility Matrix](migration/compatibility-matrix.md) |
| Understand benchmarks | [Benchmark Methodology](performance/methodology.md) |
| Review caveats first | [Known Limitations](trust/limitations.md) |

## Why WolfXL

Most Excel tooling optimizes one slice of the problem. WolfXL targets practical production workflows where read/write/modify speed and correctness both matter.

### Migration-first

Keep openpyxl-style workbook, worksheet, and cell patterns.

### Evidence-first

Benchmark claims are tied to reproducible commands and raw artifacts.

### Correctness-first

Performance is paired with fidelity checks, not assumed.

## Benchmark Environment

Unless otherwise noted, published benchmark runs were measured on:

- Apple MacBook Pro M4 Pro
- 24 GB RAM

See [Benchmark Methodology](performance/methodology.md) and [Benchmark Results](performance/benchmark-results.md).

## Documentation Map

- `getting-started/` - install and first-run guides
- `migration/` - compatibility and migration from openpyxl/legacy imports
- `performance/` - reproducible benchmark process and troubleshooting
- `fidelity/` - scoring model and fixture strategy
- `api/` - workbook/worksheet/cell/style API reference
- `operations/` - production rollout and upgrade checklists
- `trust/` - limitations, changelog, integrity notes

## Build Docs Locally

From repository root:

```bash
uv run --with mkdocs-material mkdocs serve -f docs/mkdocs.yml
```
