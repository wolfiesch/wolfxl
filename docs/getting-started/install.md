# Install

## PyPI

```bash
pip install wolfxl
```

## Verify Installation

```bash
python3 -c "import wolfxl; from wolfxl import _rust; print('ok', wolfxl.__all__, _rust.build_info())"
```

## Local Development (from this repo)

```bash
uv sync
uv run python -c "import wolfxl; print('wolfxl import ok')"
```

If you are working on the Rust extension itself, build the local extension in editable mode in the Rust subproject environment.

## Troubleshooting

- See [Performance Troubleshooting](../performance/perf-troubleshooting.md) for runtime issues.
- See [Operational Troubleshooting](../operations/troubleshooting.md) for packaging/import issues.
