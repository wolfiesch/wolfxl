# Operational Troubleshooting

## Import errors

### `ModuleNotFoundError: wolfxl._rust`

- Ensure package install succeeded.
- Verify Python environment is the one you execute from.
- Run: `python3 -c "import wolfxl._rust as m; print(m.build_info())"`

### Legacy import failures (`excelbench_rust`)

- Install/update shim package if legacy paths are still used.
- Prefer migrating runtime imports to `wolfxl`/`wolfxl._rust`.

## Save failures

- Confirm write or modify mode is used before `save()`.
- Check filesystem permissions and target path validity.

## Data mismatch reports

- Reproduce with smallest workbook possible.
- Attach expected vs actual output and code snippet.
- Reference [Known Differences](../fidelity/known-differences.md).
