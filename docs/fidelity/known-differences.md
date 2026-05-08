# Known Differences

This page tracks meaningful differences versus openpyxl/Excel behavior.

## Current guidance

- WolfXL focuses on high-value openpyxl-style workflows, not complete API parity.
- Feature support should be validated via the compatibility matrix and fixture tests.

## Modify mode notes

- Modify mode uses a surgical patcher for targeted changes.
- Preserve-value behavior for style-only edits is regression-tested.

## App-level evidence limits

- PowerView-bearing workbooks are treated as app-unsupported evidence in the
  local Microsoft Excel smoke harness. Excel can block those files behind a
  read-only unsupported-content prompt, so they are useful for package-fidelity
  radar coverage but not for proving clean editable Excel open/close behavior.
  A Microsoft Error Reporting `SIG_FORCE_QUIT` / `merp` log after this kind of
  blocked prompt is not, by itself, treated as an OOXML repair/corruption
  signal; the repair signal remains an explicit Excel repair/error dialog.
- The interactive Excel evidence gate is intentionally stricter than package
  preservation. On 2026-05-08, targeted source-workbook probes passed for:
  pivot state, slicer state, timeline state, external-link state, macro project
  presence, and embedded-control state. The combined strict audit over all six
  source-workbook probe reports now reports `ready=true` with no incomplete
  reports:
  `/tmp/wolfxl-interactive-evidence-external-oracle-all-20260508.json`.

## Contributing a difference report

When reporting a difference, include:

1. Minimal workbook sample
2. Expected behavior (Excel/openpyxl)
3. Actual WolfXL behavior
4. Repro code snippet

Link issue reports to benchmark/fidelity artifacts when possible.
