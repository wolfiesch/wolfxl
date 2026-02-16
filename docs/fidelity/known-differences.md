# Known Differences

This page tracks meaningful differences versus openpyxl/Excel behavior.

## Current guidance

- WolfXL focuses on high-value openpyxl-style workflows, not complete API parity.
- Feature support should be validated via the compatibility matrix and fixture tests.

## Modify mode notes

- Modify mode uses a surgical patcher for targeted changes.
- Preserve-value behavior for style-only edits is regression-tested.

## Contributing a difference report

When reporting a difference, include:

1. Minimal workbook sample
2. Expected behavior (Excel/openpyxl)
3. Actual WolfXL behavior
4. Repro code snippet

Link issue reports to benchmark/fidelity artifacts when possible.
