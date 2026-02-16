# Security and Integrity

## File integrity principles

1. Minimize file-surface changes during modify workflows
2. Validate critical workbook outputs after save
3. Keep regression tests for high-risk edit paths

## Reporting an integrity issue

Include:

- input workbook (or minimal repro)
- exact code used
- expected output vs actual output
- WolfXL version and environment

## Safe rollout suggestions

- Use side-by-side comparison in early rollout stages.
- Keep fallback path to existing engine until confidence is established.
- Track regressions explicitly in changelog.
