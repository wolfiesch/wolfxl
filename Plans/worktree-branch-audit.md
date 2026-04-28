# Worktree and Sprint Branch Audit

Date: 2026-04-27
Scope: post-PR #23 recovery, before any release or public launch.

## Release freeze

Do not publish, tag, release, or post externally from this state. The current
goal is to preserve and audit the sprint work until the remaining openpyxl
edge capabilities have either been integrated, explicitly deferred, or proven
out of scope.

## Current repository state

- `main` is at `3bf845c` (`feat(writer): native xlsx writer`), the squash
  merge commit for PR #23.
- PR #23 was merged after green CI and addressed review comments.
- The old `feat/native-writer` remote branch was deleted after merge.
- The stale worktree at
  `/Users/wolfgangschoenberger/Projects/wolfxl-worktrees/synthgl-speed-cell-records`
  was removed and `git worktree prune` was run.
- `git worktree list` now reports only the main checkout.
- The only uncommitted item in the checkout is the pre-existing untracked
  `logs/` directory.

## Important graph finding

`main` and tag `v2.0.0` are not ancestors of each other. They are sibling
integration histories with merge base `102a0d8` (`Expose cached formula and
visibility records (#21)`).

That means:

- Unmerged sprint branches cannot be mass-deleted just because current
  parity tests pass.
- Whole-branch merges from the old sprint branches are unsafe because many
  branch tips would delete or regress current `main` files.
- Recovery should be semantic: mine tests, focused fixes, and docs from the
  old branch tips into fresh branches based on current `main`.

## Cleanup already completed

The following local branches were proven merged into `main` by Git and deleted:

- `codex/bottom-alignment-read`
- `codex/wolfxl-multiformat-cli`
- `core-multi-format`
- `core-styles-walker`
- `feat/formulas-integration`
- `pr-cleanup-e280fe9`
- `pyo3-bridge-add`
- `pyo3-bridge-wire`
- `release/wolfxl-0.5.0`

Remote merged branches still exist and should be deleted only in a deliberate
remote cleanup pass:

- `origin/core-multi-format`
- `origin/core-styles-walker`
- `origin/feat/batch-apis-v0.2.0`
- `origin/feat/formulas-integration`
- `origin/pr-cleanup-e280fe9`
- `origin/pyo3-bridge-add`
- `origin/pyo3-bridge-wire`

## Current parity result

Current `main` reports:

- `tests/parity/openpyxl_surface.py`: 66 surface entries, 66 supported, 0
  known-gap entries.
- `uv run pytest tests/parity -q -x`: 445 passed, 4 skipped, 1 warning.

This closes the tracked SynthGL/openpyxl parity contract, but it does not close
every edge capability named in the docs matrix.

## Remaining production-grade edge audit

These started as the known non-launch-blocking-but-worth-auditing areas from
`docs/migration/compatibility-matrix.md` and the current parity docs. Initial
code/test inspection shows that some public docs are stale: current `main`
already has code and focused tests for several items still listed as "Not Yet"
in the docs.

- Slicers (`xl/slicers/`, `xl/slicerCaches/`): implemented in current `main`
  at the model/patcher/copy level. Evidence:
  `python/wolfxl/pivot/_slicer.py`, `Workbook.add_slicer_cache`,
  `Worksheet.add_slicer`, `tests/test_slicer_copy_worksheet.py`,
  `tests/diffwriter/test_advanced_pivots_bytes.py`, and
  `tests/parity/test_advanced_pivots_parity.py`.
- Pivot calculated fields (`<calculatedField>`) and calculated items
  (`<calculatedItem>`): implemented/tested in current `main`. Evidence:
  `python/wolfxl/pivot/_calc.py`, `tests/test_pivot_calculated_fields.py`,
  and `tests/test_pivot_advanced_styling.py`.
- Pivot GroupItems/date-range grouping (`<fieldGroup>`): still needs a focused
  proof pass; current tests mention advanced pivot coverage, but this should be
  isolated before docs are updated.
- Pivot-table styling beyond the named-style picker: partially implemented
  enough for current advanced-pivot tests; still needs docs/API review before
  flipping public docs.
- In-place pivot edits in modify mode: planned v2.2 territory.
- Combination charts / multi-plot charts: post-v2.0.
- Value-axis display units (`<c:displayUnits>`): still needs a focused proof
  pass.
- Per-data-point chart overrides (`<c:dPt>`): model support exists in
  `python/wolfxl/chart/marker.py` and `python/wolfxl/chart/series.py`; needs
  a focused parity/doc pass before public docs flip.
- Removing charts that survive from a source workbook: v1.8 follow-up in docs.
- `.ods`: out of scope.
- `.xlsb` / `.xls` writes: out of scope; xlsx-only transcribe path.
- Style accessors on `.xlsb` / `.xls` reads: not currently available from
  calamine.

Focused audit command run after this scan:

```bash
uv run pytest \
  tests/test_slicer_copy_worksheet.py \
  tests/diffwriter/test_advanced_pivots_bytes.py \
  tests/parity/test_advanced_pivots_parity.py \
  tests/test_pivot_calculated_fields.py \
  tests/test_pivot_advanced_styling.py \
  tests/parity/test_charts_parity.py \
  -q -x
```

Result: 87 passed, 1 skipped, 1 warning.

## Verification checkpoint

After the audit, the Rust tree was normalized with `cargo fmt --all`.

- `cargo fmt --all -- --check`: passed.
- `uv run ruff check .`: passed.
- `uv run maturin develop`: passed; rebuilt and installed editable
  `wolfxl-2.0.0`.
- `uv run pytest`: passed with 2228 passed, 24 skipped, 10 warnings.
- `cargo test --workspace --exclude wolfxl`: passed for the pure Rust crates.
- `cargo test --workspace`: currently fails while linking the `wolfxl` PyO3
  crate tests on macOS because the standalone Rust test binary cannot resolve
  Python C API symbols such as `_PyBaseObject_Type`, `_PyBool_Type`, and
  related symbols. Treat this as a release-gate/tooling issue to resolve before
  calling the full audit complete; the Python extension itself builds and the
  Python-facing suite is green.

## Branch audit queue

There are 66 unmerged local branch heads after cleanup. All show unique patches
relative to current `main`, but most are old stacked histories and are not safe
to merge wholesale.

Highest-priority recovery branches:

- `feat/sprint-pi-pod-alpha`: page breaks and dimension construction coverage.
- `feat/sprint-pi-pod-epsilon`: RFC-060 cleanup / stale stub annotations.
- `feat/sprint-omicron-pod-3-5`: advanced pivot/slicer cross-cutting tests.
- `feat/sprint-omicron-pod-3`: pivot area formatting / chart formatting.
- `feat/sprint-omicron-pod-1a5`: sheet setup / autofilter / security flush
  guard fixes.
- `feat/sprint-omicron-pod-2`: openpyxl-path drop-in import parity.
- `feat/sprint-nu-pod-epsilon`: v2.0 final docs and launch claims; use only
  as evidence, not as publish-ready material.
- `w2a/styles-xml`, `w2b/sheet-xml`, `w2c/sst-wb-xml`, `w3a/comments-vml`,
  `w3b/tables-xml`, `w3c/cf-dv`: native-writer review fix branches; mine
  individual fixes/tests only.
- `pyo3-bump`: Python 3.14 / PyO3 dependency upgrade candidate; handle as a
  separate dependency modernization PR after capability audit.

## Recommended recovery procedure

1. Keep `main` as the base. Do not merge old sprint branches wholesale.
2. For each priority branch, compare `main..branch` by file and by test intent.
3. Cherry-pick or manually port only narrow tests/fixes that still apply.
4. Run the relevant focused test module after each port.
5. Run full parity and writer suites before any PR.
6. Only after all capability audit items are either green or explicitly
   deferred should release/public-launch work resume.
