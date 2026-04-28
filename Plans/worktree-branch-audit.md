# Worktree and Sprint Branch Audit

Date: 2026-04-27
Scope: post-PR #23 recovery, before any release or public launch.

## Release freeze

Do not publish, tag, release, or post externally from this state. The current
goal is to preserve and audit the sprint work until the remaining openpyxl
edge capabilities have either been integrated, explicitly deferred, or proven
out of scope.

## Current repository state

- `main` started this audit at `3bf845c` (`feat(writer): native xlsx writer`),
  the squash merge commit for PR #23. The local checkout is now ahead with
  audit commits that have not been pushed.
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
  `tests/test_pivot_slicers.py`, `tests/test_advanced_pivots_save_smoke.py`,
  `tests/diffwriter/test_advanced_pivots_bytes.py`, and
  `tests/parity/test_advanced_pivots_parity.py`.
- Pivot calculated fields (`<calculatedField>`) and calculated items
  (`<calculatedItem>`): implemented/tested in current `main`. Evidence:
  `python/wolfxl/pivot/_calc.py`, `tests/test_pivot_calculated_fields.py`,
  `tests/test_pivot_calculated_items.py`, and
  `tests/test_pivot_advanced_styling.py`.
- Pivot GroupItems/date-range grouping (`<fieldGroup>`): implemented/tested.
  Evidence: `tests/test_pivot_group_items.py`.
- Pivot-table styling beyond the named-style picker: partially implemented.
  PivotArea formats and pivot-scoped conditional formats are covered by
  `tests/test_pivot_advanced_styling.py`; broader theme/banded-format polish
  remains a partial capability.
- In-place pivot edits in modify mode: planned v2.2 territory.
- Combination charts / multi-plot charts: post-v2.0.
- Value-axis display units (`<c:dispUnits>`): implemented in this audit slice
  across Python dict serialization, PyO3 parsing, and Rust XML emit.
- Per-data-point chart overrides (`<c:dPt>`): implemented in this audit slice
  across Python dict serialization, PyO3 parsing, and Rust XML emit.
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

Additional focused proof commands:

```bash
uv run pytest \
  tests/test_page_breaks.py \
  tests/test_dimension_helpers.py \
  tests/parity/test_page_breaks_parity.py \
  tests/diffwriter/test_page_breaks_bytes.py \
  tests/test_pivot_group_items.py \
  tests/test_pivot_advanced_styling.py \
  -q
```

Result: 147 passed.

```bash
uv run pytest tests/test_charts_write.py tests/parity/test_charts_parity.py -q
cargo test -p wolfxl-writer --test charts
cargo test -p wolfxl-writer --lib emit::charts
```

Result: 47 passed / 2 skipped for Python chart tests, 47 passed for the
writer chart integration tests, and 13 passed for chart emitter unit tests.

Post-PR #23 branch-mining proof:

```bash
uv run pytest tests/test_compat_shims.py tests/parity/test_openpyxl_path_compat.py -q
```

Result: 271 passed, 2 skipped.

```bash
uv run pytest \
  tests/test_pivot_calculated_fields.py \
  tests/test_pivot_calculated_items.py \
  tests/test_pivot_group_items.py \
  tests/test_pivot_advanced_styling.py \
  tests/test_pivot_slicers.py \
  tests/test_slicer_copy_worksheet.py \
  tests/test_advanced_pivots_save_smoke.py \
  tests/diffwriter/test_advanced_pivots_bytes.py \
  tests/parity/test_advanced_pivots_parity.py \
  -q
cargo test -p wolfxl-pivot
```

Result: 193 passed / 1 warning for the Python advanced-pivot slice, and 55
passed for `wolfxl-pivot`.

```bash
uv run pytest \
  tests/test_page_breaks.py \
  tests/test_dimension_helpers.py \
  tests/parity/test_page_breaks_parity.py \
  tests/diffwriter/test_page_breaks_bytes.py \
  -q
cargo test -p wolfxl-writer
```

Result: 77 passed for the Pi-alpha page-break/dimension slice, and 355 unit
tests / 47 chart integration tests / 2 roundtrip tests / 1 doctest passed for
`wolfxl-writer`.

Release-claim branch audit:

- `feat/sprint-nu-pod-epsilon` was audited as release-claim evidence only.
  Do not merge it wholesale: its tip sits on an old stacked sprint history, and
  the branch-wide diff would regress current `main`.
- The useful v2.0 docs surfaces already exist on current `main`
  (`README.md`, `CHANGELOG.md`, `docs/release-notes-2.0.md`,
  `docs/migration/compatibility-matrix.md`,
  `docs/migration/openpyxl-migration.md`, `Plans/launch-posts.md`, and
  `tests/parity/KNOWN_GAPS.md`), but they are not publish-ready.
- Current release/public-launch copy still contains `<!-- TBD: BENCHMARK
  NUMBERS -->`, `<!-- TBD: SHA -->`, and checklist placeholders. Those must be
  filled from fresh benchmark and release evidence before any tag, PyPI
  publish, Twitter/X post, or HN post.
- Public claims need a final truth pass before launch. In particular, verify
  or soften "full openpyxl replacement" and "only Python OOXML library that
  constructs pivotCacheRecords" language; decide whether partial pivot styling,
  OLAP/external pivot caches, and in-place pivot edits are clearly deferred.
- `CHANGELOG.md` still has stale v2.1/future-work language around slicers,
  calculated fields/items, and GroupItems even though the post-PR #23 audit
  slice has code and focused tests for those capabilities. Fix this before
  treating the changelog as release collateral.
- Release readiness still requires GitHub Actions plus clean install/import
  smoke from the published artifact/wheels. The local full-suite proof is
  necessary evidence, not sufficient release evidence.

Dependency-modernization branch audit:

- `pyo3-bump` was audited and is superseded by current `main`. No branch port
  is needed.
- Current `main` already uses `pyo3 = "0.28"` and
  `pyo3-build-config = "0.28"` in `Cargo.toml`, with maturin enabling the
  `extension-module` feature through `pyproject.toml`.
- The current implementation is better than the old branch tip because
  `extension-module` is no longer enabled for ordinary Cargo workspace tests;
  this is why the full `cargo test --workspace` proof now passes on macOS.
- Remaining PyO3 work should be a fresh modernization PR, not an old-branch
  merge: replace deprecated `Bound::downcast` / `downcast_into` usage with the
  PyO3 0.28 `cast` / `cast_into` API, remove warning-only unused imports, then
  rerun `uv run maturin develop`, `cargo test --workspace`, and a clean
  Python 3.14 import smoke.

## Verification checkpoint

After the audit, the Rust tree was normalized with `cargo fmt --all`.

- `cargo fmt --all -- --check`: passed.
- `uv run ruff check .`: passed.
- `uv run maturin develop`: passed; rebuilt and installed editable
  `wolfxl-2.0.0`.
- `uv run pytest`: passed with 2248 passed, 24 skipped, 10 warnings on the
  2026-04-28 post-branch-mining proof.
- `cargo test --workspace --exclude wolfxl`: passed for the pure Rust crates
  (two non-snake-case test-name warnings remain in `wolfxl-rels`).
- `cargo test --workspace`: passed on the 2026-04-28 post-branch-mining
  proof after splitting PyO3's
  `extension-module` feature out of the default Cargo test build and enabling
  it only through maturin. This fixed the macOS Python C API linker failure.
- `cargo test -p wolfxl --lib`: 225 passed after refreshing stale assertions
  and fixing VML `o:idmap` local-name parsing.
- `uv run pytest tests/test_sheet_setup_copy_worksheet.py
  tests/diffwriter/test_sheet_setup_bytes.py
  tests/parity/test_print_settings_parity.py
  tests/parity/test_sheet_protection_parity.py
  tests/parity/test_workbook_security_parity.py
  tests/test_autofilter_filters.py
  tests/parity/test_openpyxl_path_compat.py tests/test_compat_shims.py -q`:
  341 passed, 2 skipped.

## Branch audit queue

There are 66 unmerged local branch heads after cleanup. All show unique patches
relative to current `main`, but most are old stacked histories and are not safe
to merge wholesale.

Highest-priority recovery branches:

- `feat/sprint-pi-pod-alpha`: audited against current `main`; page-break and
  dimension tests/code are already present and included in the 77-test focused
  proof above. The only newer branch-only deltas in this slice are unused
  imports, so nothing was ported.
- `feat/sprint-pi-pod-epsilon`: audited/mined against current `main`; the
  useful RFC-060 cleanup was ported by removing stale `(stub)` annotations and
  adding representative former-stub constructors to `tests/test_compat_shims.py`.
- `feat/sprint-omicron-pod-3-5`: audited against current `main`; its advanced
  pivot/slicer cross-cutting tests are already present and included in the
  193-test focused proof above.
- `feat/sprint-omicron-pod-3`: audited against current `main`; calculated
  fields/items, field grouping, PivotArea formatting, pivot CFs, and
  chart-format coverage are already present and included in the focused proof
  above.
- `feat/sprint-omicron-pod-1a5`: audited against current `main`; its
  sheet-setup / page-break / autofilter / workbook-security no-op guard
  queues are present in `XlsxPatcher::do_save`, and the focused Python proof
  above is green.
- `feat/sprint-omicron-pod-2`: audited against current `main`; the
  openpyxl-path drop-in import parity test is present and included in the
  341-test focused proof.
- `feat/sprint-nu-pod-epsilon`: v2.0 final docs and launch claims; use only
  as evidence, not as publish-ready material. Audited on 2026-04-28; current
  docs still need benchmark/SHA replacement and a final truth pass before
  publication.
- `w2a/styles-xml`: audited against current `main`; RGB attribute escaping for
  styles plus regression tests are present in `crates/wolfxl-writer/src/emit/styles_xml.rs`.
- `w2b/sheet-xml`: audited against current `main`; unstyled blank cells are
  excluded from `<dimension ref>` and styled blanks still count.
- `w2c/sst-wb-xml`: audited against current `main`; `_xlnm.Print_Area`
  user-defined-name dedupe is present in `workbook_xml.rs`.
- `w3a/comments-vml`: audited against current `main`; the VML shape-id formula
  and anchor tuple comments are present.
- `w3b/tables-xml`: audited against current `main`; multi-table `tableParts`
  relationship-id coverage is present.
- `w3c/cf-dv`: audited against current `main`; empty `<conditionalFormatting>`
  wrappers are skipped when every rule is a stub variant.
- `pyo3-bump`: audited on 2026-04-28 and superseded by current `main`; the
  PyO3 0.28 dependency upgrade and `extension-module` feature split are
  already present. Remaining work is warning cleanup / Python 3.14 smoke on a
  fresh branch.

## Recommended recovery procedure

1. Keep `main` as the base. Do not merge old sprint branches wholesale.
2. For each priority branch, compare `main..branch` by file and by test intent.
3. Cherry-pick or manually port only narrow tests/fixes that still apply.
4. Run the relevant focused test module after each port.
5. Run full parity and writer suites before any PR.
6. Only after all capability audit items are either green or explicitly
   deferred should release/public-launch work resume.
