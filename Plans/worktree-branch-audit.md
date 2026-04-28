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

Chart parity truth pass after installing `lxml` locally:

- `tests/parity/_chart_helpers.py` now canonicalizes chart XML with namespace
  prefix rewriting so default-vs-prefixed namespace differences do not mask
  real structural drift.
- The chart writer now matches openpyxl more closely for plot-area/title
  elision, series-title formulas, default scatter style omission, series
  reindexing, default marker/tick behavior, and chart-level data labels.
- `uv run --no-sync pytest tests/parity/test_charts_parity.py -q`: 28 passed,
  8 skipped.
- `uv run --no-sync pytest tests/test_charts_write.py
  tests/test_charts_modify.py tests/test_pivot_charts.py
  tests/test_copy_worksheet_modify.py tests/test_copy_worksheet_libreoffice.py
  -q`: 101 passed, 3 skipped, 1 warning.
- `cargo test -p wolfxl-writer --test charts`: 47 passed.

LibreOffice/copy-worksheet truth pass:

- Two namespace/composition bugs were fixed in the patcher:
  1. pivot-cache and sheet-copy workbook mutations now add `xmlns:r` on the
     workbook root when later workbook-level `r:id` attributes need it, even
     if the source workbook only declared `xmlns:r` on a child `<sheet>`.
  2. defined-name merge now composes against Phase 2.7's in-progress
     `xl/workbook.xml` patch instead of re-reading the source ZIP and
     dropping the cloned `<sheet>` entry.
- `WOLFXL_RUN_LIBREOFFICE_SMOKE=1 uv run --no-sync pytest
  tests/diffwriter/soffice_smoke.py tests/test_copy_worksheet_libreoffice.py
  tests/test_array_formula_libreoffice.py
  tests/test_pivot_charts.py::test_pivot_chart_libreoffice_renders -q`:
  47 passed, 28 warnings.

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
- `CHANGELOG.md` stale v2.1/future-work language around slicers, calculated
  fields/items, and GroupItems has been cleaned up in the audit pass. It now
  treats those as shipped/supported and keeps only broader pivot styling,
  OLAP/external caches, and in-place pivot edits as deferred.
- Release readiness still requires GitHub Actions plus clean install/import
  smoke from the published artifact/wheels. The local full-suite proof is
  necessary evidence, not sufficient release evidence.

Local release-artifact smoke, 2026-04-28:

- `uv run --no-sync maturin build --release --out dist` passed and built the
  macOS arm64 CPython 3.14 wheel
  `wolfxl-2.0.0-cp314-cp314-macosx_11_0_arm64.whl`.
- A fresh temporary venv installed that wheel with `openpyxl` and `Pillow`.
  Import/version smoke passed: `wolfxl.__version__ == "2.0.0"` and
  `wolfxl._rust.build_info()` reported enabled backends
  `calamine-styles`, `wolfxl`, and `native`.
- The fresh-venv artifact smoke created a write-mode chart workbook via
  `Workbook()`/`add_chart`, reopened it with openpyxl, and confirmed chart
  parts were present.
- The same fresh-venv smoke created a modify-mode pivot workbook via
  `load_workbook(..., modify=True)`, `add_pivot_cache`, and
  `add_pivot_table`, reopened it with openpyxl, and confirmed both
  `xl/pivotTables/pivotTable*` and `xl/pivotCache/pivotCacheRecords*` parts
  were present.
- This is useful local artifact evidence only. It does not replace CI matrix
  wheel builds, cross-platform fresh installs, benchmark refresh, or manual
  Excel/LibreOffice visual checks before publish.

Advanced pivot/slicer cross-renderer smoke, 2026-04-28:

- `WOLFXL_RUN_LIBREOFFICE_SMOKE=1 uv run --no-sync pytest
  tests/test_advanced_pivots_save_smoke.py tests/test_pivot_slicers.py
  tests/test_slicer_copy_worksheet.py
  tests/diffwriter/test_advanced_pivots_bytes.py
  tests/parity/test_advanced_pivots_parity.py
  tests/test_pivot_advanced_styling.py -q`: 96 passed, 1 warning
  (openpyxl warns that the unknown slicer extension is unsupported and will
  be removed on its own rewrite).
- A fresh temporary workbook containing a pivot cache/table, calculated
  field, calculated item, grouped revenue field, pivot-area format, slicer
  cache, and slicer presentation was generated with the public modify-mode
  APIs. `/opt/homebrew/bin/soffice --headless --convert-to xlsx` and
  `--convert-to pdf` both exited 0, produced valid outputs, and emitted no
  repair/error stderr. The source ZIP contained the expected
  `xl/pivotCache/*`, `xl/pivotTables/*`, `xl/slicerCaches/*`, and
  `xl/slicers/*` parts.
- This closes an automated LibreOffice smoke slice for advanced pivots and
  slicers. It still does not replace manual Excel GUI inspection of the
  rendered layout and slicer interactivity before public release.

ExcelBench local benchmark/fidelity smoke, 2026-04-28:

- The sibling ExcelBench venv was still importing `wolfxl==0.1.0`, so it was
  updated locally to install this checkout as `wolfxl==2.0.0` before any
  benchmark smoke.
- ExcelBench initially did not register the `wolfxl` adapter because it still
  expected the old `RustXlsxWriterBook` PyO3 class. WolfXL 2.0 exposes
  `NativeWorkbook`; the sibling ExcelBench adapter was patched with a
  backwards-compatible writer-class fallback and committed there as
  `e0d5431` (`fix: support wolfxl native writer adapter`).
- `uv run --no-sync excelbench perf --tests fixtures/excel --output
  /tmp/wolfxl-2.0-excelbench-full-smoke --adapter openpyxl --adapter wolfxl
  --warmup 1 --iters 5 --memory-mode getrusage` completed across the 19
  committed fixture features with no adapter notes. This is a smoke-scale
  performance run only, not the publishable benchmark replacement. In that
  local run WolfXL was faster than openpyxl on every reported p50 read/write
  fixture cell; publishable numbers still require the standard iteration count,
  hardware/runtime capture, raw artifact retention, and dashboard regeneration.
- The first fidelity run exposed dashboard-facing gaps in image metadata,
  hyperlink display text, no-style/totals-row tables, double underline, and
  dataBar/colorScale conditional formatting. Those were split into a sibling
  ExcelBench adapter patch (native writer registration, OOXML image-anchor
  reads, hyperlink display-cell writes) and a WolfXL native-writer patch
  (double-underline styles, dataBar/colorScale emission, no-style table
  preservation, and `totalsRowCount` output).
- `uv run --no-sync excelbench benchmark --tests fixtures/excel --output
  /tmp/wolfxl-2.0-excelbench-full-fidelity-fix --adapter openpyxl --adapter
  wolfxl` completed after those fixes. Both openpyxl and WolfXL now report
  125/125 read tests and 125/125 write tests passed, 18/18 green scored
  features, and no diagnostics in this local fixture smoke. The public
  dashboard still needs an intentional artifact refresh, but the earlier
  fidelity triage list is closed locally.

Dependency-modernization branch audit:

- `pyo3-bump` was audited and is superseded by current `main`. No branch port
  is needed.
- Current `main` already uses `pyo3 = "0.28"` and
  `pyo3-build-config = "0.28"` in `Cargo.toml`, with maturin enabling the
  `extension-module` feature through `pyproject.toml`.
- The current implementation is better than the old branch tip because
  `extension-module` is no longer enabled for ordinary Cargo workspace tests;
  this is why the full `cargo test --workspace` proof now passes on macOS.
- PyO3 warning cleanup was completed from current `main`, not by merging the
  old branch: deprecated `Bound::downcast` / `downcast_into` usage was replaced
  with the PyO3 0.28 `cast` / `cast_into` API, warning-only import/dead-code
  noise was narrowed, and the remaining `wolfxl-rels` test-name warnings were
  removed.

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
- PyO3 modernization proof after the cleanup:
  `cargo check -p wolfxl`, `cargo test -p wolfxl --lib`,
  `cargo test --workspace`, `cargo fmt --all -- --check`,
  `git diff --check`, `uv run maturin develop`, `uv run ruff check .`,
  `uv run pytest`, and a modify-mode pivot construction smoke all passed.
- Post-chart/LibreOffice truth pass proof:
  `git diff --check`, `cargo fmt --all -- --check`,
  `uv run --no-sync ruff check .`, `cargo test --workspace`, and
  `uv run --no-sync pytest -q` all passed. The full Python suite reported
  2278 passed, 29 skipped, 10 warnings.

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
- `feat/sprint-omicron-pod-1a`: audited against current `main`; its branch
  tip was deferral documentation for unfinished sheet-setup scope. Current
  `main` already has the closed follow-up doc at
  `Plans/followups/omicron-pod-1a-deferrals.md`, with Pod 1A.5 marking the
  native-writer emit, patcher drain, PyO3 parser/bindings, copy-worksheet
  deep-clone, and tests as landed.
- `feat/sprint-omicron-pod-1a5`: audited against current `main`; its
  sheet-setup / page-break / autofilter / workbook-security no-op guard
  queues are present in `XlsxPatcher::do_save`, and the focused Python proof
  above is green.
- `feat/sprint-omicron-pod-2`: audited against current `main`; the
  openpyxl-path drop-in import parity test is present and included in the
  341-test focused proof.
- `feat/sprint-omicron-pod-1b`: audited against current `main`; AutoFilter
  model/emit/evaluate support is present via `crates/wolfxl-autofilter`,
  `python/wolfxl/worksheet/filters.py`, and the worksheet flush path.
  Focused proof: `tests/test_autofilter_filters.py` plus the openpyxl-path
  ratchet passed in the 405-test constructor/import proof; `cargo test -p
  wolfxl-autofilter` passed 49 tests.
- `feat/sprint-omicron-pod-1c`: audited against current `main`; RFC-057
  `ArrayFormula` / `DataTableFormula` plumbing is present in the cell and
  worksheet layers with diffwriter, parity, and LibreOffice coverage.
  Focused proof: `tests/test_array_formula.py`,
  `tests/test_data_table_formula.py`, and related import parity passed in the
  405-test proof.
- `feat/sprint-omicron-pod-1d`: audited against current `main`; workbook
  protection / file sharing classes plus native-writer and patcher security
  emission are present. Focused proof: `tests/test_workbook_protection.py`,
  `tests/parity/test_workbook_security_parity.py` in earlier broad runs, and
  `cargo test -p wolfxl-writer --lib parse::` passed 40 parse tests including
  workbook-security coverage.
- `feat/sprint-omicron-pod-1e`: audited against current `main`; public
  exceptions, `IndexedList`, and cell-class/path shims are present and covered
  by `tests/test_indexed_list.py`, `tests/parity/test_openpyxl_path_compat.py`,
  and `tests/test_compat_shims.py`.
- `feat/sprint-nu-pod-epsilon`: v2.0 final docs and launch claims; use only
  as evidence, not as publish-ready material. Audited on 2026-04-28; current
  docs still need benchmark/SHA replacement and a final truth pass before
  publication.
- `feat/sprint-pi-pod-beta`, `feat/sprint-pi-pod-gamma`, and
  `feat/sprint-pi-pod-delta`: inspected after the post-chart truth pass.
  Their visible unique branch tips are the shared Sprint Π pre-dispatch
  scaffold on top of old stacked history; no pod-specific beta/gamma/delta
  implementation commits were found to port wholesale. Current `main` already
  carries the Sprint Π constructor/import ratchet for merge/table/copier,
  style constructors, workbook internals, and re-export paths. Focused proof:
  `uv run --no-sync pytest tests/test_autofilter_filters.py
  tests/test_array_formula.py tests/test_data_table_formula.py
  tests/test_workbook_protection.py tests/test_indexed_list.py
  tests/parity/test_sprint_pi_constructors.py
  tests/parity/test_openpyxl_path_compat.py tests/test_compat_shims.py -q`
  reported 405 passed, 2 skipped.
- Older T1/T1.5 RFC branches `feat/rfc-012-formula-xlator`,
  `feat/rfc-021-defined-names`, `feat/rfc-023-comments`,
  `feat/rfc-024-tables`, `feat/rfc-030-insert-delete-rows`,
  `feat/rfc-031-insert-delete-cols`, `feat/rfc-034-move-range`,
  `feat/rfc-035-allocator-and-planner`,
  `feat/rfc-035-patcher-and-coordinator`,
  `feat/rfc-035-tests-and-parity`, and `feat/rfc-036-move-sheet`:
  audited as superseded by current `main`. Their core surfaces are present in
  the patcher and Rust helper crates, and current focused proof is green:
  `uv run --no-sync pytest tests/test_defined_names_modify.py
  tests/test_comments_modify.py tests/test_tables_modify.py
  tests/test_axis_shift_modify.py tests/test_col_shift_modify.py
  tests/test_move_range_modify.py tests/test_copy_worksheet_modify.py
  tests/test_copy_worksheet_byte_stable.py tests/test_copy_worksheet_write_mode.py
  tests/test_move_sheet_modify.py tests/parity/test_copy_worksheet_parity.py -q`
  reported 127 passed. `cargo test -p wolfxl-structural`,
  `cargo test -p wolfxl-formula`, and `cargo test -p wolfxl-rels` also passed.
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
- `feat/release-1.1-notes-and-fuzz`, `feat/rfc-001-w5-ripout`,
  `feat/rfc-035-copy-worksheet-spec`, and
  `feat/rfc-035-bugfixes-and-status`: audited as historical release/RFC
  collateral. Current `main` already has the RFC-035 spec, v1.1 release notes,
  rust-xlsxwriter ripout, and later copy-worksheet fixes; no branch-wide port
  is needed.
- `feat/sprint-theta-pod-a`, `feat/sprint-theta-pod-b`,
  `feat/sprint-theta-pod-c`, and `feat/sprint-theta-pod-d`: audited as
  superseded by current `main`. The useful self-closing `<sheets/>`
  permissive loader, CDATA-safe splice path, image deep-clone follow-up, and
  v1.2 release notes are already present. Current copy-worksheet proof is
  covered by the 127-test RFC-035 focused run above plus the full-suite proof.
- `feat/sprint-iota-pod-alpha`, `feat/sprint-iota-pod-beta`,
  `feat/sprint-iota-pod-gamma`, and `feat/sprint-iota-pod-delta`: audited as
  superseded. Current `main` carries rich-text reads, streaming reads,
  password reads, and v1.3 release collateral. Focused proof remains in the
  broad parity/full-suite runs; password read/write surfaces were rechecked
  in the 145-test trust/surface proof below.
- `feat/sprint-kappa-pod-alpha`, `feat/sprint-kappa-pod-beta`,
  `feat/sprint-kappa-pod-gamma`, and `feat/sprint-kappa-pod-delta`: audited
  as present on current `main`. `.xlsb` / `.xls` value reads, bytes /
  `BytesIO` / file-like dispatch, format classification, fixtures, and
  parity tests are present. Focused proof:
  `uv run --no-sync pytest tests/test_format_dispatch.py
  tests/parity/test_xlsb_reads.py tests/parity/test_xls_reads.py -q`
  passed as part of the 145-test trust/surface proof below.
- `feat/sprint-lambda-pod-alpha`, `feat/sprint-lambda-pod-beta`,
  `feat/sprint-lambda-pod-gamma`, and `feat/sprint-lambda-pod-delta`:
  audited as present on current `main`. Write-side Agile encryption,
  image construction / add-image, streaming datetime hardening, and v1.5
  release notes are present. A stale trust/docstring truth-pass was ported
  from repo reality, not by merging the old branch tips.
- `feat/sprint-mu-pod-alpha`, `feat/sprint-mu-pod-beta`,
  `feat/sprint-mu-pod-gamma`, `feat/sprint-mu-pod-delta`,
  `feat/sprint-mu-pod-epsilon`, `feat/sprint-mu-prime-pod-alpha`,
  `feat/sprint-mu-prime-pod-beta`, `feat/sprint-mu-prime-pod-gamma`, and
  `feat/sprint-mu-prime-pod-delta`: audited as present on current `main`.
  Chart construction, modify-mode chart adds, chart fixtures, deep-clone
  tests, 3D-family tests, and v1.6/v1.6.1 notes are already in-tree. Focused
  proof included `tests/test_charts_write.py`, `tests/test_charts_3d.py`, and
  `tests/test_copy_worksheet_chart_deep_clone.py` in the 206-test proof.
- `feat/sprint-nu-pod-gamma` and `feat/sprint-nu-pod-delta`: audited as
  present on current `main`. Pivot modify-mode integration, pivot-bearing
  copy tests, and pivot-chart linkage tests are in-tree. Focused proof
  included `tests/test_pivot_modify.py`,
  `tests/test_copy_worksheet_pivot_deep_clone.py`, and
  `tests/test_pivot_charts.py` in the 206-test proof.
- `codex/synthgl-cell-records-contract`: audited as superseded. Its branch
  tip only cached style IDs before record-format population; current `main`
  already has the same `record_style_id` /
  `populate_record_format_for_style_id` shape in
  `src/calamine_styled_backend.rs`.
- `backup-main-pre-rebase`: audited as historical only. The branch tip is an
  old pre-workspace snapshot (`feat(read): bulk styled cell records +
  dimension hardening`) and its branch-wide diff would delete the modern
  writer/pivot/structural crates, release docs, fixtures, and parity tests.
  Current `main` already carries the useful `cell_records()` and dimension
  hardening surfaces.

Trust/docs truth-pass after this branch audit:

```bash
uv run --no-sync pytest \
  tests/test_format_dispatch.py tests/parity/test_xlsb_reads.py \
  tests/parity/test_xls_reads.py tests/test_encrypted_writes.py \
  tests/parity/test_encrypted_write_parity.py tests/test_images_write.py \
  tests/test_images_modify.py tests/parity/test_images_parity.py \
  tests/test_copy_worksheet_deep_clone_images.py tests/test_charts_write.py \
  tests/test_charts_3d.py tests/test_copy_worksheet_chart_deep_clone.py \
  tests/test_pivot_modify.py tests/test_copy_worksheet_pivot_deep_clone.py \
  tests/test_pivot_charts.py -q
```

Result: 206 passed, 4 skipped, 1 warning.

```bash
uv run --no-sync pytest \
  tests/parity/test_surface_smoke.py tests/test_format_dispatch.py \
  tests/parity/test_xlsb_reads.py tests/parity/test_xls_reads.py \
  tests/test_encrypted_writes.py tests/parity/test_encrypted_write_parity.py \
  tests/test_images_write.py tests/test_images_modify.py \
  tests/parity/test_images_parity.py -q
```

Result: 145 passed, 3 skipped. `uv run --no-sync ruff check
python/wolfxl/__init__.py tests/parity/openpyxl_surface.py` passed.

## Recommended recovery procedure

1. Keep `main` as the base. Do not merge old sprint branches wholesale.
2. For each priority branch, compare `main..branch` by file and by test intent.
3. Cherry-pick or manually port only narrow tests/fixes that still apply.
4. Run the relevant focused test module after each port.
5. Run full parity and writer suites before any PR.
6. Only after all capability audit items are either green or explicitly
   deferred should release/public-launch work resume.
