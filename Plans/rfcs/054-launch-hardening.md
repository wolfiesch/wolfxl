# RFC-054: v2.0.0 launch hardening + docs + README rewrite

Status: Approved (pre-dispatch; Pod-ε can scaffold in parallel)
Owner: Sprint Ν Pod-ε
Phase: 5 (2.0)
Estimate: M
Depends-on: RFC-047, RFC-048, RFC-049 (all three SHIPPED before this RFC closes — but doc scaffolding may proceed in parallel)
Unblocks: v2.0.0 PyPI publish + public-launch posts

## 1. Background — Problem Statement

After Sprint Ξ (v1.7.0) shipped, only one construction-side gap
remains: pivot tables. Sprint Ν closes it via RFC-047/048/049.
RFC-054 is the launch-day envelope: docs / README / migration
guide / launch posts / KNOWN_GAPS close-out / version bump /
PyPI publish / public posts.

This RFC has no Rust or Python code to ship. It's all docs and
release artifacts. The "doc-pod scaffolds in parallel with the
code pods using TBD markers" pattern (lesson #3) applies.

## 2. Surface Area

### 2.1 New / rewritten docs

* `docs/release-notes-2.0.md` (new) — full v2.0 release notes,
  including pivot-construction examples and the "deferred to
  v2.1" out-of-scope section.
* `docs/migration/openpyxl-migration.md` — extend the v1.7
  rewrite with a new "Pivot tables" section.
* `docs/migration/compatibility-matrix.md` — flip "Pivot table
  construction" from ❌ to ✅; update ecosystem comparison.
* `tests/parity/KNOWN_GAPS.md` — close out the pivot row;
  "Closed in 2.0" section; "Out of scope" reduces to slicers
  + calculated fields + GroupItems + OLAP + pivot-styling.
* `CHANGELOG.md` — prepend v2.0.0 entry (RFC-047/048/049/054
  shipped + version bumps).

### 2.2 README rewrite

* Drop the "for the 95th-percentile case" qualifier.
* Headline: **"Full openpyxl replacement, drop-in compatible,
  10×-100× faster."**
* Update the feature matrix table: pivot tables ✅, pivot charts
  ✅.
* Refresh the benchmark numbers table to v2.0 (recomputed by
  Sprint Ξ Pod-γ; v2.0 should be equivalent unless pivot-emit
  benchmarks shift the headline).
* Add a "Pivot tables in 6 lines" snippet near the top.

### 2.3 Launch posts

`Plans/launch-posts.md` was scaffolded in Sprint Ξ
(RFC-053). Pod-ε finalizes:
* HN "Show HN" post text.
* 8-tweet Twitter/X thread.
* r/Python post.
* dev.to long-form post.
* GitHub Discussions launch announcement.
* Pre-launch checklist (PyPI verified install, doc site live,
  benchmark dashboard live).
* Post-launch checklist (monitor issues, respond on HN/r/Python,
  bug-fix point release plan if needed).

### 2.4 Version bump

* `pyproject.toml` `version = "2.0.0"`.
* `Cargo.toml` (workspace + each crate) `version = "2.0.0"`.
* `python/wolfxl/__init__.py` — verify `__version__` reads from
  `wolfxl._rust._VERSION` which reads from `CARGO_PKG_VERSION`
  (lesson #15 — Cargo.toml is canonical).
* PyPI classifier stays `Development Status :: 5 - Production/Stable`
  (promoted in Sprint Ξ).

### 2.5 Tag + publish

* `git tag v2.0.0` after `cargo test --workspace` and `pytest`
  green.
* `maturin publish --release --strip` for x86_64-linux,
  aarch64-linux, x86_64-darwin, aarch64-darwin, x86_64-windows.
* Verify `pip install wolfxl==2.0.0` works on a fresh venv on
  each target.

### 2.6 ExcelBench dashboard

* Run `WOLFXL_TEST_EPOCH=0 python scripts/bench-all.py
  --include-pivot --output benchmark-results-v2.0.json`.
* Push refreshed dashboard.

## 3. Verification Matrix

This RFC has no Rust/Python code; the matrix is doc-validation:

1. `mkdocs serve` builds without errors.
2. Every code snippet in `docs/migration/openpyxl-migration.md`
   parses + runs against an installed v2.0.0 wheel
   (`pytest docs/migration/test_snippets.py`).
3. `tests/parity/KNOWN_GAPS.md` "construction-side gaps" list
   is empty.
4. `CHANGELOG.md` and `docs/release-notes-2.0.md` reference the
   actual final commit SHAs (Pod-ε scaffolds with `<!-- TBD: SHA
   -->` markers; integrator finalizes during the merge commit).
5. `git tag --list | grep v2.0.0` succeeds.
6. `pip install wolfxl==2.0.0` on fresh venv works on all 5
   targets.

## 4. Acceptance

- README.md "Full openpyxl replacement" headline live.
- Compatibility matrix shows pivots ✅.
- KNOWN_GAPS.md "Out of scope" only lists slicers + calc fields
  + GroupItems + OLAP + pivot-styling.
- Tag `v2.0.0` cut.
- PyPI shows `wolfxl==2.0.0` on Production/Stable classifier.
- Launch posts go live (HN → +30 min Twitter → r/Python → dev.to
  → GH Discussions).

## Acceptance

(Filled in after Shipped.)

- Commit: `<TBD: SHA>` — Sprint Ν Pod-ε merge
- Tag: `v2.0.0` at `<TBD: SHA>`
- PyPI: `wolfxl==2.0.0` published
- Launch: <TBD>
