# Sprint Ξ ("Xi") — v1.7.0 launch slice (no pivot tables)

**Status**: In progress (kicked off post-v1.6.1, 2026-04-27).
**Target tag**: `v1.7.0`.
**Calendar**: ~2 weeks parallel-pod.
**Predecessor**: v1.6.1 (Sprint Μ-prime — chart contract reconciliation).
**Successor**: v2.0.0 (Sprint Ν — pivot tables + pivot charts + public-launch-with-pivots).

## Why this sprint exists

After v1.6.1 the openpyxl-parity construction surface is exhaustively
shipped **except for pivot tables + pivot charts**. The original
roadmap routed straight to Sprint Ν (v2.0.0 + public launch). The
user picked **Option C** (this sprint) on 2026-04-27 to:

1. Burn down the small chart-stack debt that v1.6.1 left behind, so
   Sprint Ν starts from a clean slate.
2. Ship a **public-launch-ready v1.7** that earns the **"openpyxl
   parity for the 95th-percentile case (pivot tables preserved on
   round-trip but not yet constructible)"** marketing claim.
3. Defer pivot tables to v2.0.0 / Sprint Ν without blocking launch
   on it.

## In-scope

### Code (Pod-α)

1. **Version sync** — `pyproject.toml` (currently `0.5.0`) and
   `Cargo.toml` package version both bump to `1.7.0` (and stay
   tracked from now on). `wolfxl.__version__` re-reads from
   `CARGO_PKG_VERSION` via the existing `src/lib.rs:41`
   `m.add("__version__", env!("CARGO_PKG_VERSION"))?` line, so the
   single source of truth is `Cargo.toml [package].version`.
2. **`Worksheet.remove_chart(chart)` + `Worksheet.replace_chart(old, new)`**
   (RFC-050). Mirror `add_chart` but mutate the pending list in
   write mode and queue a chart-removal request in modify mode.
   Deferred from v1.6.1 release notes.
3. **`chart.title = RichText(...)`** support (closes the only remaining
   `xfail` in `tests/test_charts_write.py::test_line_chart_title_rich_text`).
4. **RFC-046 §13 — legacy chart-dict key sunset**. Pod-α′'s
   `parse_graphical_properties` accepts both §10
   (`solid_fill`, `ln: {solid_fill, w_emu, prst_dash}`) and legacy
   (`fill_color`, `line_color`, `line_dash`, `line_width_emu`).
   v1.7 emits `DeprecationWarning` on legacy-key use; sunset and
   remove in v2.0. Already gone from the Python emitter (per
   2026-04-27 grep); only the Rust parser carries the shim.

### Docs (Pod-β)

5. **`docs/migration/` overhaul** (RFC-051). Three current files
   (`openpyxl-migration.md`, `compatibility-matrix.md`,
   `legacy-shim.md`) are 38 / 67 / N lines and cover the
   pre-1.0 read-side surface only. Rewrite to v1.7 status with
   construction-side walkthroughs:
   * Charts (8 2D + 8 3D/Stock/Surface families)
   * Images (RFC-045)
   * Encrypted reads + writes (RFC-042 + RFC-044)
   * Streaming reads (RFC-041)
   * Structural ops (RFC-030/031/034/035/036)
   * Modify-mode (T1.5 — RFC-020/021/022/023/024/025/026)
   * `.xlsb` / `.xls` reads (RFC-043)
   * Pivot tables — explicit "preserved on round-trip; not yet
     constructible (Sprint Ν / v2.0.0)" callout.

6. **`docs/migration/compatibility-matrix.md`**. Rewrite the matrix
   with the ~25-row openpyxl surface (vs. 16 today) and flip
   "Full openpyxl API parity: Partial" to "Full openpyxl API parity
   (construction): Yes (except pivot tables)" with the table
   showing the supported families. Match the wording in
   `tests/parity/KNOWN_GAPS.md`.

### Perf (Pod-γ)

7. **`docs/performance/` refresh** (RFC-052).
   * `benchmark-results.md` — replace the mostly-empty "see
     ExcelBench dashboard" page with v1.7 read / write / modify-mode
     numbers on a 1k / 10k / 100k row matrix vs openpyxl 3.1.5,
     pulled into the docs and linked.
   * `methodology.md` — verify still accurate; add a section on
     chart-construction perf (new in v1.6+).
   * `run-on-your-files.md` — add chart-construction and
     copy-worksheet bench harnesses alongside the existing
     read/write/modify ones.

### Launch (Pod-δ)

8. **`Plans/launch-posts.md`** (RFC-053). The INDEX has referenced
   this file as "kept in conversation transcript" since
   day one. Materialize as a real document with drafts for:
   * **HN** post (technical, "Show HN: WolfXL — openpyxl-compatible
     Excel I/O with a Rust backend")
   * **r/Python**
   * **Twitter/X** thread
   * **dev.to** long-form ("How we shipped 23 RFCs in 6 weeks
     with parallel-pod orchestration")
   * **GitHub Discussions** announcement post
   * Each draft includes embedded charts, the comparison-matrix
     screenshot, and links to the new docs.

### Release artifacts (Integrator)

9. **`CHANGELOG.md`** prepend v1.7.0 entry.
10. **`docs/release-notes-1.7.md`** — full release notes following
    the v1.6.x format.
11. **`Plans/rfcs/INDEX.md`** — add 050 / 051 / 052 / 053 rows;
    bump status table; flip 046's "Unblocks" to remove the
    "v1.6.1 follow-up" reference.
12. **Tag `v1.7.0`** after `cargo test --workspace` and `pytest`
    are green.

## Out of scope (deferred)

- **Pivot tables + pivot charts** — Sprint Ν / v2.0.0.
- **Combination charts** (multi-plot — bar + line on shared axes).
  Tracked as a v1.6.x follow-up; defer to post-v1.7 unless the
  launch surfaces specific demand.
- **OpenDocument (`.ods`)** — out of scope; not on roadmap.
- **CSS-style chart theme tokens** — not yet specified.

## RFCs

Four micro-RFCs land in this sprint. They are S/M sized and follow
the standard `Plans/rfcs/000-template.md`:

| RFC | Title | Pod | Estimate |
|---|---|---|---|
| 050 | `Worksheet.remove_chart` / `replace_chart` + RichText title | α | M |
| 051 | docs/migration overhaul for v1.7 | β | M |
| 052 | docs/performance refresh for v1.7 | γ | M |
| 053 | Public launch posts + materialise `Plans/launch-posts.md` | δ | S |

## Pod plan

Sequential merge order — α first (carries the version bump and
chart-API additions that β/γ/δ reference), then β/γ/δ in any order.

| Pod | Branch | Deliverable |
|---|---|---|
| α | `feat/sprint-xi-pod-alpha` | Version sync; `remove_chart` / `replace_chart`; RichText title; legacy-key DeprecationWarning |
| β | `feat/sprint-xi-pod-beta`  | `docs/migration/` rewrite + Compatibility Matrix v1.7 |
| γ | `feat/sprint-xi-pod-gamma` | `docs/performance/` refresh with v1.7 numbers + harness expansions |
| δ | `feat/sprint-xi-pod-delta` | `Plans/launch-posts.md` drafts |
| Integrator | `feat/native-writer` | CHANGELOG / release-notes-1.7 / INDEX / tag |

For Sprint Ξ we don't need the full worktree-fanout pattern — the
integrator can author α's code changes inline (small surface),
β/γ/δ are docs-only and can land as sequential commits on the
integration branch.

## Lessons applied

- Pre-dispatch contract spec (Sprint Μ-prime lesson #12): RFC-050
  spec written into the RFC document BEFORE Pod-α starts.
- Doc pod can scaffold with `<!-- TBD: SHA -->` markers (lesson
  #3); release notes drafted before tag.
- Ratchet flip is post-merge integrator work (lesson #7) — RFC-050
  adds parity ratchet entries for `Worksheet.remove_chart` /
  `replace_chart`.
- Version bump must touch BOTH `Cargo.toml` (Rust source of truth
  for `__version__`) and `pyproject.toml` (PyPI source of truth)
  — neither alone is sufficient.

## Acceptance criteria

1. `python3 -c "import wolfxl; print(wolfxl.__version__)"` reports
   `1.7.0`.
2. `cargo test --workspace --exclude wolfxl` green.
3. `pytest tests/` ≥ v1.6.1 baseline (1340+) with the lone v1.6.1
   xfail (RichText title) flipped to passing.
4. `pytest tests/parity/` green; new ratchet entries for
   `remove_chart` / `replace_chart` flipped post-merge.
5. `docs/migration/` accurately reflects v1.7 surface.
6. `docs/performance/` carries v1.7 benchmark numbers.
7. `Plans/launch-posts.md` exists with at least 3 channel drafts.
8. `docs/release-notes-1.7.md` exists.
9. `CHANGELOG.md` 1.7.0 entry prepended.
10. `git tag v1.7.0` cut.
