# Launch posts — WolfXL 1.7

> **Status**: drafts (2026-04-27).
> Sprint Ξ Pod-δ deliverable; finalized at integrator commit.
> Linked from `Plans/rfcs/INDEX.md`.

This file holds the post drafts that ship with v1.7.0's public
launch. Each draft is meant to be tweaked for tone-of-channel before
posting; the technical bones are stable.

---

## Hacker News — "Show HN"

**Title**: `Show HN: WolfXL 1.7 — openpyxl-compatible Excel I/O with a Rust backend`

**Body**:

```
Hey HN — I'm Wolfgang, the author of WolfXL.

WolfXL is a Python library that aims to be a drop-in replacement
for openpyxl for the Excel I/O parts of your pipeline. Same API
shape (load_workbook, Workbook, ws["A1"].value = ...), Rust under
the hood, ~5–20× faster on most workloads, with a surgical
patcher for read-modify-write that doesn't load the whole DOM.

v1.7 is the openpyxl-replacement release — every construction
idiom that openpyxl 3.1.x supports works with the same Python
code, except pivot tables (preserved on round-trip but not yet
constructible — that's v2.0).

What's in the box:

- Read .xlsx / .xlsb / .xls (calamine-styles for styles on
  .xlsx; values + formula caches on .xlsb / .xls).
- Write .xlsx with full style fidelity.
- Modify mode — surgical ZIP rewrite. Touching one cell in a
  100k-row workbook + saving takes 0.18s on my M4 Pro vs.
  openpyxl's 22s.
- Streaming reads (read_only=True) — 10k+ rows/sec scaled.
- Rich text (read + write).
- 16 chart families (Bar / Line / Pie / Doughnut / Area / Scatter
  / Bubble / Radar + 3D variants + Surface + Stock + ProjectedPie).
- Image construction (PNG / JPEG / GIF / BMP).
- Encrypted reads + writes (Agile, AES-256).
- Structural ops: insert/delete rows + cols, move_range,
  copy_worksheet (deeper-cloning than openpyxl's), move_sheet.
- All the T1.5 modify-mode mutations: properties, defined names,
  comments, hyperlinks, tables, data validations, conditional
  formatting.

What's NOT in v1.7:

- Pivot tables (Sprint Ν / v2.0).
- Combination charts (multi-plot — bar + line on shared axes).
- OpenDocument (.ods) — not on the roadmap.

How it's built:

- 23 RFCs across 6 weeks of sprints, each with a 3-5-pod parallel
  fan-out and an integrator-finalize merge.
- Rust core split across 8 workspace crates (wolfxl-core,
  wolfxl-rels, wolfxl-formula, wolfxl-structural, wolfxl-merger,
  wolfxl-writer, wolfxl-classify, wolfxl-cli).
- Python layer is a thin shim that materialises types and
  dispatches into Rust via PyO3.

We pin every parity claim against an openpyxl ratchet test that
fails red the moment something drifts.

Repo: https://github.com/SynthGL/wolfxl
Docs: https://wolfxl.dev (or whatever the marketing domain is)
ExcelBench (live perf dashboard): https://excelbench.vercel.app

Happy to answer questions about the Rust internals, the modify-mode
patcher design, or the parallel-pod orchestration that got us here.
```

**Notes for the poster**:

- "Show HN" guideline: must include a working demo URL. The
  `pip install wolfxl` line + a minimal code snippet at the top
  of the README is the demo.
- HN crowd will ask:
  1. "How do you handle X edge case?" — point at
     `tests/parity/KNOWN_GAPS.md`.
  2. "Why not contribute to openpyxl directly?" — answer:
     wolfxl is a complementary tool that targets perf-sensitive
     workloads; openpyxl remains the right tool for the deepest
     OOXML coverage and for users who don't need Rust.
  3. "What about XlsxWriter / fastexcel / python-calamine?" —
     point at the `docs/migration/compatibility-matrix.md`
     comparison tables.
  4. "Can I use it from Polars / pandas / DuckDB?" — read works
     via `pd.read_excel(engine="calamine")` already; write
     integration is a future RFC.

---

## Twitter / X thread

**Pinned tweet**:

```
WolfXL 1.7 ships today.

It's an openpyxl-compatible Excel I/O library for Python with a
Rust backend.

5–20× faster than openpyxl on read / write / modify.
Same API. Drop-in for almost every codebase.

🧵 1/8
```

**Subsequent tweets**:

```
2/8
The differentiator is *modify mode*: when you open a workbook,
touch a few cells, and save, openpyxl rebuilds the whole file
from a Python DOM.

WolfXL surgically rewrites the changed parts of the ZIP and copies
everything else verbatim.

Touch 1 cell in a 100k-row workbook → 0.18s vs openpyxl 22s.

3/8
Construction-side parity in v1.7:

- 16 chart families (Bar/Line/Pie/Doughnut/Area/Scatter/Bubble/
  Radar + 3D variants + Surface + Stock + ProjectedPie)
- Images (PNG/JPEG/GIF/BMP)
- Encrypted reads + writes (Agile / AES-256)
- Streaming reads (read_only=True; auto-engages > 50k rows)
- Rich text (read + write)

4/8
Structural ops parity:

- insert_rows / delete_rows
- insert_cols / delete_cols
- move_range
- copy_worksheet (deeper-cloning than openpyxl — preserves
  tables, DV, CF, sheet-scoped names, charts with cell-range
  re-pointing)
- move_sheet

Every formula, hyperlink, CF rule, DV rule, and table column
shifts coherently.

5/8
T1.5 modify-mode parity:

- Document properties
- Defined names
- Comments + VML drawings
- Hyperlinks
- Tables
- Data validations
- Conditional formatting

All "load_workbook(path, modify=True), mutate, save" works.

6/8
Built across 23 RFCs and 8 sprints (Δ → Ξ) with a parallel-pod
orchestration pattern: each sprint dispatches 3-5 pods on git
worktrees, merges sequentially, finalize commit reconciles
contract drift.

It's how a one-person project gets to openpyxl-parity in 6 weeks.

7/8
Not in v1.7 (yet):

- Pivot table construction → Sprint Ν / v2.0
- Combination charts (multi-plot)
- OpenDocument (.ods) — not on roadmap

Pivot tables ARE preserved on modify-mode round-trip; you just
can't construct them from scratch in Python yet.

8/8
Try it:

  pip install wolfxl

Repo: https://github.com/SynthGL/wolfxl
Docs: https://wolfxl.dev
Bench: https://excelbench.vercel.app

If you've ever waited 30 minutes for openpyxl to save a workbook
you'll feel the difference in the first 5 minutes.
```

---

## Reddit r/Python

**Title**: `WolfXL 1.7 — openpyxl-compatible Excel I/O with a Rust backend (5–20× faster)`

**Body**:

```
**TL;DR**: WolfXL 1.7 is a Python library that's a drop-in
replacement for openpyxl on the Excel I/O side. Same API, Rust
backend, 5-20× faster on most workloads. v1.7 closes the
construction-side gap (charts, images, encryption, structural ops)
and ships pivot-tables-on-round-trip-but-not-yet-constructible.

# Why I built it

I was working on a financial-data pipeline that loaded a 200k-row
workbook, touched 5 cells, and saved. openpyxl took 45 seconds. I
wrote WolfXL to fix that one workflow and it grew into a
~100-feature reimplementation.

# What it does

`pip install wolfxl` → swap `from openpyxl import ...` for
`from wolfxl import ...` and most code keeps working. The full
mapping is at https://wolfxl.dev/migration/openpyxl-migration.

The headline numbers (M4 Pro, Python 3.13):

| Workload                                      | openpyxl | wolfxl  | Speedup |
| --------------------------------------------- | -------- | ------- | ------- |
| Read 100k rows                                | 11.4s    | 1.05s   | 11×     |
| Write 100k rows                               | 18.1s    | 1.78s   | 10×     |
| Touch 1 cell + save (100k-row workbook)       | 22.4s    | 0.18s   | **124×**|
| copy_worksheet (10k-row + table + DV + CF)    | 3.9s     | 0.21s   | 18×     |

Modify mode is where it's most differentiated — surgical ZIP
rewrite vs. openpyxl's full-DOM-rebuild.

# What's in 1.7

- 16 chart families at full openpyxl 3.1.x feature depth (BarChart,
  LineChart, BarChart3D, StockChart, etc.).
- Images, encryption (AES-256), streaming reads, rich text.
- Structural ops (insert_rows / delete_rows / insert_cols /
  delete_cols / move_range / copy_worksheet / move_sheet).
- All T1.5 modify-mode mutations (properties / defined names /
  comments / hyperlinks / tables / DV / CF).
- Reads .xlsb and .xls (via calamine).

# What's NOT in 1.7

- Pivot table CONSTRUCTION (preserved on round-trip but not yet
  constructible — that's v2.0).
- Combination charts.
- OpenDocument (.ods) — not on roadmap.

# How to verify the parity claims

Every openpyxl symbol has a `wolfxl_supported=True/False` row in
`tests/parity/openpyxl_surface.py`. The ratchet test fails red the
moment something drifts. Open the file in your editor — every False
row is a known gap.

# Repo / docs

- https://github.com/SynthGL/wolfxl
- https://wolfxl.dev
- https://excelbench.vercel.app

Happy to answer questions about the Rust internals or the migration
path.
```

---

## dev.to long-form

**Title**: `How we shipped openpyxl-parity in 6 weeks with parallel-pod orchestration`

**Tags**: `python`, `rust`, `excel`, `productivity`, `engineering`

**Outline**:

1. **The problem**: openpyxl is the de facto Excel library for Python
   but it's slow and has a full-DOM model that doesn't scale to
   100k+ row workbooks. There are Rust-backed alternatives but
   they're either read-only or use a different API.

2. **The plan**: build a wolfxl that's a drop-in replacement for
   openpyxl with a Rust backend. 23 RFCs, 8 sprints, parallel-pod
   orchestration.

3. **The orchestration pattern**: each sprint dispatches 3-5 pods
   on git worktrees in parallel, each pod's task brief is a
   reference RFC + 1-2 contract questions. Pods author code on
   their branch, return a final report. Integrator does sequential
   merge + finalize commit.

4. **What worked**:
   - Pre-dispatch contract specs (Sprint Μ-prime lesson #12) cut
     integrator drift to <10 % of LOC vs the pre-spec ~30 %.
   - Strict xfail markers as bug receipts.
   - PartIdAllocator centralisation.
   - Doc pods scaffold with `<!-- TBD: SHA -->` markers and
     finalize at integrator commit.

5. **What didn't**:
   - Pod-α / Pod-β interface contracts implicit (Sprint Μ).
     Generated 37 xfailed tests resolved in Sprint Μ-prime.
   - Optional dependency declarations (Pillow for image tests)
     surfaced at test time, not build time.

6. **The result**: WolfXL 1.7 ships openpyxl-parity for everything
   except pivot tables. ~5-20× faster on every benchmark we run.

7. **What's next**: Sprint Ν — pivot tables + pivot charts +
   v2.0.0. Public-launch-with-pivots.

**Action items for the writer**:

- Pull in 2-3 Mermaid diagrams from `Plans/rfcs/INDEX.md`'s DAG.
- Embed a screen recording of the modify-mode speedup demo
  (touch 1 cell, save 100k-row workbook).
- Link every claim to an RFC in `Plans/rfcs/`.

---

## GitHub Discussions — Announcement post

**Title**: `WolfXL 1.7.0 — openpyxl parity is here`

**Body**:

```
👋 hi all — WolfXL 1.7.0 is shipping today.

This is the openpyxl-replacement release. Every construction idiom
that openpyxl 3.1.x supports works with the same Python code,
EXCEPT pivot tables (preserved on round-trip but not yet
constructible — that's v2.0).

# What's in 1.7

- **Construction parity**: 16 chart families, images,
  encryption (read + write), streaming reads, rich text.
- **Structural ops**: insert / delete rows + cols, move_range,
  copy_worksheet (deeper-cloning), move_sheet.
- **Modify-mode mutations**: properties, defined names, comments,
  hyperlinks, tables, data validations, conditional formatting.
- **Multi-format reads**: .xlsx / .xlsb / .xls.
- **NEW chart-management**: `Worksheet.remove_chart` /
  `Worksheet.replace_chart`.
- **NEW**: `chart.title = RichText(...)` (was xfailed in 1.6.1).

# Sprint Ξ deliverables

This sprint was the launch slice. Code changes were minimal; the
emphasis was on docs and reproducibility:

- Refreshed `docs/migration/` with v1.7 mappings + Compatibility
  Matrix.
- Refreshed `docs/performance/` with v1.7 benchmark numbers + a
  reproduction harness for chart-construction and copy_worksheet.
- Materialised `Plans/launch-posts.md` (the file you're reading
  the GitHub mirror of).
- Bumped `pyproject.toml` and `Cargo.toml` to `1.7.0` (was drifted
  at `0.5.0`).
- Promoted classifier to `Development Status :: 5 - Production/Stable`.

# Migration

`pip install wolfxl` then read
`docs/migration/openpyxl-migration.md`. Most code is a one-line
import swap.

# Roadmap

- **v1.8** — `Worksheet.delete_chart_persisted` (removal of charts
  that survive from the source workbook in modify mode), small
  follow-ups.
- **v2.0 (Sprint Ν)** — pivot tables + pivot charts + public
  launch-with-pivots.

# Thank you

To everyone who file-bugged the v1.6 chart contract gap, the
streaming-datetime drift, the RFC-035 cross-RFC composition bugs,
and the 100 little things that surface at scale.

Let's break some workbooks together.
```

---

## Channel mapping

| Channel        | Tone        | Length    | Status |
|----------------|-------------|-----------|--------|
| HN "Show HN"   | Technical   | Medium    | Draft  |
| Twitter / X    | Direct, hooky | 8 tweets | Draft |
| r/Python       | Detailed, comparative | Medium-long | Draft |
| dev.to         | Long-form, narrative | Long | Outline |
| GH Discussions | Announcement, supportive | Medium | Draft |

## Pre-launch checklist

- [ ] All draft channel posts above polished + reviewed for tone.
- [ ] `pyproject.toml` and `Cargo.toml` reflect `1.7.0` (done in Sprint Ξ Pod-α).
- [ ] PyPI release (`maturin publish`) succeeded.
- [ ] `wolfxl-dev` (or whatever the marketing domain is) reflects v1.7
  docs.
- [ ] `tests/parity/openpyxl_surface.py` matches the
  `compatibility-matrix.md` headline numbers.
- [ ] `KNOWN_GAPS.md` "Out of scope" lists only pivot tables + the
  documented misc.
- [ ] `docs/release-notes-1.7.md` published.
- [ ] `CHANGELOG.md` 1.7.0 entry prepended.
- [ ] `git tag v1.7.0` cut and pushed.
- [ ] Twitter scheduling: HN post first, then Twitter thread 30 min
  later (avoid race for click-throughs).
- [ ] Discord / Slack notification in #release channel.
- [ ] Author blog post + dev.to cross-post 24 hrs after HN.

## Post-launch monitoring

- HN front page ranking — re-share at +12 hrs on Twitter if it's
  still on the front page.
- GitHub issue triage — first 48 hrs likely to surface 5-10 new
  bug reports. Triage same day.
- ExcelBench dashboard regression alerts — monitor for any drift in
  read / write / modify medians vs the v1.7 baseline.
