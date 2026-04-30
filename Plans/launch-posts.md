# Launch posts — WolfXL 2.0

> **Status**: internal drafts only (Sprint Ν Pod-ε ε.6, finalize from
> Sprint Ξ scaffold).
> Linked from `Plans/rfcs/INDEX.md` (RFC-054) and
> `docs/release-notes-2.0.md`.
> SHAs and benchmark numbers are `<!-- TBD -->` until the release
> truth pass reconciles them against the final audited commit and
> release artifacts.
> Post-PR #23 audit note: release/public launch remains frozen.
> Older draft copy described slicers, calculated fields/items,
> GroupItems, and chart display units/dPt overrides as future work;
> those are now implemented, but this file still needs a final manual
> truth pass before publication.

This file holds launch-post drafts for the v2.0.0 public release. They
must not be posted until benchmark numbers, release SHAs, clean
artifact smoke, and the final public-claim wording pass are
complete.

The body drafts below intentionally preserve some bolder launch phrasing for
editing context; treat any benchmark, replacement, first, or only wording as
unsafe until the benchmark and ecosystem-claim proof gates are complete.
Pivot construction examples must use
`load_workbook(..., modify=True)`; the current v2.0 API intentionally
routes cache/table construction through the patcher, not fresh
`Workbook()` write mode.

---

## Hacker News — "Show HN"

**Title**: `Show HN: WolfXL 2.0 — Python OOXML pivot-table construction with pre-aggregated records`

**Body**:

```
Hey HN — I'm Wolfgang, the author of WolfXL.

WolfXL is a Python library aiming at openpyxl-shaped Excel
automation. Same API shape (load_workbook, Workbook,
ws["A1"].value = ...), Rust under the hood, with a surgical
patcher for read-modify-write that doesn't load the whole DOM.
The speedup headline is withheld until the release benchmark
refresh lands.

v2.0 closes the tracked construction-side parity roadmap, including
pivot tables. The final release copy still needs the exact
"replacement" wording audited against the remaining caveats below.

The pivot-table piece is the differentiator. WolfXL constructs
pivot tables with a pre-aggregated `pivotCacheRecords` snapshot —
the saved workbook opens in Excel / LibreOffice / openpyxl with
the pivot's data already populated, no Excel-side refresh
round-trip required.

(Caveat: openpyxl preserves pivot tables on round-trip but does
not provide a Python-side constructor that emits the records
part. XlsxWriter doesn't support pivots at all.)

What you can do in v2.0:

  import wolfxl
  from wolfxl.chart import Reference
  from wolfxl.pivot import PivotCache, PivotTable

  wb = wolfxl.load_workbook("source-data.xlsx", modify=True)
  ws = wb.active
  src = Reference(ws, min_col=1, min_row=1, max_col=4, max_row=100)
  cache = wb.add_pivot_cache(PivotCache(source=src))
  pt = PivotTable(
      cache=cache, location="F2",
      rows=["region"], cols=["quarter"], data=[("revenue", "sum")],
  )
  ws.add_pivot_table(pt)
  wb.save("pivot.xlsx")

The saved workbook opens in Excel without a refresh.

What's also in the box (carried over from v1.7):

- Read .xlsx natively with styles and workbook metadata; read .xlsb / .xls
  through Calamine-backed value and formula-cache compatibility paths.
- Write .xlsx with full style fidelity.
- Modify mode — surgical ZIP rewrite. Touching one cell in a
  100k-row workbook + saving takes <!-- TBD: BENCHMARK NUMBERS -->
  on M4 Pro vs. openpyxl's <!-- TBD: BENCHMARK NUMBERS -->.
- Streaming reads (read_only=True) — auto-engages > 50k rows.
- Rich text (read + write).
- 16 chart families (Bar / Line / Pie / Doughnut / Area / Scatter
  / Bubble / Radar + 3D variants + Surface + Stock + ProjectedPie).
- Pivot-chart linkage (chart.pivot_source = pt) on all 16
  families — emits <c:pivotSource> + per-series <c:fmtId>.
- Image construction (PNG / JPEG / GIF / BMP).
- Encrypted reads + writes (Agile, AES-256).
- Structural ops: insert/delete rows + cols, move_range,
  copy_worksheet (deeper-cloning than openpyxl's; pivot-bearing
  sheets clone cleanly in v2.0), move_sheet.
- All the T1.5 modify-mode mutations: properties, defined names,
  comments, hyperlinks, tables, data validations, conditional
  formatting.

Still partial / not in v2.0:

- Pivot-table styling beyond the named-style picker remains partial
  (PivotArea formats and pivot-scoped CF exist; broader theme/banded
  styling is still limited).
- In-place pivot edits in modify mode — v2.2.
- Combination charts (multi-plot — bar + line on shared axes).
- OpenDocument (.ods) — not on the roadmap.

How it's built:

- 24 RFCs across 9 sprints (Δ → Ν), each with a 3-5-pod
  parallel fan-out and an integrator-finalize merge.
- Pre-dispatch §10 contract specs (Sprint Μ-prime lesson #12)
  for every cross-pod data structure.
- Rust core split across 9 workspace crates (wolfxl-core,
  wolfxl-rels, wolfxl-formula, wolfxl-structural, wolfxl-merger,
  wolfxl-writer, wolfxl-classify, wolfxl-cli, **wolfxl-pivot**
  new in v2.0).
- Python layer is a thin shim that materialises types and
  dispatches into Rust via PyO3.

Every parity claim is pinned against an openpyxl ratchet test
that fails red the moment something drifts.

Repo: https://github.com/SynthGL/wolfxl
Docs: https://wolfxl.dev
ExcelBench (live perf dashboard): https://excelbench.vercel.app
Release notes: https://github.com/SynthGL/wolfxl/blob/v2.0.0/docs/release-notes-2.0.md

Happy to answer questions about the pivot-construction internals,
the modify-mode patcher design, or the parallel-pod orchestration
that got us here.
```

**Notes for the poster**:

- "Show HN" guideline: must include a working demo URL. The
  `pip install wolfxl` line plus the modify-mode pivot snippet at the
  top of the README is the demo.
- HN crowd will ask:
  1. "Is the pivotCacheRecords ecosystem claim proven?" — point at
     [openpyxl issue tracker / source for pivot.cache.CacheDefinition](https://foss.heptapod.net/openpyxl/openpyxl):
     the openpyxl `CacheDefinition` class only round-trips
     records; no `to_tree()` path generates a fresh records part.
     XlsxWriter's docs explicitly call out "pivot tables not
     supported".
  2. "How do you handle X edge case?" — point at
     `tests/parity/KNOWN_GAPS.md` "Out of scope" and the
     post-PR #23 audit notes.
  3. "Why not contribute pivot construction to openpyxl?" —
     answer: wolfxl is a complementary tool that targets
     perf-sensitive workloads with a Rust backend; the
     pivot-construction layer is a 9-crate Rust effort that
     doesn't drop neatly into openpyxl's pure-Python design.
  4. "What about XlsxWriter / fastexcel / python-calamine?" —
     point at the `docs/migration/compatibility-matrix.md`
     comparison tables (now showing wolfxl as the pivot-construction
     entry in the compared set).
  5. "Does the records snapshot stay in sync if I edit the
     source data later?" — no, mirrors openpyxl's behaviour.
     Excel rebuilds the cache on next "Refresh"; until then the
     records are the v2.0-emitted snapshot. Document this in
     the FAQ.

---

## Twitter / X thread

**Pinned tweet**:

```
WolfXL 2.0 ships today.

It's a Python library aiming at openpyxl-shaped Excel automation,
with a Rust backend and surgical modify-mode patcher.

v2.0 closes the tracked construction-side parity roadmap.
Pivot tables included.

🧵 1/8
```

**Subsequent tweets**:

```
2/8
The headline: WolfXL constructs pivot tables with pre-aggregated
records.

Open the saved workbook in Excel / LibreOffice / openpyxl —
the pivot's data is already populated. No refresh-on-open.

(openpyxl preserves on round-trip but doesn't construct.
XlsxWriter doesn't support pivots at all.)

3/8
Pivot construction:

  wb = wolfxl.load_workbook("source-data.xlsx", modify=True)
  ws = wb.active
  src = Reference(ws, min_col=1, min_row=1,
                  max_col=4, max_row=100)
  cache = wb.add_pivot_cache(PivotCache(source=src))
  pt = PivotTable(
      cache=cache, location="F2",
      rows=["region"], cols=["quarter"],
      data=[("revenue", "sum")],
  )
  ws.add_pivot_table(pt)
  wb.save("pivot.xlsx")

4/8
Pivot-chart linkage:

  chart = BarChart()
  chart.pivot_source = pt
  ws.add_chart(chart, "F18")

Excel renders it as a true PivotChart — right-click shows
PivotChart actions; chart.pivotSource.name reads back via
openpyxl.

5/8
Other v2.0 deltas vs v1.7:

- copy_worksheet now deep-clones pivot-bearing sheets. Cache
  aliased, table fresh-id'd, source-range hint re-pointed.
- RFC-046 §13 legacy chart-dict shim removed (deprecated in
  v1.7, removed v2.0).
- Migration guide + Compatibility Matrix updated.
- README rewrite: drop-in-oriented compatibility, pivot coverage,
  and benchmark claims gated on audited numbers.

6/8
The differentiator across v1.x → v2.0 is *modify mode*. When
you open a workbook, touch a few cells, and save, openpyxl
rebuilds the whole file from a Python DOM.

WolfXL surgically rewrites the changed parts of the ZIP and
copies everything else verbatim.

Touch 1 cell in a 100k-row workbook → <!-- TBD: BENCHMARK
NUMBERS --> vs openpyxl <!-- TBD: BENCHMARK NUMBERS -->.

7/8
Still partial / not in v2.0:

- Pivot-table styling beyond named-style picker remains partial
- In-place pivot edits in modify mode → v2.2
- Combination charts (multi-plot)
- OpenDocument (.ods) — not on roadmap

8/8
Try it:

  pip install wolfxl

Repo: github.com/SynthGL/wolfxl
Docs: wolfxl.dev
Bench: excelbench.vercel.app
Release notes: docs/release-notes-2.0.md

If you've ever waited 30 minutes for openpyxl to save a
pivot-bearing workbook, you'll feel the difference in the
first 5 minutes.
```

---

## Reddit r/Python

**Title**: `WolfXL 2.0 — Python OOXML pivot-table construction with pre-aggregated records`

**Body**:

```
**TL;DR**: WolfXL 2.0 is a Python library aiming at openpyxl-shaped
Excel automation with a Rust backend. v2.0 closes the tracked
construction-side parity roadmap, including pivot tables. Benchmark
claims and the final "replacement" wording remain gated on the release
truth pass.

# What's new in 2.0

WolfXL constructs pivot tables with pre-aggregated
`pivotCacheRecords`. The saved
workbook opens in Excel / LibreOffice / openpyxl with the
pivot's data already populated — no Excel-side refresh
round-trip.

(Caveat: openpyxl preserves pivot tables on round-trip but
does NOT provide a Python-side constructor for the records
snapshot. XlsxWriter doesn't support pivots at all.)

## Pivot construction:

```python
import wolfxl
from wolfxl.chart import Reference
from wolfxl.pivot import PivotCache, PivotTable

wb = wolfxl.load_workbook("source-data.xlsx", modify=True)
ws = wb.active

src = Reference(ws, min_col=1, min_row=1, max_col=4, max_row=100)
cache = wb.add_pivot_cache(PivotCache(source=src))
pt = PivotTable(
    cache=cache, location="F2",
    rows=["region"], cols=["quarter"],
    data=[("revenue", "sum")],
)
ws.add_pivot_table(pt)
wb.save("pivot.xlsx")
```

## Link a chart to the pivot:

```python
from wolfxl.chart import BarChart

chart = BarChart()
chart.pivot_source = pt
ws.add_chart(chart, "F18")
```

# Why I built it

I was working on a financial-data pipeline that loaded a 200k-row
workbook, touched 5 cells, and saved. openpyxl took 45 seconds. I
wrote WolfXL to fix that one workflow and it grew into a
~100-feature reimplementation — with pivot tables now closing the
last construction-side gap.

# What it does (full list)

`pip install wolfxl` → swap `from openpyxl import ...` for
`from wolfxl import ...` and most code keeps working. The full
mapping is at https://wolfxl.dev/migration/openpyxl-migration.

The headline numbers (M4 Pro, Python 3.13):

| Workload                                      | openpyxl | wolfxl  | Speedup |
| --------------------------------------------- | -------- | ------- | ------- |
| Read 100k rows                                | <!-- TBD: BENCHMARK NUMBERS --> | <!-- TBD: BENCHMARK NUMBERS --> | <!-- TBD: BENCHMARK NUMBERS --> |
| Write 100k rows                               | <!-- TBD: BENCHMARK NUMBERS --> | <!-- TBD: BENCHMARK NUMBERS --> | <!-- TBD: BENCHMARK NUMBERS --> |
| Touch 1 cell + save (100k-row workbook)       | <!-- TBD: BENCHMARK NUMBERS --> | <!-- TBD: BENCHMARK NUMBERS --> | <!-- TBD: BENCHMARK NUMBERS --> |
| copy_worksheet (10k-row + table + DV + CF)    | <!-- TBD: BENCHMARK NUMBERS --> | <!-- TBD: BENCHMARK NUMBERS --> | <!-- TBD: BENCHMARK NUMBERS --> |
| Pivot construction (100k source rows)         | <!-- TBD: BENCHMARK NUMBERS --> | <!-- TBD: BENCHMARK NUMBERS --> | <!-- TBD: BENCHMARK NUMBERS --> |

Modify mode is where it's most differentiated — surgical ZIP
rewrite vs. openpyxl's full-DOM-rebuild.

# What's in 2.0 (full surface)

- **NEW: Pivot tables** — `PivotCache`, `PivotTable`, 11
  aggregator functions, 16-chart-family `pivot_source` linkage,
  `copy_worksheet` of pivot-bearing sheets.
- 16 chart families at full openpyxl 3.1.x feature depth.
- Images, encryption (AES-256), streaming reads, rich text.
- Structural ops (insert_rows / delete_rows / insert_cols /
  delete_cols / move_range / copy_worksheet / move_sheet).
- All T1.5 modify-mode mutations (properties / defined names /
  comments / hyperlinks / tables / DV / CF).
- Reads .xlsb and .xls (via calamine).

# Still partial / not in 2.0

- Pivot-table styling beyond the named-style picker remains partial.
- In-place pivot edits in modify mode (v2.2).
- Combination charts.
- OpenDocument (.ods) — not on roadmap.

# How to verify the parity claims

Every openpyxl symbol has a `wolfxl_supported=True/False` row in
`tests/parity/openpyxl_surface.py`. The ratchet test fails red
the moment something drifts. As of v2.0 the construction-side
False rows are zero.

# Repo / docs

- https://github.com/SynthGL/wolfxl
- https://wolfxl.dev
- https://excelbench.vercel.app
- Release notes: docs/release-notes-2.0.md

Happy to answer questions about the pivot-construction
internals, the Rust crate split, or the migration path.
```

---

## dev.to long-form

**Title**: `Shipping pre-aggregated pivot-table construction in Python`

**Tags**: `python`, `rust`, `excel`, `pivot`, `engineering`

**Outline**:

1. **The problem**: openpyxl preserves pivot tables on round-trip
   but doesn't construct them. XlsxWriter doesn't support pivots.
   If you want a Python pipeline that writes a pivot-bearing
   workbook, your only option (until v2.0) was: emit the source
   data with one library, open in Excel, build the pivot manually,
   save.

2. **The OOXML pivot anatomy**: three XML parts —
   `pivotCacheDefinition{N}.xml` (schema + SharedItems),
   `pivotCacheRecords{N}.xml` (denormalised data snapshot), and
   `pivotTable{N}.xml` (layout). The `Records` part is the one
   nobody else writes.

3. **Why pre-aggregate?** Without the records snapshot, Excel
   forces a "Refresh" round-trip on first open. Inside corporate
   pipelines that auto-distribute pivot reports, that round-trip
   is a non-starter. WolfXL writes the records part from scratch
   so the workbook is "viewable" out of the gate.

4. **The implementation slice**: new Rust crate `wolfxl-pivot`
   (PyO3-free, deterministic emit), Python `wolfxl.pivot.*`
   module replacing the v0.5+ `_make_stub`, patcher Phase 2.5m
   for the workbook splice (`<pivotCaches>` + rels), and a
   chart-side `pivot_source` linkage.

5. **The orchestration**: 5 parallel pods + integrator on git
   worktrees, pre-dispatch §10 contract specs (Sprint Μ-prime
   lesson #12), sequential merge α → β → γ → δ → ε.

6. **What worked**:
   - Pre-dispatch contract specs cut integrator drift to <10 %
     of LOC vs the pre-spec ~30 %.
   - Doc pod scaffolds with `<!-- TBD: SHA -->` markers (Sprint
     Δ lesson #3) — integrator finalize fills them after γ / δ
     merge.
   - Rust `wolfxl-pivot` crate is PyO3-free; PyO3 lives in the
     `_rust` boundary crate that consumes the §10 dict shapes.

7. **What didn't**:
   - The `pivotCacheRecords` shape has a "shared vs inline"
     value mode that openpyxl reads but doesn't write. We had to
     reverse-engineer it from a known-working Excel-emitted
     fixture (RFC-047 §10.9).
   - The pivot-axis layout pre-computer
     (`python/wolfxl/pivot/_table.py`) is the heaviest pure-
     Python piece in v2.0. Pod-β's instinct was to push it into
     Rust, but the §10 contract is cleaner with a Python-side
     pre-compute that emits the dict shape Rust serialises
     verbatim.

8. **The result**: WolfXL 2.0 ships pivot construction. Speedup
   claims stay withheld until the release benchmark refresh is run
   against the final artifacts. The "full openpyxl replacement"
   phrase remains gated on the final truth pass.

9. **What's next**: finish the production-readiness audit, manual
   advanced-pivot visual checks, and benchmark replacement before any
   publish. v2.2 — in-place pivot edits in modify mode.

**Action items for the writer**:

- Pull in 2-3 Mermaid diagrams from
  `Plans/sprint-nu.md` (the OOXML pivot-anatomy diagram is
  ready-made).
- Embed a screen recording of the 6-line pivot snippet
  producing a workbook that opens in Excel with data populated.
- Link every claim to an RFC in `Plans/rfcs/047-049,054`.

---

## GitHub Discussions — Announcement post

**Title**: `WolfXL 2.0.0 — openpyxl-shaped Excel automation, pivot tables included`

**Body**:

```
👋 hi all — WolfXL 2.0.0 is shipping today.

This release closes the tracked construction-side parity roadmap,
**including pivot tables, pivot caches, and pivot-chart linkage**.
The final public "replacement" wording remains gated on the launch
truth pass.

# Headline

WolfXL constructs pivot tables with pre-aggregated
`pivotCacheRecords`. Pivots open in Excel / LibreOffice /
openpyxl with data populated, no refresh-on-open required.
(openpyxl preserves on round-trip but doesn't construct;
XlsxWriter doesn't support pivots at all.)

# Pivot construction

```python
import wolfxl
from wolfxl.chart import Reference
from wolfxl.pivot import PivotCache, PivotTable

wb = wolfxl.load_workbook("source-data.xlsx", modify=True)
ws = wb.active
src = Reference(ws, min_col=1, min_row=1, max_col=4, max_row=100)
cache = wb.add_pivot_cache(PivotCache(source=src))
pt = PivotTable(
    cache=cache, location="F2",
    rows=["region"], cols=["quarter"], data=[("revenue", "sum")],
)
ws.add_pivot_table(pt)
wb.save("pivot.xlsx")
```

Link a chart:

```python
from wolfxl.chart import BarChart

chart = BarChart()
chart.pivot_source = pt
ws.add_chart(chart, "F18")
```

# What's in 2.0

- **NEW: Pivot tables** — `wolfxl.pivot.PivotCache` /
  `PivotTable` / `RowField` / `ColumnField` / `DataField` /
  `PageField` / `PivotItem` real classes (replaces the v0.5+
  `_make_stub`); 11 aggregator functions; pre-aggregated
  records emit.
- **NEW: Pivot-chart linkage** — `chart.pivot_source = pt` on
  every one of the 16 chart families; emits `<c:pivotSource>` +
  per-series `<c:fmtId>` per ECMA-376 §21.2.2.158.
- **NEW: Deep-clone of pivot-bearing sheets** — RFC-035 limit
  lifted; `copy_worksheet` of a sheet with a pivot table now
  works.
- **REMOVED: RFC-046 §13 legacy chart-dict keys** —
  `fill_color` / `line_color` / `line_dash` / `line_width_emu`
  shim removed (deprecated in v1.7, removed v2.0).
- **README rewrite**: openpyxl-shaped compatibility, pivot coverage,
  and benchmark claims gated on audited release numbers.
- **Migration guide + Compatibility Matrix v2.0**.
- **`Plans/launch-posts.md` finalized for v2.0**.
- **CHANGELOG.md v2.0.0 entry** prepended.

# Sprint Ν deliverables

This sprint was the v2.0 launch slice. Five parallel pods +
integrator on git worktrees:

- Pod-α: Rust `wolfxl-pivot` crate (model + deterministic emit).
- Pod-β: Python `wolfxl.pivot.*` module (replaces `_make_stub`;
  layout pre-compute).
- Pod-γ: Patcher Phase 2.5m + `Workbook.add_pivot_cache` /
  `Worksheet.add_pivot_table` public APIs + RFC-035 deep-clone
  extension.
- Pod-δ: `chart.pivot_source = pt` on all 16 chart families.
- Pod-ε (this slice): docs + CHANGELOG + release notes + README
  rewrite + migration guide + Compatibility Matrix v2.0 +
  KNOWN_GAPS close-out + launch posts.

Sequential merge α → β → γ → δ → ε. Pod-ε scaffolded with
`<!-- TBD: SHA -->` markers (Sprint Δ lesson #3); integrator
finalize commit filled them after γ / δ merge.

# Migration

`pip install --upgrade wolfxl` → `wolfxl.__version__ == "2.0.0"`.

Most code is a one-line import swap. The new pivot APIs live at
`wolfxl.pivot.*`. See
[`docs/migration/openpyxl-migration.md`](../docs/migration/openpyxl-migration.md)
"Pivot tables (Sprint Ν / v2.0)" for the full mapping.

# Roadmap

- **Pre-release audit** — manual advanced-pivot/slicer visual
  verification, benchmark replacement, and final truth pass before
  any publish.
- **v2.1** — broader pivot styling beyond named-style picker.
- **v2.2** — in-place pivot edits in modify mode (source-range
  edit, field re-order, subtotal toggle).
- **v2.x** — combination charts.

# Thank you

To everyone who file-bugged the v1.6 chart contract gap, the
streaming-datetime drift, the RFC-035 cross-RFC composition
bugs, the v1.7 launch-day surprises, and the 100 little things
that surface at scale.

The construction-side parity roadmap is now close to exhausted,
subject to the final manual audit and benchmark replacement.
```

---

## Channel mapping

| Channel        | Tone        | Length    | Status |
|----------------|-------------|-----------|--------|
| HN "Show HN"   | Technical   | Medium    | Draft — blocked on benchmarks, artifact smoke, and claim verification |
| Twitter / X    | Direct, hooky | 8 tweets | Draft — blocked on benchmarks, artifact smoke, and claim verification |
| r/Python       | Detailed, comparative | Medium-long | Draft — blocked on benchmarks, artifact smoke, and claim verification |
| dev.to         | Long-form, narrative | Long | Outline — focuses on the pivot-engineering story |
| GH Discussions | Announcement, supportive | Medium | Draft — changelog mirror needs final truth pass |

## Pre-launch checklist

- [ ] All draft channel posts above polished + reviewed for tone.
- [ ] `pyproject.toml` and `Cargo.toml` reflect `2.0.0` (integrator finalize).
- [ ] `wolfxl.__version__ == "2.0.0"` verified post-bump.
- [ ] `cargo test --workspace` GREEN.
- [ ] `uv run pytest` GREEN.
- [ ] `pytest tests/parity/` GREEN, no strict-xfail markers.
- [ ] `tests/parity/KNOWN_GAPS.md` "Out of scope" matches the
      post-PR #23 audit.
- [ ] `tests/parity/openpyxl_surface.py` `wolfxl.pivot.PivotTable` flipped to `wolfxl_supported=True`.
- [ ] `docs/release-notes-2.0.md` `<!-- TBD: SHA -->` and `<!-- TBD: BENCHMARK NUMBERS -->` markers all filled by integrator.
- [ ] `Plans/launch-posts.md` (this file) `<!-- TBD -->` markers all filled by integrator.
- [ ] `CHANGELOG.md` v2.0.0 entry prepended (replaces the v2.0.0-dev WIP).
- [ ] `mkdocs build` succeeds with no warnings.
- [ ] **PyPI release** — `maturin publish --release --strip` succeeded for x86_64-linux, aarch64-linux, x86_64-darwin, aarch64-darwin, x86_64-windows.
- [ ] **PyPI verified install on all 5 wheel targets** — fresh venv per target; `pip install wolfxl==2.0.0`; smoke test with the 6-line pivot snippet.
- [ ] **Doc site live** — `wolfxl.dev` (or marketing domain) reflects v2.0 docs; v2.0 release notes accessible.
- [ ] **Benchmark dashboard live** — `WOLFXL_TEST_EPOCH=0 python scripts/bench-all.py --include-pivot --output benchmark-results-v2.0.json`; ExcelBench reflects the v2.0 numbers; pivot-construction microbenchmark visible.
- [ ] `git tag v2.0.0` cut and pushed.
- [ ] Twitter scheduling: HN post first, then Twitter thread 30 min later (avoid race for click-throughs).
- [ ] Discord / Slack notification in #release channel.
- [ ] Author blog post + dev.to cross-post 24 hrs after HN.

## Post-launch monitoring

- **HN front page ranking** — re-share at +12 hrs on Twitter if it's
  still on the front page.
- **GitHub issue triage** — first 48 hrs likely to surface
  10-20 new bug reports (the pivot surface adds new code paths
  that haven't seen production use). Triage same day; tag
  pivot-related issues with `area:pivot` for fast routing.
- **r/Python comments** — answer every top-level question within
  2 hours during NA business hours.
- **HN comments** — answer the top 5 questions within the first
  hour; subsequent within the day.
- **ExcelBench dashboard regression alerts** — monitor for any
  drift in read / write / modify medians vs the v2.0 baseline,
  plus the new pivot-construction microbenchmark.
- **PyPI download counters** — monitor v2.0.0 vs v1.7.0 split for
  the first week (target: > 50 % of new downloads on v2.0 by day
  7).
- **openpyxl-interop bug reports** — the pivot read-back via
  `openpyxl.load_workbook(...).pivots[0]` is the most likely
  source of v2.0-specific bug reports. The `Plans/sprint-nu.md`
  risk register #1 calls this out; pre-pin a fix branch
  (`hotfix/2.0.x-openpyxl-interop`) ready to receive a fast
  patch release.

## Bug-fix point release plan

If a critical bug surfaces in the first 72 hrs:

- **Day 0** — triage, reproduce, file an RFC patch (RFC-055+ for
  any new contract).
- **Day 1** — write the fix on a `hotfix/2.0.1` branch off
  `v2.0.0`. Cherry-pick into `feat/native-writer` for the next
  minor.
- **Day 1-2** — merge to `main`, tag `v2.0.1`, `maturin publish`.
- **Day 2** — Twitter / Discord / GH Discussions update note +
  CHANGELOG.md `2.0.1` entry.

If the bug is non-critical:

- Bundle into a `v2.0.x` rolling point release at the next
  natural sync (typically 1-2 weeks).

A regression on the pivot-construction or pivot-chart-linkage
path is automatically critical (the launch headline depends on
those paths working). A regression on a non-pivot path is
critical only if it's in the modify-mode patcher (the perf
differentiator).
