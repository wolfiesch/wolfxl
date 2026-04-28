# wolfxl 2.0.0 audit draft — pivot tables included

_Date: 2026-04-27_

> **Audit status**: draft release collateral. Do not publish, tag, or
> post externally from this file until the v2.0 benchmark replacement,
> manual advanced-pivot/slicer checks, clean release-artifact smoke, and
> final wording truth pass are complete.

WolfXL 2.0.0 closes the last construction-side gap on the
openpyxl-parity roadmap: **pivot tables, pivot caches, and
pivot-chart linkage**. After 24 RFCs across 9 sprints
(Δ → Ν), every construction idiom that openpyxl 3.1.x supports
works with the same Python code.

The marketing goal shifts from "openpyxl parity for the
95th-percentile case" (v1.7) toward a full replacement claim, but that
public wording remains gated on the final audit.

## TL;DR

- ✅ **Pivot table construction** — `wolfxl.pivot.PivotCache` /
  `PivotTable` / `RowField` / `ColumnField` / `DataField` /
  `PageField` / `PivotItem` are real classes with full layout
  pre-computation. Replaces the v0.5+ `_make_stub`.
- ✅ **`pivotCacheRecords{N}.xml` emit from scratch.** wolfxl
  constructs pivot tables with a pre-aggregated records snapshot.
  Pivots open in Excel, LibreOffice, and openpyxl with data
  populated — without requiring an Excel-side refresh round-trip.
- ✅ **Pivot-chart linkage** — `chart.pivot_source = pt` on every
  one of the 16 chart families. Emits `<c:pivotSource>` at the
  start of `<c:chart>` and per-series `<c:fmtId val="0"/>` per
  ECMA-376 §21.2.2.158.
- ✅ **RFC-035 deep-clone of pivot-bearing sheets.** The v1.6
  "sheets with pivots raise on copy" limit is lifted; cloned
  pivots round-trip with cell-range re-pointing on the
  source-range hint, alias-sharing on the cache (one cache
  serves many tables), and fresh `pivotTable{N}` allocations.
- ✅ **`pyproject.toml` and `Cargo.toml` → `2.0.0`.**
  `wolfxl.__version__` reports `2.0.0`. PyPI classifier stays
  `Development Status :: 5 - Production/Stable` (promoted in
  v1.7).
- ✅ **README rewritten for the audit** — openpyxl-compatible Excel
  automation with pivot construction; final benchmark claims are held
  until the v2.0 ExcelBench refresh lands.
- ✅ **Migration guide + Compatibility Matrix updated.**
  "Pivot table construction" flips from ❌ to ✅. The current
  comparison identifies wolfxl as the sole Python OOXML library in
  scope with pivot construction backed by pre-aggregated records;
  public "first/only" wording remains gated on the final truth pass.
- ✅ **RFC-046 §13 legacy chart-dict key sunset.** The Rust
  parser's accept-also for `fill_color` / `line_color` /
  `line_dash` / `line_width_emu` (deprecated in v1.7) is
  removed. Only the §10.9 `solid_fill` + nested `ln`
  form is accepted.

## Three things you can now do

### 1. Construct a pivot in 6 lines

```python
import wolfxl
from wolfxl.chart import Reference
from wolfxl.pivot import PivotCache, PivotTable

wb = wolfxl.Workbook()
ws = wb.active

# Source data
ws.append(["region", "quarter", "product", "revenue"])
ws.append(["NA",     "Q1",      "Widget",  100])
ws.append(["NA",     "Q2",      "Widget",  120])
ws.append(["EU",     "Q1",      "Widget",   80])
ws.append(["EU",     "Q2",      "Widget",   95])
# ... fill source data ...

src = Reference(ws, min_col=1, min_row=1, max_col=4, max_row=100)
cache = wb.add_pivot_cache(PivotCache(source=src))
pt = PivotTable(
    cache=cache, location="F2",
    rows=["region"], cols=["quarter"], data=[("revenue", "sum")],
)
ws.add_pivot_table(pt)
wb.save("pivot.xlsx")
```

Open `pivot.xlsx` in Excel, LibreOffice, or read it with
openpyxl — the pivot's row / column / data layout is already
populated; **no refresh-on-open required**.

### 2. Link a chart to a pivot

```python
from wolfxl.chart import BarChart, Reference

chart = BarChart()
chart.title = "Revenue by region × quarter"

# Bind the chart to the pivot — Excel renders it as a true
# PivotChart (right-click → "Refresh", PivotChart toolbar):
chart.pivot_source = pt

# Or with an explicit (name, fmt_id):
# chart.pivot_source = ("MyPivot", 0)

ws.add_chart(chart, "F18")
```

`<c:pivotSource>` is emitted at the start of `<c:chart>`, with
each `<c:ser>` carrying a `<c:fmtId val="0"/>` element. Read the
result via openpyxl: `chart.pivotSource.name == "MyPivot"`.

### 3. Deep-clone a sheet that has a pivot

```python
wb = wolfxl.load_workbook("template_with_pivot.xlsx", modify=True)
src = wb["Source"]
clone = wb.copy_worksheet(src)
# clone is a fresh sheet with:
#   - its own pivotTable{N+1}.xml (fresh id)
#   - the source's pivotCache aliased (one cache serves both)
#   - the cell-range hint re-pointed to the new sheet's name
#   - all the v1.6 chart-deep-clone preservation
clone.title = "Cloned"
wb.save("template_with_pivot.xlsx")
```

The "sheets with pivot tables raise on copy" limit from RFC-035
§10 is lifted in v2.0.

## What's new

### RFC-047 — pivot caches

`wolfxl.pivot.PivotCache(source=Reference(...))` builds a typed
cache around a source range. On `wb.add_pivot_cache(cache)` the
cache walks the source range, infers per-column type (string /
number / date / boolean / mixed), builds `SharedItems` for each
field, and emits two XML parts:

- `xl/pivotCache/pivotCacheDefinition{N}.xml` — schema with
  `CacheField` per source column, `SharedItems` enumeration.
- `xl/pivotCache/pivotCacheRecords{N}.xml` — denormalised
  rectangular snapshot, one `<r>` per source-data row, with
  `<x v="N"/>` indices into SharedItems for shared values and
  inline `<n v=42/>` / `<s v="text"/>` for non-shared.

The records emit is the **Option A** differentiator: in the current
project comparison, no other Python OOXML library has been identified
that writes the records snapshot from scratch, so Excel / LibreOffice /
openpyxl all read the pivot's data without an Excel-side refresh.
(openpyxl preserves records on round-trip but doesn't construct them;
XlsxWriter doesn't support pivots at all.) Keep any public "first/only"
wording gated on the final launch truth pass.

The workbook-side splice (`<pivotCaches>` collection in
`xl/workbook.xml` + a rel of type `pivotCacheDefinition` in
`xl/_rels/workbook.xml.rels`) is handled by the patcher's
Phase 2.5m. Cache-id allocation goes through `PartIdAllocator`,
the same mechanism RFC-035 / RFC-046 already use.

See `Plans/rfcs/047-pivot-caches.md` for the full §10 contract.

### RFC-048 — pivot tables

`wolfxl.pivot.PivotTable(cache=..., location=..., rows=..., cols=..., data=...)`
holds the layout. Bare-string axis specs (`rows=["region"]`)
work for the common case; explicit `RowField` / `ColumnField` /
`DataField` / `PageField` builders are available for
fine-grained control (custom captions, custom subtotals, custom
sort orders).

11 aggregator functions are supported on `DataField`:

| Aggregator | Excel name |
|---|---|
| `sum` | Sum |
| `count` | Count |
| `average` | Average |
| `max` | Max |
| `min` | Min |
| `product` | Product |
| `count_nums` | CountNums |
| `std_dev` | StdDev |
| `std_dev_p` | StdDevp |
| `var` | Var |
| `var_p` | Varp |

The "Option A" core lives in `python/wolfxl/pivot/_table.py`:
the layout pre-computer enumerates `<rowItems>` and
`<colItems>` and aggregates per-data-field values per pivot
intersection. After save, Excel does not need to recompute —
it reads the pre-aggregated `<rowItems>` / `<colItems>`
directly.

`Worksheet.add_pivot_table(pt, anchor=...)` slots into the
patcher's Phase 2.5m alongside the cache; allocates a fresh
`pivotTable{N}.xml`; wires the sheet's rels to the cache's;
emits the table XML through `file_adds`.

See `Plans/rfcs/048-pivot-tables.md` for the full §10 contract.

### RFC-049 — pivot-chart linkage

```python
chart.pivot_source = pt              # or (name, fmt_id) tuple
```

Touches all 16 chart families uniformly — the attribute lives
on `ChartBase`, so `BarChart`, `LineChart`, `PieChart`,
`BarChart3D`, `StockChart`, `ProjectedPieChart`, etc. all gain
the same setter. Validates `name` against
`^([A-Za-z_][A-Za-z0-9_]*!)?[A-Za-z_][A-Za-z0-9_ ]*$`;
validates `fmt_id` in `[0, 65535]`.

The Rust emitter (`crates/wolfxl-writer/src/emit/charts.rs`)
inserts `<c:pivotSource>` between the `<c:chart>` open and the
`<c:plotArea>` open per ECMA-376 §21.2.2.158, plus
`<c:fmtId val="0"/>` on every `<c:ser>` (required by spec when
`pivotSource` is present).

When a pivot-bearing chart is `copy_worksheet`'d, the cloned
chart's `pivot_source.name` is rewritten if the cloned pivot
table got a renamed `displayName` (mirrors the cell-range-rewrite
pattern from v1.6).

See `Plans/rfcs/049-pivot-charts.md` for the full §10 contract.

### RFC-054 — launch hardening

This RFC has no code; it's the launch-day envelope:

- `docs/release-notes-2.0.md` (this file).
- `docs/migration/openpyxl-migration.md` — new "Pivot tables"
  section.
- `docs/migration/compatibility-matrix.md` — pivot row flips ❌
  → ✅; ecosystem comparison updated.
- `tests/parity/KNOWN_GAPS.md` — pivot row moves to "Closed in
  2.0".
- `Plans/launch-posts.md` — finalized for v2.0 with pivot
  snippets in HN / Twitter / r/Python / dev.to / GH Discussions
  drafts.
- `CHANGELOG.md` — v2.0.0 entry replaces the WIP `2.0.0-dev`.
- README rewrite — openpyxl-compatible positioning with release-claim
  caveats.

## Sprint Ν acknowledgements

| Pod | Branch | Deliverable | Merge SHA |
|---|---|---|---|
| α + β | landed inline on `feat/native-writer` | Rust `wolfxl-pivot` crate; `pivot_cache_definition_xml` / `pivot_cache_records_xml` / `pivot_table_xml` deterministic emit + Python `wolfxl.pivot.*` module with layout pre-compute and 11 aggregator functions | `38234b0` |
| γ | `feat/sprint-nu-pod-gamma` → merge `407f1b7` | Patcher Phase 2.5m (queued pivot adds, `PartIdAllocator`-backed numbering, workbook + sheet rels splice); PyO3 bindings (`serialize_pivot_cache_dict`, `serialize_pivot_records_dict`, `serialize_pivot_table_dict`); RFC-035 deep-clone extension | `ba9db64`, `658e296`, `7f4081e`, `d21edb5`, `4b5e16f` |
| δ | `feat/sprint-nu-pod-delta` → merge `ac19b60` | `chart.pivot_source = pt` on every chart family; `<c:pivotSource>` block emit + mandatory per-series `<c:fmtId>` | `2fd0de0`, `501f12b`, `2ade7b5`, `04c93bc` |
| ε | `feat/sprint-nu-pod-epsilon` → merge `231182d` | This release-notes file; README rewrite; migration guide; KNOWN_GAPS close-out; launch posts; CHANGELOG finalize | `7d3d82a`, `f2df471`, `0f099c7`, `329b9ec`, `14f0d7a`, `e5762be`, `79f5b2e` |
| Integrator | `feat/native-writer` | Pre-dispatch §10 contract specs; sequential α → β → γ → δ → ε merge; reconciliation pass; `2.0.0` version bump; new openpyxl-surface ratchet entries; `v2.0.0` tag | this commit |

Sprint Ν used the parallel-pod orchestration that landed v1.6 /
v1.6.1 / v1.7. Pre-dispatch §10 contracts (RFC-047 §10, RFC-048
§10, RFC-049 §10) were authored before any pod opened a worktree
— Sprint Μ-prime lesson #12. Pod-ε scaffolded with SHA markers
(lesson #3); the post-PR #23 audit keeps final branch evidence gated
until the release truth pass reconciles it.

## Migration notes

### From v1.7

`pip install --upgrade wolfxl` → `wolfxl.__version__ == "2.0.0"`.

If your code uses `wolfxl.pivot` for round-trip preservation
only (the v1.7 stub case), nothing changes — pivot tables still
preserve verbatim through modify mode. If you want to construct
pivots from Python, see "Three things you can now do" above.

### RFC-046 §13 — legacy chart-dict key sunset

Deprecated in v1.7, **removed in v2.0**. The Rust parser
(`src/native_writer_backend.rs::parse_chart_dict`) no longer
accepts the legacy keys:

| Legacy key (removed) | §10.9 form (kept) |
|---|---|
| `fill_color: "FF0000"` | `solid_fill: "FF0000"` |
| `line_color: "0000FF"` | `ln: {solid_fill: "0000FF", w_emu: 12700, prst_dash: "solid"}` |
| `line_dash: "dash"` | `ln: {prst_dash: "dash", ...}` |
| `line_width_emu: 12700` | `ln: {w_emu: 12700, ...}` |

The Python emitter has used the §10.9 form exclusively since
v1.6.1 — only out-of-tree callers that bypassed
`Worksheet.add_chart` and built chart dicts by hand are
affected. If you hit this on upgrade, rewrite the dict per the
§10.9 form (see `Plans/rfcs/046-chart-construction.md` §10.9
for the canonical example).

### openpyxl pivot import-path differences

openpyxl's pivot construction lives at:

```python
from openpyxl.pivot.table import TableDefinition
from openpyxl.pivot.cache import CacheDefinition
```

wolfxl's lives at:

```python
from wolfxl.pivot import PivotTable, PivotCache
```

API differences:

- openpyxl exposes a one-step `ws.add_pivot(table)`; wolfxl
  splits cache and table into two steps:
  `cache = wb.add_pivot_cache(PivotCache(source=ref))` then
  `ws.add_pivot_table(PivotTable(cache=cache, ...))`. The split
  exists because OOXML caches are workbook-scoped (one cache can
  serve multiple tables) while tables are sheet-scoped.
- openpyxl's `PivotField` / `DataField` / etc. live under
  `openpyxl.pivot.table`; wolfxl's live at `wolfxl.pivot`.
- openpyxl's `Reference` for the pivot source is the same shape
  as the chart `Reference`; wolfxl re-uses
  `wolfxl.chart.Reference` for both, mirroring openpyxl 3.1.x's
  shared reference type.

See `docs/migration/openpyxl-migration.md` "Pivot tables (Sprint
Ν / v2.0)" for the full mapping.

## Out of scope / partial after post-PR #23 audit

- **Slicers** (`xl/slicers/` + `xl/slicerCaches/`) are now
  implemented and covered by zip-integrity, copy-worksheet, and
  openpyxl-load smoke tests. Full manual Excel/LibreOffice visual
  verification remains a release-readiness task.
- **Calculated fields** (`<calculatedField>`), **calculated items**
  (`<calculatedItem>`), and **GroupItems** (`<fieldGroup>`) are now
  implemented and covered by advanced pivot tests.
- **OLAP / external pivot caches**. Needs the PowerPivot
  data-model (`xl/model/`). Out of scope permanently.
- **Pivot-table styling beyond the named-style picker**.
  PivotArea formats and pivot-scoped conditional formatting are
  implemented; broader theme/banded styling is still partial.
- **In-place pivot edits in modify mode** beyond
  `add_pivot_table`. Editing an existing pivot's source range,
  field ordering, etc. v2.2.
- **Combination charts** (multi-plot charts on shared axes) remain
  post-v2.0.

## Verification

| Surface | Status | Tooling |
|---|---|---|
| Rust unit tests | full workspace green in the post-PR #23 audit | `cargo test --workspace` |
| Python unit tests | 2278 passed, 29 skipped in the post-chart/LibreOffice truth pass | `uv run --no-sync pytest -q` |
| openpyxl-parity ratchet | 445 passed, 4 skipped in the post-PR #23 parity run | `uv run pytest tests/parity -q -x` |
| LibreOffice cross-renderer | 47 opt-in smoke tests passed, including copy_worksheet, array formulas, and pivot-chart render smoke | `WOLFXL_RUN_LIBREOFFICE_SMOKE=1 uv run --no-sync pytest ...` |
| openpyxl interop | Advanced pivot fixtures save cleanly and can be opened by `openpyxl.load_workbook(...)` | `tests/parity/test_advanced_pivots_parity.py` |
| Excel-on-Windows | Manual smoke test on each pivot-fixture file | Excel 365 (latest) and Excel 2021 |
| Benchmark dashboard | v2.0 numbers refreshed | `WOLFXL_TEST_EPOCH=0 python scripts/bench-all.py --include-pivot --output benchmark-results-v2.0.json` |

### Benchmark headline (v2.0)

Benchmark numbers are intentionally withheld until the v2.0 ExcelBench
refresh is rerun and checked against the release artifact. The release
benchmark grid should cover read, write, modify/touch-one-cell,
`copy_worksheet`, and pivot construction workloads before any public
speedup claim is restored.

## Stats (post-2.0.0)

- `python -c "import wolfxl; print(wolfxl.__version__)"` →
  `2.0.0`.
- `wolfxl.pivot.PivotTable.__module__` →
  `wolfxl.pivot._table` (no longer `wolfxl._compat` stub).
- `tests/parity/openpyxl_surface.py` — pivot row flipped:
  `wolfxl_supported=True`.
- `tests/parity/KNOWN_GAPS.md` "Out of scope" section reduced
  to: OLAP/external pivots, partial pivot styling, in-place pivot
  edits, combination charts, chart removal for source-surviving
  charts, and non-xlsx write/style-accessor limits.

## Pods that landed 2.0.0

- **Pod-α** (`crates/wolfxl-pivot/`) — Rust crate with model +
  emit; PyO3-free; 25+ unit tests covering deterministic emit
  (`WOLFXL_TEST_EPOCH=0` golden tests).
- **Pod-β** (`python/wolfxl/pivot/`) — Python module replacing
  the `_make_stub`; per-class `to_rust_dict()` + layout
  pre-computer; construction-time validation; 40+ tests in
  `tests/test_pivot_construction.py`.
- **Pod-γ** — Patcher Phase 2.5m; `Worksheet.add_pivot_table` /
  `Workbook.add_pivot_cache` public APIs; RFC-035 deep-clone
  extension for pivot-bearing sheets; PyO3 bindings.
- **Pod-δ** — `chart.pivot_source = pt` on all 16 chart
  families; `<c:pivotSource>` Rust emit; per-series `<c:fmtId>`.
- **Pod-ε** (this slice) — docs, CHANGELOG finalize,
  release-notes-2.0, README rewrite, Compatibility Matrix v2.0,
  KNOWN_GAPS close-out, launch posts.

## RFCs

- `Plans/rfcs/047-pivot-caches.md` — pivot caches.
- `Plans/rfcs/048-pivot-tables.md` — pivot tables.
- `Plans/rfcs/049-pivot-charts.md` — pivot-chart linkage.
- `Plans/rfcs/054-launch-hardening.md` — launch hardening.
- `Plans/sprint-nu.md` — sprint plan.

## Next

**Next audit focus** — finish manual Excel/LibreOffice visual
checks for advanced pivots/slicers, run benchmark replacement for
the remaining TBD numbers, and decide whether partial pivot styling
is acceptable for the release claim. No publish/release step should
run until those gates are complete.
