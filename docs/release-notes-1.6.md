# wolfxl 1.6.0 (TBD-DATE) — chart construction (8 types, full depth)

_Date: <!-- TBD -->_

WolfXL 1.6 lifts the construction-side chart stub that the 1.0–1.5
arc deferred as out-of-scope. Sprint Μ ("Mu") ships **eight 2D chart
types** at full openpyxl 3.1.x depth — `BarChart`, `LineChart`,
`PieChart`, `DoughnutChart`, `AreaChart`, `ScatterChart`,
`BubbleChart`, `RadarChart` — plus `Reference`, `Series`, and
`Worksheet.add_chart(chart, anchor)`. The 3D / Stock / Surface /
ProjectedPieChart variants remain stubbed with a v1.6.1 pointer; the
pivot-chart linkage depends on Sprint Ν / v2.0.0 pivot tables.

Sprint Μ also lifts the RFC-035 `copy_worksheet` chart-aliasing
limit: charts in copied sheets now deep-clone with cell-range
re-pointing instead of aliasing the source's chart part. Modify-mode
`add_chart` ships alongside the write-mode path so both flows have
identical surface and behaviour.

## TL;DR

- **Chart construction** — `wolfxl.chart.{Bar,Line,Pie,Doughnut,Area,Scatter,Bubble,Radar}Chart`
  are real classes. `Reference`, `Series`, and `Worksheet.add_chart(chart, anchor)`
  ship together. Full per-type feature depth (gap_width, smooth,
  vary_colors, hole_size, scatter_style, bubble_3d, radar_style, …).
  Both write mode and modify mode supported.
- **`copy_worksheet` chart deep-clone** — RFC-035 §10's chart-aliasing
  limit (lines 924-929) is lifted. Charts in copied sheets now
  deep-clone with cell-range re-pointing. Self-references rewrite to
  the copy's sheet name; cross-sheet references preserved.
- **Modify-mode `add_chart`** — `wb = load_workbook(..., modify=True);
  ws.add_chart(chart, "G2"); wb.save()` works. New patcher Phase 2.5l
  drains queued chart adds and routes them through `file_adds`.

## What's new

### Chart construction (Sprint Μ Pod-α/β, RFC-046)

```python
from wolfxl.chart import BarChart, LineChart, PieChart, Reference
import wolfxl

wb = wolfxl.Workbook()
ws = wb.active
ws.append(["Region", "Q1", "Q2", "Q3", "Q4"])
ws.append(["NA",      100,  120,  90,   140])
ws.append(["EU",      80,   95,   110,  85])
ws.append(["APAC",    60,   70,   85,   100])

chart = BarChart()
chart.type = "col"
chart.style = 10
chart.title = "Quarterly Revenue"
chart.y_axis.title = "Revenue (USD)"
chart.x_axis.title = "Quarter"

data = Reference(ws, min_col=2, min_row=1, max_col=5, max_row=4)
cats = Reference(ws, min_col=1, min_row=2, max_row=4)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

ws.add_chart(chart, "G2")
wb.save("revenue.xlsx")
```

Sample code per chart type (each verified to work verbatim with
wolfxl):

#### LineChart

```python
from wolfxl.chart import LineChart, Reference
import wolfxl

wb = wolfxl.Workbook()
ws = wb.active
for row in [(1, 10, 30), (2, 40, 60), (3, 50, 70), (4, 20, 10), (5, 10, 40)]:
    ws.append(row)

chart = LineChart()
chart.title = "Line Chart"
chart.style = 13
chart.x_axis.title = "Time (s)"
chart.y_axis.title = "Voltage"
data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=5)
chart.add_data(data, titles_from_data=False)

# Per-series feature: smooth lines
for series in chart.series:
    series.smooth = True

ws.add_chart(chart, "E2")
wb.save("line.xlsx")
```

#### PieChart

```python
from wolfxl.chart import PieChart, Reference
import wolfxl

wb = wolfxl.Workbook()
ws = wb.active
data = [["Pie", "Slice"], ["Apple", 50], ["Cherry", 30], ["Pumpkin", 10], ["Chocolate", 40]]
for row in data:
    ws.append(row)

chart = PieChart()
labels = Reference(ws, min_col=1, min_row=2, max_row=5)
data_ref = Reference(ws, min_col=2, min_row=1, max_row=5)
chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(labels)
chart.title = "Pies sold by category"

ws.add_chart(chart, "D1")
wb.save("pie.xlsx")
```

#### DoughnutChart

```python
from wolfxl.chart import DoughnutChart, Reference
import wolfxl

wb = wolfxl.Workbook()
ws = wb.active
ws.append(["Slice", "Value"])
ws.append(["Apple", 50])
ws.append(["Cherry", 30])
ws.append(["Pumpkin", 10])

chart = DoughnutChart()
chart.title = "Doughnut Chart"
chart.hole_size = 50         # 50% hole
labels = Reference(ws, min_col=1, min_row=2, max_row=4)
data_ref = Reference(ws, min_col=2, min_row=1, max_row=4)
chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(labels)

ws.add_chart(chart, "D1")
wb.save("doughnut.xlsx")
```

#### AreaChart

```python
from wolfxl.chart import AreaChart, Reference
import wolfxl

wb = wolfxl.Workbook()
ws = wb.active
for row in [["Yr", "A", "B"], [2020, 30, 40], [2021, 50, 35], [2022, 70, 60], [2023, 65, 75]]:
    ws.append(row)

chart = AreaChart()
chart.title = "Area Chart"
chart.style = 13
chart.grouping = "stacked"
chart.x_axis.title = "Year"
chart.y_axis.title = "Units"
cats = Reference(ws, min_col=1, min_row=2, max_row=5)
data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=5)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

ws.add_chart(chart, "E2")
wb.save("area.xlsx")
```

#### ScatterChart

```python
from wolfxl.chart import ScatterChart, Reference, Series
import wolfxl

wb = wolfxl.Workbook()
ws = wb.active
ws.append(["X", "Y1", "Y2"])
for x in range(1, 11):
    ws.append([x, x ** 2, x ** 0.5])

chart = ScatterChart()
chart.title = "Scatter Chart"
chart.style = 13
chart.x_axis.title = "X"
chart.y_axis.title = "f(X)"
chart.scatter_style = "smoothMarker"

xvals = Reference(ws, min_col=1, min_row=2, max_row=11)
for col in (2, 3):
    yvals = Reference(ws, min_col=col, min_row=1, max_row=11)
    series = Series(yvals, xvals, title_from_data=True)
    chart.series.append(series)

ws.add_chart(chart, "E2")
wb.save("scatter.xlsx")
```

#### BubbleChart

```python
from wolfxl.chart import BubbleChart, Reference, Series
import wolfxl

wb = wolfxl.Workbook()
ws = wb.active
ws.append(("Number of Products", "Sales in USD", "Market share"))
ws.append((14,    12200,   15))
ws.append((20,    60000,   33))
ws.append((18,    24400,   10))
ws.append((22,    32000,   42))

chart = BubbleChart()
chart.title = "Bubble Chart"
chart.style = 18
chart.bubble_3d = False
chart.bubble_scale = 100

xvals = Reference(ws, min_col=1, min_row=2, max_row=5)
yvals = Reference(ws, min_col=2, min_row=2, max_row=5)
sizes = Reference(ws, min_col=3, min_row=2, max_row=5)
series = Series(yvals, xvals, zvalues=sizes, title="2024")
chart.series.append(series)

ws.add_chart(chart, "E2")
wb.save("bubble.xlsx")
```

#### RadarChart

```python
from wolfxl.chart import RadarChart, Reference
import wolfxl

wb = wolfxl.Workbook()
ws = wb.active
rows = [
    ("Quality", "Style", "Performance", "Reliability", "Comfort"),
    (5, 4, 3, 4, 5),
    (4, 5, 5, 3, 4),
]
for row in rows:
    ws.append(row)

chart = RadarChart()
chart.type = "filled"            # 'standard' | 'marker' | 'filled'
chart.style = 26
chart.title = "Radar Chart"
chart.y_axis.delete = True

cats = Reference(ws, min_col=1, min_row=1, max_col=5, max_row=1)
data = Reference(ws, min_col=1, min_row=2, max_col=5, max_row=3)
chart.add_data(data, titles_from_data=False)
chart.set_categories(cats)

ws.add_chart(chart, "A8")
wb.save("radar.xlsx")
```

### `copy_worksheet` chart deep-clone (Sprint Μ Pod-γ, RFC-035 §10 lift)

```python
import wolfxl
from wolfxl.chart import BarChart, Reference

wb = wolfxl.Workbook()
ws = wb.active
ws.title = "Source"
ws.append(["Region", "Sales"])
for row in [("NA", 100), ("EU", 80), ("APAC", 60)]:
    ws.append(row)

chart = BarChart()
data = Reference(ws, min_col=2, min_row=1, max_row=4)
cats = Reference(ws, min_col=1, min_row=2, max_row=4)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "D1")

# Copy the sheet — chart's data refs re-point to the copy
copy = wb.copy_worksheet(ws)
copy.title = "SourceCopy"

# Before 1.6: copy's chart still pointed at Source!$B$1:$B$4 (alias)
# In 1.6:   copy's chart now points at SourceCopy!$B$1:$B$4 (deep-clone + re-point)
wb.save("with_copies.xlsx")
```

The deep-clone re-points self-references (LHS of `!` matches the
source sheet name) to the copy's sheet name. Cross-sheet references
(LHS of `!` is some other sheet) are preserved verbatim — copying
`Sheet2` whose chart pulls data from `Sheet1` keeps the
`Sheet1`-pointing references intact, since `Sheet1` itself was not
copied.

Cached values inside `<c:strCache>` / `<c:numCache>` are preserved
as-is. Excel rebuilds them on next open if they go stale; explicit
chart-cache rebuild is left to the user (matches the `xl/calcChain.xml`
deferred-rebuild contract from RFC-035).

### Modify-mode `add_chart` (Sprint Μ Pod-γ)

```python
import wolfxl
from wolfxl.chart import BarChart, Reference

# Open an existing workbook in modify mode
wb = wolfxl.load_workbook("template.xlsx", modify=True)
ws = wb["Data"]

# Build a chart against the existing sheet's data
chart = BarChart()
chart.title = "Sales (auto-added)"
data = Reference(ws, min_col=2, min_row=1, max_col=4, max_row=20)
cats = Reference(ws, min_col=1, min_row=2, max_row=20)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "G2")

wb.save("template.xlsx")           # works in modify mode
```

XlsxPatcher Phase 2.5l drains the queued chart adds per sheet,
allocates fresh `chartN.xml` and `drawingN.xml` numbers via
`PartIdAllocator` (collision-free against the source ZIP listing
AND any in-flight `file_adds`), emits chart bytes through
`file_adds`, and splices `<xdr:graphicFrame>` blocks into the
sheet's existing drawing part (or creates one if the sheet had no
drawing yet). Composes cleanly with RFC-045 `add_image`: a single
drawing part can carry image `<xdr:pic>` blocks and chart
`<xdr:graphicFrame>` blocks side-by-side.

## Migration notes

### Charts (RFC-046)

* **`BarChart()` etc. previously raised `NotImplementedError`**.
  Code that defensively caught the exception (e.g. wrapped chart
  constructions in `try/except` to fall through to a "render as a
  table instead" branch) now succeeds. Audit any `try`/`except`
  around chart construction or `ws.add_chart(...)` calls and remove
  the fallback branch.
* **`Reference()` and `Series()` previously raised
  `NotImplementedError`** too. Same migration path.
* **Drop-in for openpyxl users**: every keyword arg matches openpyxl
  3.1.x verbatim. The search-replace from `from openpyxl.chart
  import …` to `from wolfxl.chart import …` is the only change
  required for the eight 2D chart families.
* **3D / Stock / Surface / ProjectedPie still raise
  `NotImplementedError`**. The error message points at v1.6.1.
  Workloads that need 3D charts in 1.6.0 should either:
  (1) construct the 2D variant (`BarChart` instead of `BarChart3D`),
  (2) round-trip a workbook that already has a 3D chart through
      modify mode (the existing chart bytes are preserved verbatim),
  or (3) wait for v1.6.1.

### `copy_worksheet` chart behaviour change (RFC-035 §10 lift)

* **Chart-aliasing in copied sheets is replaced by deep-clone with
  cell-range re-pointing**. Code that relied on the alias (the copy's
  chart pointing at the source's data) now sees the copy's chart
  pointing at the copy's data. This is the openpyxl-parity contract
  and matches what the Excel UI does when you right-click "Move or
  Copy".
* **Migration**: if you specifically want the alias behaviour (chart
  on the copy that still references the source's cells), construct
  a new chart via `BarChart()` after the copy and reference the
  source sheet's cells explicitly:
  ```python
  copy = wb.copy_worksheet(ws)
  # Replace the deep-cloned chart with a fresh chart pointing at the source
  for chart in list(copy._charts):    # private list, public list pending
      copy._charts.remove(chart)
  alias_chart = BarChart()
  alias_chart.add_data(Reference(ws, …), titles_from_data=True)  # ws, not copy
  copy.add_chart(alias_chart, "D1")
  ```
* **Cross-sheet chart references are preserved**, so a chart on
  `Sheet2` that references `Sheet1`'s data keeps the
  `Sheet1`-pointing references after copying `Sheet2`. This is
  almost always what you want; the only time the change is visible
  is for self-referencing charts.

### Modify-mode `add_chart`

* **Code that `try/except`'d `ws.add_chart` in modify mode** to fall
  back to a "regenerate the whole workbook in write mode" branch
  now succeeds. Drop the fallback.
* **No `remove_chart` or `replace_chart` API yet** — v1.6.0 is
  additive only. Existing charts on round-tripped workbooks are
  still preserved verbatim.

## Out of scope (documented, planned)

The 1.6 release closes "chart construction" for the eight 2D
families. Three large items remain on the roadmap:

* **Chart 3D / Stock / Surface / ProjectedPieChart** — scheduled for
  **v1.6.1**. Stub classes raise `NotImplementedError` in 1.6.0.
  RFC-046 §9 documents the deferral and the per-family XML element
  surface that ships in 1.6.1.
* **Pivot tables + pivot charts** — scheduled for **v2.0.0** (Sprint
  Ν). Pivot caches and pivot tables are preserved on round-trip but
  cannot be added programmatically. v2.0.0 is the public-launch
  milestone; pivots ship alongside the launch.
* **Public launch** — **v2.0.0**.

Other out-of-scope items (chart trendlines, error bars, data tables
under chart, conditional series formatting, replace / delete
existing charts) are tracked in `tests/parity/KNOWN_GAPS.md` and
RFC-046 §9 with deferral rationales.

## RFCs

- `Plans/rfcs/046-chart-construction.md` (Sprint Μ Pod-α/β/γ/δ; docs by Pod-ε) — <!-- TBD: SHA -->

## Stats (post-1.6)

- `cargo test --workspace --exclude wolfxl`: ~660 + N green
  (Pod-α adds per-type chart-emit tests, axis-id-allocation tests).
- `pytest tests/`: **~1235+ → ~1300+ passed** (Pod-β adds ~40
  write-mode tests, Pod-γ adds ~25 modify-mode + chart-deep-clone
  tests; the exact count is filled in on integrator merge).
- `pytest tests/parity/`: **~165+ → ~190+ passed** (Pod-δ adds ~25
  parity tests).
- `KNOWN_GAPS.md` "Out of scope" pruned: chart construction lifted
  to "Closed in 1.6"; chart 3D / Stock / Surface / ProjectedPie
  explicitly scheduled for v1.6.1; pivot-chart linkage explicitly
  deferred to Sprint Ν / v2.0.0.

## Acknowledgments

Sprint Μ ("Mu") pods that landed 1.6:

- **Pod-α — RFC-046 Rust chart XML emit + types.** Files:
  `crates/wolfxl-writer/src/model/chart.rs`,
  `crates/wolfxl-writer/src/emit/charts.rs`,
  `crates/wolfxl-writer/src/emit/drawings.rs` (extended for
  `<xdr:graphicFrame>`),
  `crates/wolfxl-rels/src/lib.rs` (`RT_CHART`),
  `src/lib.rs` (PyO3 binding `Workbook.add_chart_native`). Commits:
  <!-- TBD: SHA -->. Merged via <!-- TBD: SHA -->.
- **Pod-β — RFC-046 Python class hierarchy.** ~17 modules under
  `python/wolfxl/chart/`. Replaces all `_make_stub` chart classes
  with real ones; mirrors openpyxl's API verbatim. Commits:
  <!-- TBD: SHA -->. Merged via <!-- TBD: SHA -->.
- **Pod-γ — Modify-mode `add_chart` + RFC-035 chart-deep-clone.**
  XlsxPatcher Phase 2.5l; cell-range re-pointing on the chart-clone
  pass. Commits: <!-- TBD: SHA -->. Merged via <!-- TBD: SHA -->.
- **Pod-δ — Parity tests vs openpyxl.** ~40 write-mode tests +
  ~25 parity tests + 8 LibreOffice smoke + 10 surface entries.
  Commits: <!-- TBD: SHA -->. Merged via <!-- TBD: SHA -->.
- **Pod-ε (this release scaffold)** — RFC-046, INDEX update,
  KNOWN_GAPS reconciliation, this release notes scaffold, CHANGELOG
  entry, and the RFC-035 §10 chart-aliasing-limit lifted note.
  Commits: <!-- TBD: SHA -->. Merged via <!-- TBD: SHA -->.

## SHA log

| Pod | Branch | Commits | Merge |
|---|---|---|---|
| α | `feat/sprint-mu-pod-alpha`   | <!-- TBD: SHA --> | <!-- TBD: SHA --> |
| β | `feat/sprint-mu-pod-beta`    | <!-- TBD: SHA --> | <!-- TBD: SHA --> |
| γ | `feat/sprint-mu-pod-gamma`   | <!-- TBD: SHA --> | <!-- TBD: SHA --> |
| δ | `feat/sprint-mu-pod-delta`   | <!-- TBD: SHA --> | <!-- TBD: SHA --> |
| ε | `feat/sprint-mu-pod-epsilon` | <!-- TBD: SHA --> | <!-- TBD: SHA --> |

Integrator finalize commit fills these placeholders, performs the
post-merge ratchet flip on the 10 chart-related entries in
`tests/parity/openpyxl_surface.py` (`shipped-1.6` tag), and tags
`v1.6.0`.

After Sprint Μ the openpyxl-parity surface is exhausted at the read
level (1.0 → 1.4) AND at the construction level for encryption /
images / charts (1.5 → 1.6). The remaining roadmap is the v1.6.1
follow-up (3D / Stock / Surface / ProjectedPie chart families) and
the v2.0.0 public-launch milestone (pivot tables + pivot charts +
launch). Thanks to everyone who file-bugged the chart construction
stubs over the 1.0 → 1.5 cycle — every workload that hit
`NotImplementedError("Charts are preserved on modify-mode round-trip
but cannot be added programmatically.")` drove this slice.
