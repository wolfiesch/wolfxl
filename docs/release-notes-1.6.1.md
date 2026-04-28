# wolfxl 1.6.1 (TBD-DATE) — chart contract reconciliation + 3D variants

_Date: <!-- TBD -->_

WolfXL 1.6.1 closes the two debts that v1.6.0 deferred:

1. The **Pod-α/Pod-β chart-dict contract gap**: 37 advanced
   sub-feature tests in `tests/test_charts_write.py` were marked
   `xfail` in 1.6.0 because the Rust parser
   (`src/native_writer_backend.rs::parse_chart_dict`) and the
   Python emitter (per-class `to_rust_dict`) diverged on title
   runs, manual layout xy, gridlines, error bars, trendlines,
   `varyColors`, non-default grouping (`stacked`/`percentStacked`),
   scatter style, marker symbol, fill color, and invalid-input
   rejection. RFC-046 §10 promotes the chart-dict shape to
   authoritative; Pod-α′ and Pod-β′ both code to §10 verbatim. All
   37 xfailed tests now pass.

2. The **8 deferred chart families**: `BarChart3D`, `LineChart3D`,
   `PieChart3D` (alias `Pie3D`), `AreaChart3D`, `SurfaceChart`,
   `SurfaceChart3D`, `StockChart`, `ProjectedPieChart` ship as real
   classes, replacing the v1.6.0 `NotImplementedError` stubs.

Plus: modify-mode `Worksheet.add_chart(BarChart(...))` no longer
drops with `RuntimeWarning` — Pod-γ′ wires the high-level bridge
through Pod-α′'s new `serialize_chart_dict` PyO3 export.

## TL;DR

- **37 chart-write xfailed tests → pass.** The module-level
  `pytest.mark.xfail` on `tests/test_charts_write.py` is removed
  (Pod-δ′); every advanced sub-feature now passes on the green
  path.
- **8 new chart families** at full openpyxl 3.1.x depth: 3D
  variants (`BarChart3D`, `LineChart3D`, `PieChart3D` / `Pie3D`,
  `AreaChart3D`, `SurfaceChart3D`), 2D `SurfaceChart`, `StockChart`
  (OHLC), `ProjectedPieChart` (pie-of-pie / bar-of-pie). 9 new
  parity-ratchet entries in `tests/parity/openpyxl_surface.py`.
- **Modify-mode high-level `add_chart` works end-to-end.** The
  v1.6.0 warn-and-drop fallback in
  `Workbook._flush_pending_charts_to_patcher` is replaced with a
  real dict→bytes bridge via the new `serialize_chart_dict` PyO3
  helper.
- **Construction-time validation.** Empty series, bad anchor (not
  A1 and not a recognised anchor object), out-of-range `Reference`
  bounds, out-of-range `style` / `gap_width` / `overlap` /
  `hole_size` / `bubble_scale` / poly trendline `order`, and
  `display_blanks_as` outside the closed enum all raise
  `ValueError` / `TypeError` at construction or `add_chart()`
  time, not at save.

## What's new

### Chart 3D / Stock / Surface / ProjectedPie families

#### `BarChart3D` (RFC-046 §11.1)

```python
from wolfxl.chart import BarChart3D, Reference
import wolfxl

wb = wolfxl.Workbook()
ws = wb.active
ws.append(["Region", "Q1", "Q2", "Q3", "Q4"])
ws.append(["NA",      100,  120,  90,   140])
ws.append(["EU",      80,   95,   110,  85])
ws.append(["APAC",    60,   70,   85,   100])

chart = BarChart3D()
chart.title = "Quarterly Revenue (3D)"
# view_3d defaults: rot_x=15, rot_y=20, right_angle_axes=True,
# depth_percent=100 — override on the chart instance:
chart.view_3d.rot_x = 30
chart.view_3d.depth_percent = 150

data = Reference(ws, min_col=2, min_row=1, max_col=5, max_row=4)
cats = Reference(ws, min_col=1, min_row=2, max_row=4)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

ws.add_chart(chart, "G2")
wb.save("revenue_3d.xlsx")
```

#### `LineChart3D` / `AreaChart3D` / `PieChart3D` / `SurfaceChart3D`

All five 3D variants share the `view_3d` field; per-type defaults
follow openpyxl 3.1.x (see RFC-046 §11.1 table). `Pie3D` is the
openpyxl alias for `PieChart3D` and is exposed at the same name.

```python
from wolfxl.chart import LineChart3D, Reference

chart = LineChart3D()
# defaults: rot_x=15, rot_y=20, perspective=30,
#           right_angle_axes=False, depth_percent=100
chart.title = "3D revenue lines"
data = Reference(ws, min_col=2, min_row=1, max_col=4, max_row=10)
chart.add_data(data, titles_from_data=True)
ws.add_chart(chart, "G2")
```

#### `StockChart` (OHLC)

`StockChart` validates its 4-series Open / High / Low / Close
ordering at construction. Pass `series` in the wrong order and you
get a `ValueError` immediately, not on save.

```python
from wolfxl.chart import StockChart, Reference, Series

ws.append(["Date", "Open", "High", "Low", "Close"])
ws.append(["2026-04-01", 100.5, 105.0, 99.0, 104.0])
ws.append(["2026-04-02", 104.0, 107.5, 103.5, 106.0])
# … more rows …

chart = StockChart()
xvals = Reference(ws, min_col=1, min_row=2, max_row=N)
for col, name in [(2, "Open"), (3, "High"), (4, "Low"), (5, "Close")]:
    yvals = Reference(ws, min_col=col, min_row=2, max_row=N)
    chart.series.append(Series(yvals, xvals, title=name))

ws.add_chart(chart, "G2")
# Emits <c:stockChart> with <c:hiLowLines/> + <c:upDownBars/>
```

#### `SurfaceChart` / `SurfaceChart3D`

```python
from wolfxl.chart import SurfaceChart3D, Reference

chart = SurfaceChart3D(wireframe=True)
chart.title = "Surface plot"
data = Reference(ws, min_col=2, min_row=1, max_col=11, max_row=11)
cats = Reference(ws, min_col=1, min_row=2, max_row=11)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

ws.add_chart(chart, "M2")
```

#### `ProjectedPieChart` (pie-of-pie / bar-of-pie)

```python
from wolfxl.chart import ProjectedPieChart, Reference

chart = ProjectedPieChart(
    of_pie_type="bar",      # "bar" or "pie"
    split_type="percent",   # "auto" | "pos" | "percent" | "val" | "cust"
    split_pos=20,
    second_pie_size=75,
)
chart.title = "Sales — minor slices in secondary bar"
labels = Reference(ws, min_col=1, min_row=2, max_row=10)
data_ref = Reference(ws, min_col=2, min_row=1, max_row=10)
chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(labels)

ws.add_chart(chart, "G2")
```

### Validation at construction time

The 1.6.0 contract was lenient at construction time — many illegal
inputs only surfaced as cryptic XML errors at save. v1.6.1 fails
fast per RFC-046 §10.11. New raise sites:

```python
from wolfxl.chart import BarChart, Reference

# Empty series — no data to plot.
chart = BarChart()
ws.add_chart(chart, "G2")     # ValueError("BarChart has no series; add_data() first")

# Bad anchor — neither A1 nor a recognised anchor object.
chart = BarChart()
chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=4),
               titles_from_data=True)
ws.add_chart(chart, "not-an-anchor")
# ValueError("anchor must be an A1 string (e.g. 'G2') or an anchor object")

# Reference out-of-range.
Reference(ws, min_col=5, min_row=10, max_col=2, max_row=4)
# ValueError("min_col=5 > max_col=2")

# style outside 1..48
BarChart(style=99)            # ValueError("style must be in 1..48; got 99")

# gap_width outside 0..500
BarChart(gap_width=600)       # ValueError("gap_width must be in 0..500; got 600")

# hole_size outside 1..90
DoughnutChart(hole_size=95)   # ValueError("hole_size must be in 1..90; got 95")

# bubble_scale outside 0..300
BubbleChart(bubble_scale=500) # ValueError("bubble_scale must be in 0..300; got 500")

# Poly trendline order outside 2..6
trendline = Trendline(trendline_type="poly", order=10)
# ValueError("Polynomial trendline order must be in 2..6; got 10")

# display_blanks_as outside the closed enum
BarChart(display_blanks="nope")
# ValueError("display_blanks must be 'gap', 'span', or 'zero'; got 'nope'")
```

3D-only fields (`view_3d.rot_x` / `rot_y` / `perspective` /
`depth_percent` / `h_percent`) set on a 2D chart kind warn rather
than raise, matching openpyxl's tolerant posture there.

### Modify-mode high-level `add_chart` (Pod-γ′)

In v1.6.0 a high-level `add_chart(BarChart(...))` in modify mode
hit a warn-and-drop fallback in
`Workbook._flush_pending_charts_to_patcher`: the chart was queued
on `Worksheet._pending_charts` but never reached the patcher
because there was no dict→bytes bridge. The bytes-level escape
hatch (`Workbook.add_chart_modify_mode(sheet_name, chart_xml_bytes,
anchor)`) worked, but only for callers who already had pre-built
chart XML.

v1.6.1 wires the bridge:

```python
import wolfxl
from wolfxl.chart import BarChart, Reference

# Open an existing workbook in modify mode
wb = wolfxl.load_workbook("template.xlsx", modify=True)
ws = wb["Data"]

chart = BarChart()
chart.title = "Sales (auto-added in modify mode)"
data = Reference(ws, min_col=2, min_row=1, max_col=4, max_row=20)
cats = Reference(ws, min_col=1, min_row=2, max_row=20)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "G2")     # 1.6.1: bridges through serialize_chart_dict

wb.save("template.xlsx")       # NEW chart present, existing charts intact
```

The patcher Phase 2.5l (Sprint Μ Pod-γ) handles the rest exactly
as in 1.6.0.

### `serialize_chart_dict` — internal API

`wolfxl._backend.serialize_chart_dict(chart_dict)` is the new PyO3
helper that materialises a §10 chart-dict into the chart XML
bytes that the patcher's `queue_chart_add` expects. It is **not**
public surface; users should keep calling
`Worksheet.add_chart(chart, anchor)`. The helper exists so
`Workbook._flush_pending_charts_to_patcher` can do the same job
without round-tripping through a temporary write-mode file.

## Migration from 1.6.0

Most v1.6.0 → v1.6.1 changes are additive (new classes, new
validation surfaces). The breaking changes are limited to internals
that downstream code rarely touches.

### Chart-dict shape (internal, §10)

If your code calls `chart.to_rust_dict()` directly (this is
internal API and almost no one does), the dict shape changed in
the following ways:

- `solidFill` (camelCase) → `solid_fill` (snake_case). All
  graphical-property keys are now snake_case to match the Pod-α
  parser side.
- `tagname` (the openpyxl XML tag) is no longer passed through;
  Pod-β maps it to a short `kind` ("bar", "line", …) per §10.2.
- Per-type unique fields (`bar_dir`, `grouping`, `gap_width`,
  `overlap`, `smooth`, `scatter_style`, `radar_style`,
  `first_slice_ang`, `hole_size`, `bubble_3d`, `bubble_scale`,
  `show_neg_bubbles`, `size_represents`) are flat top-level keys,
  NOT nested in an `extras` dict.
- `view_3d` is a top-level dict on 3D-kind charts.
- `axes` is no longer a list; `x_axis`, `y_axis`, `z_axis` are flat
  top-level keys, each `None` when not used.

### Modify-mode warn-and-drop is gone

If your code relied on the v1.6.0 `RuntimeWarning` ("modify-mode
high-level add_chart is not yet wired; chart dropped") to fall
back to `Workbook.add_chart_modify_mode(..., chart_xml_bytes,
...)`, you can drop that fallback. The high-level path now works.
The bytes-level path keeps working unchanged.

### `NotImplementedError` on 3D / Stock / Surface / ProjectedPie is gone

Code that defensively caught `NotImplementedError` around
`BarChart3D()` / `LineChart3D()` / `PieChart3D()` / `Pie3D()` /
`AreaChart3D()` / `SurfaceChart()` / `SurfaceChart3D()` /
`StockChart()` / `ProjectedPieChart()` now succeeds. Audit any
`try/except NotImplementedError` around chart construction and
remove the fallback branch.

## Known gaps remaining (now → v2.0.0)

- **Pivot-chart linkage** (Sprint Ν / v2.0.0). A chart's
  `<c:pivotSource>` referencing a pivot cache definition cannot
  land before pivot caches are constructible. Pivot tables ship
  in v2.0.0.
- **Combination charts** (multi-plot — bar + line on shared axes).
  Not on the v1.6.x roadmap; v1.6.2 if requested.
- **Replace / delete existing charts.** The `add_chart` API stays
  additive only. `remove_chart` / `replace_chart` are not yet
  scheduled.
- **Conditional series formatting, data tables under chart,
  display units on value axes, dPt overrides.** Tracked in RFC-046
  §9 with deferral rationale; no v1.6.x commitment.

## Pods that landed 1.6.1

- **Pod-α′** — Rust parser extensions (`parse_chart_dict` covers
  §10 verbatim, 8 new `ChartKind` variants, per-3D-family emit
  fns) + the new `serialize_chart_dict` PyO3 export.
  Commits: `ba16137`, `37a3a6b`. Merged via `6c5425c`.
- **Pod-β′** — Python flat `to_rust_dict` (matches §10) + 8 new
  chart-family classes (`BarChart3D`, `LineChart3D`, `PieChart3D` /
  `Pie3D`, `AreaChart3D`, `SurfaceChart`, `SurfaceChart3D`,
  `StockChart`, `ProjectedPieChart`) + construction-time validation
  per §10.11.
  Commits: `70ebae1`, `a6d7442`. Merged via `8612189`.
- **Pod-γ′** — Modify-mode bridge (`_flush_pending_charts_to_patcher`
  uses `serialize_chart_dict`) + 3D-family round-trip tests in
  `tests/test_charts_3d.py`.
  Commits: `70ebae1`. Merged via `ed08aff`.
- **Pod-δ′** (this release scaffold) — RFC-046 §10/§11 SHA-log
  scaffold, 9 new `tests/parity/openpyxl_surface.py` entries,
  `KNOWN_GAPS.md` close-out, this release notes scaffold,
  `CHANGELOG.md` entry, removal of the module-level xfail on
  `tests/test_charts_write.py`.
  Commits: `330500e`. Merged via `620d606`.
- **Integrator finalize** — `parse_graphical_properties` reads both
  §10.9 and legacy graphical-property keys; chart-level
  `data_labels` propagates to series; `_DataLabelBase` accepts
  openpyxl-style `position=`; `Reference(None, …)` rejected;
  anchor bounds-check against XFD/1048576; 9 ratchet entries
  flipped to `wolfxl_supported=True`; `Pie3D` surface entry
  removed (openpyxl 3.1.x doesn't expose the alias).
  Commit: `70492ab`.

## SHA log

| Pod | Branch | Commits | Merge |
|---|---|---|---|
| α′ | `feat/sprint-mu-prime-pod-alpha` | `ba16137`, `37a3a6b` | `6c5425c` |
| β′ | `feat/sprint-mu-prime-pod-beta`  | `70ebae1`, `a6d7442` | `8612189` |
| γ′ | `feat/sprint-mu-prime-pod-gamma` | `70ebae1`             | `ed08aff` |
| δ′ | `feat/sprint-mu-prime-pod-delta` | `330500e`             | `620d606` |
| —  | (integrator finalize)            | `70492ab`             | n/a       |

Integrator finalize commit fills these placeholders, performs the
post-merge ratchet flip on the 9 new chart-3D entries in
`tests/parity/openpyxl_surface.py` (`shipped-1.6.1` tag), and tags
`v1.6.1`.

## Stats (post-1.6.1, projected)

- `cargo test --workspace --exclude wolfxl`: 1.6.0 + per-3D-family
  emit tests (Pod-α′) + `serialize_chart_dict` roundtrip tests.
- `pytest tests/`: **~1300+ → ~1340+ passed** (Pod-β′ adds ~30
  3D-family construction tests, Pod-γ′ adds ~10 modify-mode +
  bridge tests; Pod-δ′ removes the v1.6.0 xfail mark so the 37
  previously-xfailed tests count as passing). Final count filled
  in on integrator merge.
- `pytest tests/parity`: **~190+ → ~200+ passed** (9 new chart-3D
  ratchet entries flip to `wolfxl_supported=True` post-merge).
- `KNOWN_GAPS.md` "Out of scope" pruned: only **pivot tables +
  pivot charts (Sprint Ν / v2.0.0)** remain on the construction
  side. The openpyxl-parity surface for chart construction is
  exhaustively shipped after 1.6.1.

After 1.6.1 the only remaining v2.0.0 milestone is pivot tables +
pivot charts + public launch. Sprint Ν owns that slice; the
openpyxl-parity construction roadmap is otherwise complete.

## RFCs

- `Plans/rfcs/046-chart-construction.md` §10 (chart-dict contract)
  and §11 (3D / Stock / Surface / ProjectedPie families) — the
  Sprint Μ-prime preamble. The §12 SHA log gains the v1.6.1
  sub-table on integrator finalize.

## Acknowledgments

Sprint Μ-prime ("Mu-prime") closes out the Sprint Μ chart-construction
slice end-to-end. Thanks to everyone who file-bugged the v1.6.0
xfailed tests and the `NotImplementedError` on 3D variants — every
report drove this slice.
