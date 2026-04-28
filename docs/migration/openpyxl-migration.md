# Migrate from openpyxl

> **WolfXL 2.0** is the full openpyxl-replacement release: every
> construction idiom that openpyxl 3.1.x supports works with the
> same Python code, **including pivot tables, pivot caches, and
> pivot-chart linkage** (Sprint Ν / v2.0.0 closes the last gap).
> This guide walks through the API mapping for the eight idioms
> that cover the entire openpyxl construction surface.

## TL;DR — minimal import change

```python
# before
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image
from openpyxl.comments import Comment
from openpyxl.pivot.table import TableDefinition
from openpyxl.pivot.cache import CacheDefinition

# after
from wolfxl import load_workbook, Workbook
from wolfxl.styles import Font, PatternFill, Alignment, Border
from wolfxl.utils import get_column_letter, column_index_from_string
from wolfxl.chart import BarChart, Reference
from wolfxl.drawing.image import Image
from wolfxl.comments import Comment
from wolfxl.pivot import PivotTable, PivotCache
```

Almost every openpyxl import has the same name under `wolfxl`. The
exceptions live in [Compatibility Matrix](compatibility-matrix.md).

## What usually stays the same

```python
wb = load_workbook("data.xlsx")
ws = wb["Sheet"]
ws["A1"].value
ws["A1"].font.bold
ws["A1"] = "hello"
ws["B2"].font = Font(bold=True)
ws["B2"].fill = PatternFill("solid", fgColor="FFFF00")

ws.merge_cells("A1:B2")
ws.append(["a", "b", "c"])
for row in ws.iter_rows(min_row=2, values_only=True):
    print(row)
ws.cell(row=3, column=2, value="x")

wb.save("out.xlsx")
```

## Construction-side parity (complete in v2.0)

WolfXL 1.7 was the first release where the non-pivot construction-side
idioms all work end-to-end with the same code you'd write against
openpyxl 3.1.x. **WolfXL 2.0 closes the pivot-table gap** — the full
construction surface is now drop-in.

### Charts (v1.6 + v1.6.1)

Sixteen chart types ship at full openpyxl-3.1.x feature depth:

| Family    | Classes |
|-----------|---------|
| Bar       | `BarChart`, `BarChart3D` |
| Line      | `LineChart`, `LineChart3D` |
| Pie       | `PieChart`, `PieChart3D` (alias `Pie3D`), `DoughnutChart`, `ProjectedPieChart` |
| Area      | `AreaChart`, `AreaChart3D` |
| Scatter   | `ScatterChart` |
| Bubble    | `BubbleChart` |
| Radar     | `RadarChart` |
| Surface   | `SurfaceChart`, `SurfaceChart3D` |
| Stock     | `StockChart` (Open-High-Low-Close) |

```python
from wolfxl import Workbook
from wolfxl.chart import BarChart, Reference

wb = Workbook()
ws = wb.active
ws.append(["Region", "Q1", "Q2", "Q3", "Q4"])
ws.append(["NA", 100, 110, 120, 140])
ws.append(["EU", 80,  95,  110, 85])
ws.append(["APAC", 60, 70, 85, 100])

chart = BarChart()
chart.title = "Quarterly Revenue"
chart.style = 10
chart.x_axis.title = "Region"
chart.y_axis.title = "Revenue (USD)"

data = Reference(ws, min_col=2, min_row=1, max_col=5, max_row=4)
cats = Reference(ws, min_col=1, min_row=2, max_row=4)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

ws.add_chart(chart, "G2")
wb.save("out.xlsx")
```

Sprint Ξ (v1.7) adds two more chart-management methods:

```python
ws.remove_chart(chart)            # mirrors openpyxl ws._charts.remove(chart)
ws.replace_chart(old, new)        # convenience: keeps anchor + list position
```

Modify-mode `add_chart` works with any of the 16 families:

```python
wb = load_workbook("template.xlsx", modify=True)
ws = wb.active
chart = BarChart()
# ... configure chart ...
ws.add_chart(chart, "B10")
wb.save("template.xlsx")
```

### Images (v1.5)

```python
from wolfxl.drawing.image import Image

img = Image("logo.png")          # PNG / JPEG / GIF / BMP
img.width = 200
img.height = 100
ws.add_image(img, "A1")          # one-cell anchor

# Two-cell anchor:
from wolfxl.drawing import TwoCellAnchor, AnchorMarker
img2 = Image("chart.png")
img2.anchor = TwoCellAnchor(
    _from=AnchorMarker(col=2, row=2, colOff=0, rowOff=0),
    to=AnchorMarker(col=8, row=10, colOff=0, rowOff=0),
)
ws.add_image(img2)
```

### Encrypted reads + writes (v1.3 read; v1.5 write)

```python
# read
wb = load_workbook("encrypted.xlsx", password="hunter2")

# write — Agile (AES-256 / SHA-512) on save
wb = Workbook()
# ... build workbook ...
wb.save("secret.xlsx", password="hunter2")
```

Install with `pip install wolfxl[encrypted]` (pulls `msoffcrypto-tool`).

### Streaming reads — `read_only=True` (v1.3)

```python
wb = load_workbook("huge.xlsx", read_only=True)
for row in wb.active.iter_rows(values_only=True):
    process(row)
```

Auto-engages for sheets with > 50,000 rows even when the caller didn't
opt in. Streaming cells are immutable — assignment raises `RuntimeError`.

### `.xlsb` and `.xls` reads (v1.4)

```python
wb_b = load_workbook("data.xlsb")     # Binary OOXML
wb_x = load_workbook("data.xls")      # Legacy BIFF8

# Modify mode + read_only + password are xlsx-only.
# To round-trip a .xlsb to .xlsx, transcribe via a fresh Workbook().
```

### Modify-mode mutations (v1.0 / v1.1)

Every T1.5 mutation that openpyxl supports works in WolfXL modify mode
(surgical ZIP patching — much faster than a full DOM rewrite):

```python
wb = load_workbook("template.xlsx", modify=True)

# Document properties
wb.properties.title = "Q4 2025 Report"
wb.properties.creator = "Finance Team"

# Defined names
from wolfxl.defined_names import DefinedName
wb.defined_names["RevenueRange"] = DefinedName(
    name="RevenueRange",
    attr_text="Sheet1!$B$2:$B$100",
)

# Comments + hyperlinks + tables + DV + CF
from wolfxl.comments import Comment
ws["A1"].comment = Comment("Reviewed", "auditor")

from wolfxl.cell.hyperlink import Hyperlink
ws["B1"].hyperlink = Hyperlink(target="https://example.com", display="Source")

from wolfxl.worksheet.table import Table, TableStyleInfo
ws.add_table(Table(displayName="Sales", ref="A1:D10"))

from wolfxl.worksheet.datavalidation import DataValidation
dv = DataValidation(type="whole", operator="between", formula1=0, formula2=100)
dv.add("A1:A100")
ws.data_validations.append(dv)

from wolfxl.formatting.rule import CellIsRule
from wolfxl.styles import PatternFill
ws.conditional_formatting.add(
    "B2:B100",
    CellIsRule(operator="greaterThan", formula=["100"],
               fill=PatternFill(fgColor="FFFF00", patternType="solid")),
)

wb.save("template.xlsx")
```

### Structural ops (v1.1)

```python
# Insert / delete rows + columns; everything shifts (formulas, hyperlinks,
# CF rules, data validations, defined names, tables, conditional formatting).
ws.insert_rows(idx=2, amount=3)
ws.delete_rows(idx=10, amount=1)
ws.insert_cols(idx=2, amount=2)
ws.delete_cols(idx=5, amount=1)

# Move a 2D range
ws.move_range("B2:D5", rows=3, cols=1)

# Copy a worksheet (deep-clones tables, DV, CF, sheet-scoped defined
# names, charts with cell-range re-pointing)
ws_copy = wb.copy_worksheet(wb["Source"])

# Reorder sheets
wb.move_sheet("Sheet2", offset=-1)
```

### Rich text (v1.3)

```python
from wolfxl.cell.rich_text import CellRichText, TextBlock, InlineFont

cell = ws["A1"]
cell.value = CellRichText([
    TextBlock(InlineFont(b=True), "Bold "),
    "and ",
    TextBlock(InlineFont(color="FF0000"), "red"),
])

# Reading:
wb = load_workbook("rich.xlsx", rich_text=True)
print(ws["A1"].value)             # CellRichText(...)
print(ws["A1"].rich_text)         # always returns CellRichText (or None)
```

### Pivot tables (Sprint Ν / v2.0)

WolfXL 2.0 closes the last construction-side gap. Pivot caches,
pivot tables, and pivot-chart linkage all ship; the v0.5+
`_make_stub` is replaced with real classes.

```python
import wolfxl
from wolfxl.chart import Reference, BarChart
from wolfxl.pivot import PivotCache, PivotTable

wb = wolfxl.Workbook()
ws = wb.active
ws.append(["region", "quarter", "product", "revenue"])
ws.append(["NA",     "Q1",      "Widget",  100])
ws.append(["NA",     "Q2",      "Widget",  120])
ws.append(["EU",     "Q1",      "Widget",   80])
ws.append(["EU",     "Q2",      "Widget",   95])
# ... fill source data ...

# 1. Build a cache around the source range:
src = Reference(ws, min_col=1, min_row=1, max_col=4, max_row=100)
cache = wb.add_pivot_cache(PivotCache(source=src))

# 2. Build a pivot table referencing the cache:
pt = PivotTable(
    cache=cache,
    location="F2",
    rows=["region"],
    cols=["quarter"],
    data=[("revenue", "sum")],          # or [DataField("revenue", function="sum")]
)
ws.add_pivot_table(pt)

# 3. (Optional) link a chart to the pivot:
chart = BarChart()
chart.title = "Revenue by region × quarter"
chart.pivot_source = pt                  # emits <c:pivotSource>
ws.add_chart(chart, "F18")

wb.save("pivot.xlsx")
```

Open `pivot.xlsx` in Excel, LibreOffice, or read it with
`openpyxl.load_workbook(...)` — the pivot's data is already
populated; **no refresh-on-open is required**. WolfXL is the only
Python OOXML library that constructs pivot tables with
pre-aggregated `pivotCacheRecords` (openpyxl preserves them on
round-trip but doesn't construct them; XlsxWriter doesn't support
pivots at all).

#### Import paths

| openpyxl path | WolfXL path |
|---|---|
| `from openpyxl.pivot.table import TableDefinition` | `from wolfxl.pivot import PivotTable` |
| `from openpyxl.pivot.cache import CacheDefinition` | `from wolfxl.pivot import PivotCache` |
| `from openpyxl.pivot.fields import PivotField, DataField, RowField, ColumnField, PageField` | `from wolfxl.pivot import PivotField, DataField, RowField, ColumnField, PageField` |
| `from openpyxl.pivot.table import PivotItem, Location, PivotTableStyleInfo` | `from wolfxl.pivot import PivotItem, Location, PivotTableStyleInfo` |
| `Reference(ws, min_col=..., min_row=..., max_col=..., max_row=...)` | Same — re-uses `wolfxl.chart.Reference` (mirrors openpyxl 3.1.x's shared reference type) |

#### API differences

openpyxl exposes a one-step `ws.add_pivot(table)` (where the
cache is implied). WolfXL splits cache and table into two steps:

```python
# wolfxl — explicit two-step
cache = wb.add_pivot_cache(PivotCache(source=src))
pt    = PivotTable(cache=cache, location="F2", ...)
ws.add_pivot_table(pt)
```

The split exists because OOXML caches are workbook-scoped (one
cache can serve multiple tables) while tables are sheet-scoped.
Aliasing a cache to several tables is the same idiom wolfxl uses
for shared image media on `copy_worksheet`.

#### Aggregator functions

`DataField(function=...)` accepts:

| wolfxl name | OOXML name |
|---|---|
| `"sum"` | Sum |
| `"count"` | Count |
| `"average"` | Average |
| `"max"` | Max |
| `"min"` | Min |
| `"product"` | Product |
| `"count_nums"` | CountNums |
| `"std_dev"` | StdDev |
| `"std_dev_p"` | StdDevp |
| `"var"` | Var |
| `"var_p"` | Varp |

The bare-string axis spec (`rows=["region"]`) is shorthand for
`rows=[RowField("region")]` and uses the field's column header
as the caption. Use the explicit builders (`RowField`,
`ColumnField`, `DataField`, `PageField`) for custom captions,
custom subtotals, or custom sort orders.

#### Limits (deferred to v2.1+)

- **Slicers** — `xl/slicers/` and `xl/slicerCaches/` are not yet
  constructible. Pivots round-trip without slicers; if a source
  workbook has them, modify-mode round-trip preserves them.
- **Calculated fields** (`<calculatedField>`) and **calculated
  items** (`<calculatedItem>`) — formula expressions in the
  pivot's field list / row-or-col items. v2.1.
- **GroupItems** (date / range grouping —
  `<fieldGroup base="N"><rangePr><groupItems/></rangePr></fieldGroup>`)
  — non-trivial recursion. v2.1.
- **OLAP / external pivot caches** — needs the PowerPivot
  data-model (`xl/model/`). Out of scope permanently.
- **Pivot-table styling beyond the named-style picker** —
  themes, banded formats, pivot-cell conditional formatting.
  v2.1.
- **In-place pivot edits in modify mode** beyond
  `add_pivot_table` — editing an existing pivot's source range,
  field ordering, subtotals, etc. v2.2.

#### Caveat — empty cache

If the chart's `pivot_source` points at a `PivotTable` whose
cache walked an empty source range, the chart will render blank
in Excel. This matches openpyxl's behaviour and is documented in
RFC-049 §10.4. Validate the source range has at least one data
row before constructing.

## What to validate during migration

1. **Style fidelity** in your critical sheets — open the saved
   workbook in Excel and diff visually. WolfXL's
   `tests/parity/openpyxl_surface.py` ratchet tracks every flaky
   serialisation.
2. **Formula behavior** in your downstream consumers — formulas are
   preserved verbatim; cached results are recomputed when Excel opens.
3. **Pivot tables** — fully constructible from Python (v2.0+) with
   pre-aggregated records, so the saved workbook opens in any
   OOXML reader (Excel / LibreOffice / openpyxl) without
   refresh-on-open. See "Pivot tables (Sprint Ν / v2.0)" above.
4. **Rare openpyxl APIs** — see the [Compatibility Matrix](compatibility-matrix.md)
   for anything that's `Partial` or `Not Yet`.

## Migration playbook

1. Swap imports in **one** workflow.
2. Run your existing test suite — wolfxl's read+write is a strict
   superset of openpyxl's, so most tests should pass unchanged.
3. Compare a representative output workbook in Excel side-by-side
   with the openpyxl-produced version.
4. Measure runtime/memory — see [Performance](../performance/benchmark-results.md).
5. Roll out gradually to other pipelines.

## Edge cases worth knowing

- **`Worksheet.max_row` / `max_column`** — public properties (not
  methods).
- **`merged_cells`** — backed by `_MergedCellsProxy` in read mode.
- **`Cell.coordinate`** — always uppercase (e.g. `"A1"`).
- **`Cell.number_format`** — accepts the same Excel format strings
  openpyxl does.
- **`copy_worksheet`** — diverges from openpyxl in five
  documented ways (preserves tables, DV, CF, sheet-scoped defined
  names, image media). See `tests/parity/KNOWN_GAPS.md`
  "RFC-035 — `copy_worksheet` divergences from openpyxl" for the
  full record. WolfXL's behaviour is strictly more useful in every
  case.

## When to keep openpyxl alongside

- You construct OpenDocument (`.ods`) files.
- You need the deepest features of openpyxl's chart layer
  (combination charts, `<c:displayUnits>` on value axes,
  per-data-point overrides via `dPt`).
- You construct pivot tables that need slicers, calculated
  fields / items, GroupItems, or pivot-styling beyond the
  named-style picker. Those are deferred to v2.1+ (see "Limits"
  in the Pivot tables section above).

For everything else, v2.0 is a drop-in replacement.

## Further reading

- [Compatibility Matrix](compatibility-matrix.md) — exhaustive table
  of API support.
- [Legacy Shim](legacy-shim.md) — `excelbench_rust` → `wolfxl._rust`
  shim notes (only relevant if you're upgrading from a pre-1.0
  ExcelBench install).
- [Performance](../performance/benchmark-results.md) — v2.0 numbers (incl. pivot construction).
- [Release notes](../release-notes-2.0.md) — full v2.0 changelog.
- [Trust](../trust/) — fidelity, security, supply-chain provenance.
