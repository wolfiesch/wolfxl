# RFC-061 — Advanced pivot construction (slicers + calc fields + calc items + GroupItems + styling)

> **Status**: Approved
> **Phase**: 5 (2.0 — Sprint Ο)
> **Depends-on**: 047 (pivot caches), 048 (pivot tables), 049 (pivot charts), 026 (CF — for `dxfId` allocation)
> **Unblocks**: v2.0.0 launch
> **Pod**: 3

## 1. Goal

Close every Tier 3 pivot gap surfaced by the audit, EXCEPT
OLAP / external pivot caches (per user scope decision —
no PowerPivot data-model support).

This RFC covers four sub-features:

1. **Slicers** — interactive UI filter cards (10 days)
2. **Calculated fields** — formula-based cache fields (3 days)
3. **Calculated items** — formula-based row/col items (3 days)
4. **GroupItems** — date / range grouping in cache (7 days)
5. **Pivot styling beyond named-style picker** — pivot-scoped CF, banded formats (5 days)

Total: ~28 days, ~4500 LOC.

## 2. Public API

### 2.1 Slicers

```python
import wolfxl

# Slicer cache — workbook-scoped, references a pivot cache.
slicer_cache = wolfxl.pivot.SlicerCache(
    name="region_slicer_cache",
    source_pivot_cache=cache,      # PivotCache instance
    field="region",
    sort_order="ascending",
    custom_list_sort=False,
    hide_items_with_no_data=False,
)
wb.add_slicer_cache(slicer_cache)

# Slicer presentation — sheet-scoped.
slicer = wolfxl.pivot.Slicer(
    name="region_slicer",
    cache=slicer_cache,
    caption="Filter by Region",
    row_height=204,                # EMU
    column_count=1,
    show_caption=True,
    style="SlicerStyleLight1",
    locked=True,
)
ws.add_slicer(slicer, anchor="H2")
```

### 2.2 Calculated fields

```python
# Add to cache
cache.add_calculated_field(
    name="profit",
    formula="= revenue - cost",
)
# Field becomes available in PivotTable like any other:
pt = wolfxl.pivot.PivotTable(
    cache=cache, location="A1",
    rows=["region"],
    data=[("profit", "sum")],     # use the calc field
)
```

Formula uses pivot field names as bare identifiers; the
`wolfxl-formula` crate (RFC-012) parses + validates. Excel
evaluates on open. wolfxl does not pre-compute calc-field
values into cache records.

### 2.3 Calculated items

```python
# Add to a specific field of the table
pt.add_calculated_item(
    field="region",
    item_name="east_west_combined",
    formula='= east + west',
)
```

Calc items live inside pivot table XML, not cache XML.

### 2.4 GroupItems

```python
# Group by date precision
cache.group_field(
    field="order_date",
    by="months",        # "years"|"quarters"|"months"|"days"|"hours"|"minutes"|"seconds"
)

# Or by numeric range
cache.group_field(
    field="age",
    start=0,
    end=100,
    interval=10,        # 0-9, 10-19, 20-29, ...
)
```

Recursive: a grouped field can be source for further grouping
(e.g. group by year, then by quarter within year). RFC-061
caps recursion depth at 4.

### 2.5 Pivot styling

```python
pt.style_info = wolfxl.pivot.PivotTableStyleInfo(
    name="PivotStyleMedium2",
    show_row_headers=True,
    show_col_headers=True,
    show_row_stripes=True,
    show_col_stripes=False,
    show_last_column=True,
)

# Pivot-scoped cell-level formatting (specific cells in the pivot)
pt.add_format(
    pivot_area=wolfxl.pivot.PivotArea(
        field=0,                    # field index
        type="data",                # "all"|"data"|"labelOnly"|"button"
        data_only=True,
        grand_row=False,
        grand_col=False,
    ),
    dxf=DifferentialStyle(font=Font(bold=True), fill=PatternFill(...)),
    action="formatting",            # "formatting"|"blank"
)

# Pivot-scoped CF
pt.add_conditional_format(
    rule=ColorScale(...),
    pivot_area=wolfxl.pivot.PivotArea(field=1, type="data"),
)
```

## 3. OOXML output

### 3.1 Slicer parts

New parts:
- `xl/slicers/slicer{N}.xml` — slicer presentation per worksheet.
- `xl/slicerCaches/slicerCache{N}.xml` — slicer cache.
- Workbook `<extLst>` extension carrying `<x14:slicerCaches>`
  collection (extension URI:
  `{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}`).
- Sheet `<extLst>` extension carrying `<x14:slicerList>`
  (extension URI: `{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}`).
- Drawing-level rels — slicers anchor like images via
  `<sl:slicer>` extension under `<extLst>` of the
  `<xdr:graphicFrame>`.

### 3.2 Calculated fields

Inserted into pivot cache XML:

```xml
<cacheFields count="N">
  ...
  <cacheField name="profit" numFmtId="0">
    <fieldGroup base="..."/>
  </cacheField>
</cacheFields>
<calculatedItems count="1">
  <calculatedItem fld="N" formula="revenue - cost"/>
</calculatedItems>
```

### 3.3 GroupItems

```xml
<cacheField name="order_date" numFmtId="14">
  <sharedItems containsDate="1" .../>
  <fieldGroup par="0" base="0">
    <rangePr groupBy="months" startDate="2020-01-01T00:00:00"
             endDate="2025-12-31T23:59:59"/>
    <groupItems count="14">
      <s v="<01/01/2020"/>
      <s v="Jan"/>
      ...
      <s v="Dec"/>
      <s v=">12/31/2025"/>
    </groupItems>
  </fieldGroup>
</cacheField>
```

Recursive grouping uses `par="N"` to point at the parent
group's index.

### 3.4 PivotArea / Format / CF

Inside pivot table XML:

```xml
<formats count="N">
  <format dxfId="0" action="formatting">
    <pivotArea field="0" type="data" dataOnly="1"/>
  </format>
</formats>
<conditionalFormats count="N">
  <conditionalFormat scope="data" type="all" priority="1">
    <pivotAreas count="1">
      <pivotArea field="1" type="data"/>
    </pivotAreas>
    <!-- references CF rule via dxfId -->
  </conditionalFormat>
</conditionalFormats>
```

## 4. Modify mode

Patcher Phase 2.5p (after pivots — Phase 2.5m, before charts —
Phase 2.5l) drains:
- Slicer caches → workbook extLst splice.
- Slicers → sheet extLst splice + drawing graphicFrame.
- Calc fields → modify in-place inside pivot cache definition
  XML (re-emit).
- Calc items → re-emit pivot table XML.
- Group items → re-emit pivot cache definition XML.
- Pivot CF / formats → re-emit pivot table XML.

**Phase ordering decision**: pivots emitted first (cache + table
in Phase 2.5m), then advanced extensions overlaid in Phase 2.5p.
The pivot table's pre-computed `<rowItems>` / `<colItems>` are
NOT re-computed when calc-fields are added — Excel re-computes
on open. This matches the v2.0 "no runtime evaluation" stance.

## 5. Native writer

`crates/wolfxl-pivot` extends with:
- `slicer.rs` — Slicer + SlicerCache models + emit (~600 LOC).
- `calc.rs` — CalculatedField + CalculatedItem models + emit
  (~200 LOC).
- `group.rs` — FieldGroup + GroupItems models + emit (~400 LOC).
- `styling.rs` — PivotArea + Format + PivotConditionalFormat
  (~300 LOC).

## 6. RFC-035 deep-clone

- **Slicer caches**: workbook-scoped → alias on copy_worksheet
  (one cache serves source + clone).
- **Slicers**: sheet-scoped → deep-clone with cache_id
  preserved (alias) and presentation re-emitted.
- **Calc fields**: cache-scoped → already aliased via the
  parent cache.
- **Calc items**: table-scoped → deep-clone with table.
- **Group items**: cache-scoped → aliased.
- **Pivot styling**: table-scoped → deep-clone with table.

## 7. Testing

- `tests/test_pivot_slicers.py` (~40 tests).
- `tests/test_pivot_calculated_fields.py` (~30 tests).
- `tests/test_pivot_calculated_items.py` (~25 tests).
- `tests/test_pivot_group_items.py` (~30 tests).
- `tests/test_pivot_advanced_styling.py` (~25 tests).
- `tests/parity/test_pivot_advanced_parity.py` (~10 tests).
- `tests/diffwriter/test_pivot_advanced_*.py` (~5 tests).
- LibreOffice + Excel-online fixture regression (~5 manual).

## 8. Out of scope (v2.1+)

- **OLAP / external pivot caches** — `xl/model/`, PowerPivot
  data-model. User scope decision: no.
- **In-place pivot edits** beyond `add_*` (editing existing
  field assignments, subtotal toggling).
- **Pivot table evaluation** — wolfxl pre-computes pivot
  values on construction (Sprint Ν Option A); calc-field
  values stay un-evaluated and Excel computes on open.

## 9. References

- ECMA-376 Part 1 §18.10 (CT_PivotCacheDefinition extensions)
- ECMA-376 Part 4 — Slicers (extension namespace)
- openpyxl 3.1.x `openpyxl.pivot.cache` source (FieldGroup,
  GroupItems).
- openpyxl 3.1.x `openpyxl.pivot.table` source (PivotArea,
  Format, ConditionalFormat).

## 10. Dict contracts

### 10.1 Slicer cache

```python
{
    "name": str,
    "source_pivot_cache_id": int,
    "source_field_index": int,
    "sort_order": "ascending" | "descending" | "none",
    "custom_list_sort": bool,
    "hide_items_with_no_data": bool,
    "show_missing": bool,
    "items": [
        {"name": str, "hidden": bool, "no_data": bool},
        ...
    ],
}
```

### 10.2 Slicer

```python
{
    "name": str,
    "cache_name": str,
    "caption": str,
    "row_height": int,           # EMU
    "column_count": int,
    "show_caption": bool,
    "style": str | None,
    "locked": bool,
    "anchor": str,               # A1 cell
}
```

### 10.3 Calculated field

```python
{"name": str, "formula": str, "data_type": "string"|"number"|"boolean"|"date"}
```

### 10.4 Calculated item

```python
{"field_name": str, "item_name": str, "formula": str}
```

### 10.5 Field group

```python
{
    "field_index": int,           # base field
    "parent_index": int | None,   # for nested grouping
    "kind": "date" | "range" | "discrete",
    "date": {
        "group_by": str,          # "years"|"quarters"|...|"seconds"
        "start_date": str,
        "end_date": str,
    } | None,
    "range": {
        "start": float,
        "end": float,
        "interval": float,
    } | None,
    "items": [
        {"name": str},
        ...
    ],
}
```

### 10.6 Pivot area

```python
{
    "field": int,                 # field index
    "type": "all" | "data" | "labelOnly" | "button",
    "data_only": bool,
    "label_only": bool,
    "grand_row": bool,
    "grand_col": bool,
    "cache_index": int | None,
    "axis": str | None,
    "field_position": int | None,
}
```

### 10.7 Format

```python
{
    "action": "formatting" | "blank",
    "dxf_id": int,
    "pivot_area": <§10.6>,
}
```

PyO3 bindings:
- `serialize_slicer_cache_dict(d) -> bytes`
- `serialize_slicer_dict(d) -> bytes`
- `serialize_pivot_with_advanced_dict(table_d, cache_d, ...) -> bytes`
  (re-emits the entire pivot table or pivot cache XML with calc
  fields / items / groups / formats / CF inlined).

## 11. Acceptance

- All 5 sub-features constructible.
- openpyxl reads back slicers + calc fields + calc items +
  group items + pivot CF correctly.
- LibreOffice renders slicers (manual verification).
- Recursive group-by (e.g. year → quarter → month) works
  for date fields.
- Pivot CF rule references workbook-scoped dxf table (uses
  RFC-026's `dxfId` allocator).
- ~155 tests green.
- `KNOWN_GAPS.md` "Out of scope" reduces to OLAP + in-place
  edits + runtime evaluation only.
