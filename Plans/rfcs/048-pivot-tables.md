# RFC-048: Pivot tables — `wolfxl.pivot.PivotTable` (layout + emit)

Status: Approved (pre-dispatch contract spec authored; pods not yet dispatched)
Owner: Sprint Ν Pods α / β / γ
Phase: 5 (2.0)
Estimate: XL
Depends-on: RFC-010 (rels), RFC-013 (patcher), RFC-035 (deep-clone), RFC-046 (chart Reference shared), **RFC-047 (pivot caches)**
Unblocks: RFC-049 (pivot charts), v2.0.0 launch

## 1. Background — Problem Statement

Companion to RFC-047. The cache holds the data; the
**pivot table** holds the layout (which fields go on rows / cols /
data / page filters; which aggregator (sum / avg / count / etc.);
totals, styling, captions). One workbook can have N pivot tables
sharing M pivot caches.

The user-visible API:

```python
from wolfxl.pivot import PivotTable, PivotCache, Reference, DataField

src = Reference(ws, min_col=1, min_row=1, max_col=4, max_row=100)
cache = wb.add_pivot_cache(PivotCache(source=src))
pt = PivotTable(
    cache=cache,
    location="F2",
    rows=["region"],
    cols=["quarter"],
    data=[DataField("revenue", function="sum")],
    page=["customer"],
    name="MyPivot",
)
ws.add_pivot_table(pt)
```

Today every line raises `NotImplementedError` because the module
is `_make_stub`. RFC-048 closes the construction path: typed
Python class, `to_rust_dict()` returning the §10 shape, Rust crate
emit (`crates/wolfxl-pivot/src/emit/pivot_table.rs`), patcher
integration (`Worksheet.add_pivot_table(pt)`), and a Sprint Μ-prime
lesson #12 §10 contract specified verbatim before pod dispatch.

## 2. OOXML Spec Surface

ECMA-376 Part 1 §18.10.1.73 (`pivotTableDefinition`).

### 2.1 Part

| Path | Content type | Rels type |
|---|---|---|
| `xl/pivotTables/pivotTable{N}.xml` | `application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml` | `…/relationships/pivotTable` |

The table has its own rels file pointing back at the cache:

```xml
<!-- xl/pivotTables/_rels/pivotTable1.xml.rels -->
<Relationships xmlns="…">
  <Relationship Id="rId1"
    Type="…/relationships/pivotCacheDefinition"
    Target="../pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>
```

The owning sheet's rels file points at the table:

```xml
<!-- xl/worksheets/_rels/sheet1.xml.rels -->
<Relationship Id="rId5"
  Type="…/relationships/pivotTable"
  Target="../pivotTables/pivotTable1.xml"/>
```

### 2.2 Skeleton

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      name="MyPivot"
                      cacheId="0"
                      dataOnRows="0"
                      applyNumberFormats="0"
                      applyBorderFormats="0"
                      applyFontFormats="0"
                      applyPatternFormats="0"
                      applyAlignmentFormats="0"
                      applyWidthHeightFormats="1"
                      dataCaption="Values"
                      updatedVersion="6"
                      minRefreshableVersion="3"
                      showCalcMbrs="0"
                      useAutoFormatting="1"
                      itemPrintTitles="1"
                      createdVersion="6"
                      indent="0"
                      outline="1"
                      outlineData="1"
                      multipleFieldFilters="0">
  <location ref="F2:I20" firstHeaderRow="0" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="4">
    <pivotField axis="axisRow" showAll="0">
      <items count="5">
        <item x="0"/>
        <item x="1"/>
        <item x="2"/>
        <item x="3"/>
        <item t="default"/>
      </items>
    </pivotField>
    <pivotField dataField="1" showAll="0"/>
    <pivotField axis="axisCol" showAll="0">
      <items count="5">
        <item x="0"/><item x="1"/><item x="2"/><item x="3"/>
        <item t="default"/>
      </items>
    </pivotField>
    <pivotField axis="axisPage" showAll="0">
      <items count="2"><item x="0"/><item t="default"/></items>
    </pivotField>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <rowItems count="5">
    <i><x/></i>
    <i><x v="1"/></i>
    <i><x v="2"/></i>
    <i><x v="3"/></i>
    <i t="grand"><x/></i>
  </rowItems>
  <colFields count="1"><field x="2"/></colFields>
  <colItems count="5">
    <i><x/></i><i><x v="1"/></i><i><x v="2"/></i><i><x v="3"/></i>
    <i t="grand"><x/></i>
  </colItems>
  <pageFields count="1">
    <pageField fld="3" hier="-1"/>
  </pageFields>
  <dataFields count="1">
    <dataField name="Sum of revenue" fld="1" baseField="0" baseItem="0"/>
  </dataFields>
  <pivotTableStyleInfo name="PivotStyleLight16"
                       showRowHeaders="1" showColHeaders="1"
                       showRowStripes="0" showColStripes="0"
                       showLastColumn="1"/>
</pivotTableDefinition>
```

### 2.3 The 25 attrs we support in v2.0

`TableDefinition` has 69 attrs. Sprint Ν supports the **25 used in
real-world fixtures**; the rest emit Excel defaults via attribute
omission. The supported attrs (with their default if omitted):

| Attr | Default | Notes |
|---|---|---|
| `name` | required | |
| `cacheId` | required | |
| `dataOnRows` | `false` | data fields → cols by default |
| `dataCaption` | `"Values"` | |
| `grandTotalCaption` | `None` (omit) | |
| `errorCaption` | `None` (omit) | |
| `missingCaption` | `None` (omit) | |
| `showError` | `false` | |
| `showMissing` | `true` | |
| `pivotTableStyle` | `None` (omit) | the pre-2010 attr; superseded by `pivotTableStyleInfo` child |
| `applyNumberFormats` | `false` | |
| `applyBorderFormats` | `false` | |
| `applyFontFormats` | `false` | |
| `applyPatternFormats` | `false` | |
| `applyAlignmentFormats` | `false` | |
| `applyWidthHeightFormats` | `true` | |
| `useAutoFormatting` | `true` | |
| `outline` | `true` | |
| `outlineData` | `true` | |
| `compact` | `true` | |
| `compactData` | `true` | |
| `rowGrandTotals` | `true` | |
| `colGrandTotals` | `true` | |
| `multipleFieldFilters` | `false` | |
| `itemPrintTitles` | `true` | |

The other 44 attrs (`asteriskTotals`, `mdxSubqueries`,
`fieldListSortAscending`, etc.) emit at Excel default by being
omitted. v2.1 expands the supported set if user demand surfaces.

## 3. openpyxl Reference

`openpyxl/pivot/table.py` (~41 kB). Key classes:

* `TableDefinition` (69 attrs) — root.
* `Location` — `<location>`.
* `PivotField` — per-cache-field appearance/role.
* `RowFields`, `ColFields`, `PageFields`, `DataFields` —
  axis-specific child lists.
* `RowItems`, `ColItems` — pre-computed pivot row/col labels.
* `DataField` — aggregator + numFmtId.
* `PageField` — slicer-style filter.
* `PivotTableStyleInfo` — named-style picker.

What we DO NOT copy:
* `formats` collection (per-cell formatting overrides). Out of
  scope; v2.1.
* `conditionalFormats`. Out of scope.
* `chartFormats`. Tied to PivotChart legacy formatting; v2.1.
* `pivotHierarchies` / OLAP. Out permanently.
* `kpis`. Out permanently.
* `filters` collection (label / value filters). v2.1.

## 4. WolfXL Surface Area

### 4.1 Python

* `python/wolfxl/pivot/__init__.py` — re-exports.
* `python/wolfxl/pivot/_table.py` — `PivotTable`, `PivotField`,
  `DataField`, `RowField`, `ColumnField`, `PageField`,
  `Location`, `PivotTableStyleInfo`. Per-class `to_rust_dict()`.
* `python/wolfxl/_worksheet.py` — `Worksheet.add_pivot_table(pt)`
  (modify mode → enqueue on patcher; write mode → push onto
  worksheet's `_pivot_tables` list).

### 4.2 Patcher

* `crates/wolfxl-pivot/src/emit/pivot_table.rs` (Pod-α) —
  `pivot_table_xml(pt: &PivotTable, w: &mut Writer)`.
* `src/wolfxl/mod.rs` (Pod-γ) — PyO3 binding
  `wolfxl._rust.serialize_pivot_table_dict(d) -> bytes`.
* `XlsxPatcher.queue_pivot_table_add(sheet_name, table_xml)` +
  Phase 2.5m: allocate `pivotTable{N}.xml`, splice into sheet's
  rels, add to content-types.

### 4.3 Native writer

`crates/wolfxl-writer/src/emit/pivot_table.rs` (new) wraps the
crate's emit and handles ZIP entry creation. Worksheet's
`_pivot_tables: Vec<PivotTable>` populated by Python.

## 5. Implementation Sketch

### 5.1 Python `PivotTable` constructor

```python
class PivotTable:
    def __init__(
        self,
        *,
        cache: PivotCache,
        location: str | tuple[str, str],          # "F2" or ("F2", "I20")
        rows: list[str | RowField] | None = None,
        cols: list[str | ColumnField] | None = None,
        data: list[str | DataField] | None = None,
        page: list[str | PageField] | None = None,
        name: str = "PivotTable1",
        style_name: str = "PivotStyleLight16",
        row_grand_totals: bool = True,
        col_grand_totals: bool = True,
        outline: bool = True,
        compact: bool = True,
    ):
        if not isinstance(cache, PivotCache):
            raise TypeError("PivotTable.cache must be a PivotCache")
        if cache._cache_id is None:
            raise ValueError(
                "PivotTable.cache must have been registered via "
                "Workbook.add_pivot_cache(cache) first"
            )
        self.cache = cache
        self.location = self._normalize_location(location)
        self.rows = self._normalize_field_list(rows, "row")
        self.cols = self._normalize_field_list(cols, "col")
        self.data = self._normalize_field_list(data, "data")
        self.page = self._normalize_field_list(page, "page")
        self.name = name
        self.style_name = style_name
        self.row_grand_totals = row_grand_totals
        self.col_grand_totals = col_grand_totals
        self.outline = outline
        self.compact = compact
        self._validate_field_names_against_cache()
```

### 5.2 Layout pre-computation (Pod-β)

`PivotTable._compute_layout()` walks the cache's records to
pre-compute:
- `row_items`: ordered list of distinct row-axis tuples.
- `col_items`: ordered list of distinct col-axis tuples.
- `aggregated_values`: dict keyed by (row_tuple, col_tuple) →
  per-data-field aggregated value (sum, count, avg, etc.).

The aggregations populate `<rowItems>` / `<colItems>` children,
so Excel doesn't need to refresh on open. This is the core of
"Option A — full pivot construction".

### 5.3 Emit (Pod-α)

```rust
fn emit_pivot_table_xml(pt: &PivotTable, w: &mut Writer) -> Result<()> {
    write_xml_decl(w);
    open_root(w, "pivotTableDefinition", PIVOT_NS,
              pt.attrs_for_root());
    emit_location(w, &pt.location);
    emit_pivot_fields(w, &pt.cache_fields_with_axis_assignments());
    if !pt.row_field_indices.is_empty() {
        emit_row_fields(w, &pt.row_field_indices);
        emit_row_items(w, &pt.row_items);
    }
    if !pt.col_field_indices.is_empty() {
        emit_col_fields(w, &pt.col_field_indices);
        emit_col_items(w, &pt.col_items);
    }
    if !pt.page_fields.is_empty() {
        emit_page_fields(w, &pt.page_fields);
    }
    if !pt.data_fields.is_empty() {
        emit_data_fields(w, &pt.data_fields);
    }
    emit_style_info(w, &pt.style_info);
    close(w, "pivotTableDefinition");
    Ok(())
}
```

### 5.4 Patcher Phase 2.5m (Pod-γ)

Mirrors the cache emit pattern from RFC-047 §5.5; allocates a
`pivotTable{N}.xml` part-id, builds the table+cache rels file,
adds the sheet → table rel, and splices `<contentTypes>`.

Phase 2.5m order: caches first (RFC-047), then tables (RFC-048),
then chart pivot-source linkage (RFC-049). Tables are the only
piece that touches sheet rels (caches are workbook-scoped).

## 6. Cross-RFC: RFC-035 deep-clone

When `copy_worksheet(ws)` is called and `ws` has pivot tables,
RFC-035 §10 lifts the limit. The cloned sheet gets:
- A deep-clone of each pivot table part (each table is sheet-
  scoped; can't alias).
- An aliased reference to the same pivot cache (caches are
  workbook-scoped; deep-cloning would duplicate ~MB-scale
  records).
- Cell-range re-pointing on the table's `<location ref="…">`
  attribute, mirroring the chart cell-range rewrite from
  RFC-035 §5.4.

## 7. Verification Matrix

1. **Rust unit tests** (`cargo test -p wolfxl-pivot`):
   round-trip emit on synthetic 4-field pivot; row/col/page/data
   axis combinations; grand totals on/off; named-style picker.
2. **Golden round-trip**: `tests/diffwriter/test_pivot_table.py`.
3. **openpyxl parity**: construct via wolfxl, load via openpyxl,
   assert `wb.active._pivots[0].name`,
   `wb.active._pivots[0].rowFields`, etc., match.
4. **LibreOffice cross-renderer**: open
   `tests/fixtures/pivots/single_table.xlsx` produced by wolfxl;
   pivot grid renders region/quarter/sum-of-revenue populated.
5. **Cross-mode**: write-mode + modify-mode equivalence.
6. **Regression fixture**:
   `tests/fixtures/pivots/{single_table,multi_data_field,page_filter}.xlsx`.

## 8. Risks

| # | Risk | Likelihood | Impact | Mitigation |
|---|---|---|---|---|
| 1 | `<rowItems>` / `<colItems>` pre-aggregation diverges from cache records → Excel "data inconsistent" warning. | high | high | Pin a §10 contract that includes the full pre-aggregated layout (not just axis assignments); §10.6 enumerates `row_items` / `col_items` shapes verbatim. |
| 2 | 25 supported attrs vs 69 → user expectation gap. | med | low | RFC-048 §2.3 documents the supported set; KNOWN_GAPS lists deferred attrs. |
| 3 | Aggregator logic (sum / count / avg / min / max / countNums / stdDev / var) must match Excel's exactly. | med | high | Pod-β implements `_aggregate(records, aggregator)` against well-defined formulas; parity tests on each. |
| 4 | Multi-data-field layout (`dataOnRows="1"` vs default) is a UI-visible flip. | low | med | Default to `dataOnRows="0"` (Excel's default); only flip when explicit kwarg passed. |

## 9. Effort Breakdown

| Slice | Estimate | Notes |
|---|---|---|
| Pod-α: emit_pivot_table_xml + unit tests | 4d | Builds on `wolfxl-pivot` crate from RFC-047 |
| Pod-β: PivotTable + aggregator + layout pre-compute | 6d | The hard part is `_compute_layout()` aggregation |
| Pod-γ: Worksheet.add_pivot_table + Phase 2.5m + sheet rels | 3d | Mirrors chart pattern |
| Tests | 3d | |
| RFC-035 deep-clone | 2d | Pod-γ |
| **Total (parallel-pod)** | ~10 calendar days | |

## 10. Pivot-table-dict contract (Sprint Ν, v2.0.0)

> **Status**: Authoritative. Both Pod-α (Rust
> `parse::pivot_table_dict`) and Pod-β (Python
> `PivotTable.to_rust_dict`) MUST produce/consume this exact
> shape. Lesson #12: write the contract BEFORE pod dispatch.

### 10.1 Top-level keys

```python
{
    # Required
    "name": str,                              # "PivotTable1"
    "cache_id": int,                          # foreign key into cache definition
    "location": <location_dict>,              # see §10.2
    "pivot_fields": [<pivot_field_dict>, ...],  # one per cache field, IN CACHE FIELD ORDER

    # Axis enumerations — indices into pivot_fields[]
    "row_field_indices": [int, ...],
    "col_field_indices": [int, ...],
    "page_fields": [<page_field_dict>, ...],
    "data_fields": [<data_field_dict>, ...],

    # Pre-computed layouts — see §10.6
    "row_items": [<axis_item_dict>, ...],
    "col_items": [<axis_item_dict>, ...],

    # Layout
    "data_on_rows": bool,                     # default False
    "outline": bool,                          # default True
    "compact": bool,                          # default True
    "row_grand_totals": bool,                 # default True
    "col_grand_totals": bool,                 # default True

    # Captions (None → omit attr)
    "data_caption": str,                      # default "Values"
    "grand_total_caption": str | None,
    "error_caption": str | None,
    "missing_caption": str | None,

    # Apply-format flags (the 6 booleans)
    "apply_number_formats": bool,             # default False
    "apply_border_formats": bool,             # default False
    "apply_font_formats": bool,               # default False
    "apply_pattern_formats": bool,            # default False
    "apply_alignment_formats": bool,          # default False
    "apply_width_height_formats": bool,       # default True

    # Style picker (the named-style block)
    "style_info": <style_info_dict> | None,   # see §10.7

    # Versioning (Excel-rejected if not 6/3 — set by Pod-β to defaults)
    "created_version": int,                   # default 6
    "updated_version": int,                   # default 6
    "min_refreshable_version": int,           # default 3
}
```

### 10.2 `location` shape

```python
{
    "ref": str,                               # "F2:I20" — A1-style absolute
    "first_header_row": int,                  # default 0 (relative to ref's top)
    "first_data_row": int,                    # default 1
    "first_data_col": int,                    # default 1
    "row_page_count": int | None,             # for multi-page-field layouts
    "col_page_count": int | None,
}
```

### 10.3 `pivot_field` shape

One entry per cache field, in cache field order. Fields not on
any axis still appear (with `axis=None`). The Pod-β
`PivotTable._normalize_field_list` resolves a string field name
to the cache's field index via lookup.

```python
{
    "name": str | None,                       # override cache field name (rare)
    "axis": "axisRow" | "axisCol" | "axisPage" | "axisValues" | None,
    "data_field": bool,                       # default False; True → field is on dataFields
    "show_all": bool,                         # default False — `<pivotField showAll="0">`
    "default_subtotal": bool,                 # default True
    "sum_subtotal": bool,                     # default False
    "count_subtotal": bool,                   # default False
    "avg_subtotal": bool,                     # default False
    "max_subtotal": bool,                     # default False
    "min_subtotal": bool,                     # default False
    "items": [<pivot_item_dict>, ...] | None,  # see §10.4 — None → no <items> child
    "outline": bool,                          # default True
    "compact": bool,                          # default True
    "subtotal_top": bool,                     # default True
}
```

### 10.4 `pivot_item` shape

```python
{
    "x": int | None,                          # index into the cache field's shared_items.items
    "t": "default" | "sum" | "count" | "avg" | "max" | "min" | "blank" | "grand" | None,
    "h": bool,                                # hidden; default False
    "s": bool,                                # has format; default False
    "n": str | None,                          # display caption override
}
```

A `default`-typed item is the catch-all "(blank)" / total row.
`x=N` items reference `shared_items.items[N]` from the cache.

### 10.5 `data_field` shape

```python
{
    "name": str,                              # "Sum of revenue" — display name
    "field_index": int,                       # index into pivot_fields[]
    "function": "sum" | "count" | "average" | "max" | "min"
                | "product" | "countNums" | "stdDev" | "stdDevp"
                | "var" | "varp",
    "show_data_as": "normal" | "difference" | "percent" | "percentDiff"
                    | "runTotal" | "percentOfRow" | "percentOfCol"
                    | "percentOfTotal" | "index" | None,
    "base_field": int,                        # default 0; only used when show_data_as != normal
    "base_item": int,                         # default 0; ditto
    "num_fmt_id": int | None,                 # default None → inherit from cache field
}
```

### 10.6 `axis_item` shape (row_items / col_items)

```python
{
    "indices": [int, ...] | None,             # sequence of x values, one per axis field
    "t": "data" | "default" | "sum" | "count" | "avg" | "max" | "min"
         | "product" | "countNums" | "stdDev" | "stdDevp" | "var" | "varp"
         | "grand" | "blank" | None,
    "r": int | None,                          # repeat count (number of leading indices to repeat from previous item)
    "i": int | None,                          # data-field index when multi-data-field
}
```

`indices` is the path through the row/col axis fields — e.g. for
rows=[region, customer], a row entry might be `indices=[0, 3]`
(region=North, customer=Globex). The `t="grand"` row is the
total. `r` is OOXML's run-length compression for the leading
indices.

### 10.7 `style_info` shape

```python
{
    "name": str,                              # "PivotStyleLight16"
    "show_row_headers": bool,                 # default True
    "show_col_headers": bool,                 # default True
    "show_row_stripes": bool,                 # default False
    "show_col_stripes": bool,                 # default False
    "show_last_column": bool,                 # default True
}
```

The `name` is the Excel built-in style name; we don't (yet) emit
custom pivot styles.

### 10.8 `page_field` shape

```python
{
    "field_index": int,                       # index into pivot_fields[]
    "name": str | None,
    "item_index": int,                        # default 0; -1 for "(All)"
    "hier": int,                              # default -1 (no hierarchy)
    "cap": str | None,                        # display caption
}
```

### 10.9 Validation rules

* `cache_id` not registered → "PivotTable references unknown cache id"
* `row_field_indices` out of range → "row field index out of range"
* `data_fields` empty → "PivotTable requires ≥1 data field"
* `aggregator` not in §10.5 enum → "unknown aggregator function"
* Same field on multiple axes → "field {name} appears on multiple axes"
* `location.ref` overlaps source range → "pivot location overlaps source range" (warn, not error — Excel allows)

### 10.10 PyO3 boundary

```python
# wolfxl._rust
def serialize_pivot_table_dict(d: dict) -> bytes: ...
```

### 10.11 Versioning & legacy keys

This is the v2.0.0 introduction. No prior shape; nothing to
sunset. Future additions (calculated fields, filters, formats)
slot in via new optional keys at §10.1.

## 11. Out of Scope

Same list as RFC-047 §11, plus:
* **Pivot-table filters** (label/value filters). v2.1.
* **Per-cell formats** (`<formats>` collection). v2.1.
* **Calculated items**. v2.1.
* **`pivotHierarchies`** (OLAP). Out permanently.
* **`kpis`**. Out permanently.

## Acceptance

(Filled in after Shipped.)

- Commit: `<TBD: SHA>` — Sprint Ν Pod-α/β/γ merge
- Verification: `python scripts/verify_rfc.py --rfc 048` GREEN at `<TBD: SHA>`
- Date: <TBD>
