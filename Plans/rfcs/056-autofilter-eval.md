# RFC-056 — AutoFilter conditions + filter evaluation

> **Status**: Approved
> **Phase**: 5 (2.0 — Sprint Ο)
> **Depends-on**: 011 (xml-merger), 013 (patcher extensions), 026 (CF — for `dxfId` allocation)
> **Unblocks**: 060
> **Pod**: 1B

## 1. Goal

Close the AutoFilter gap. Today `ws.auto_filter.ref = "A1:B10"`
works but no filter conditions are exposed. openpyxl exposes 11
filter classes plus sort-state. **User chose: evaluate filters
in patcher** (matches user expectations; openpyxl itself
round-trips XML only).

## 2. Public API

### 2.1 Filter classes

```python
class BlankFilter:
    """Filter blank/non-blank — matches openpyxl <blank/>."""

class ColorFilter:
    dxf_id: int
    cell_color: bool = True       # True=cell color, False=font color

class CustomFilter:
    operator: Literal["equal", "lessThan", "lessThanOrEqual",
                      "notEqual", "greaterThanOrEqual", "greaterThan"]
    val: str

class CustomFilters:
    filters: list[CustomFilter]
    and_: bool = False             # logical AND across filters

class DateGroupItem:
    year: int
    month: int | None = None
    day: int | None = None
    hour: int | None = None
    minute: int | None = None
    second: int | None = None
    date_time_grouping: Literal["year", "month", "day",
                                 "hour", "minute", "second"]

class DynamicFilter:
    type: Literal["null", "aboveAverage", "belowAverage", "tomorrow",
                  "today", "yesterday", "nextWeek", "thisWeek",
                  "lastWeek", "nextMonth", "thisMonth", "lastMonth",
                  "nextQuarter", "thisQuarter", "lastQuarter",
                  "nextYear", "thisYear", "lastYear", "yearToDate",
                  "Q1", "Q2", "Q3", "Q4",
                  "M1", "M2", ..., "M12"]
    val: float | None = None       # for above/belowAverage
    val_iso: str | None = None
    max_val_iso: str | None = None

class IconFilter:
    icon_set: str                   # "3Arrows", "5Quarters", etc.
    icon_id: int                    # 0-based index into the set

class NumberFilter:
    filters: list[float]            # values to keep
    blank: bool = False             # also include blanks
    calendar_type: str | None = None

class StringFilter:
    values: list[str]               # values to keep (logical OR)

class Top10:
    top: bool = True                # True=top, False=bottom
    percent: bool = False           # True=top N%, False=top N
    val: float                      # N value
    filter_val: float | None = None # actual threshold (read-only)

class FilterColumn:
    col_id: int                     # 0-based column index inside auto_filter.ref
    hidden_button: bool = False
    show_button: bool = True
    filter: BlankFilter | ColorFilter | CustomFilters | DynamicFilter \
          | IconFilter | NumberFilter | StringFilter | Top10 | None
    date_group_items: list[DateGroupItem] = []

class SortCondition:
    ref: str                        # range like "A2:A100"
    descending: bool = False
    sort_by: Literal["value", "cellColor", "fontColor", "icon"] = "value"
    custom_list: str | None = None
    dxf_id: int | None = None
    icon_set: str | None = None
    icon_id: int | None = None

class SortState:
    sort_conditions: list[SortCondition]
    column_sort: bool = False
    case_sensitive: bool = False
    ref: str | None = None

class AutoFilter:
    ref: str | None = None
    filter_columns: list[FilterColumn] = []
    sort_state: SortState | None = None

    def add_filter_column(self, col_id: int, filter, ...) -> FilterColumn: ...
    def add_sort_condition(self, ref: str, descending: bool = False,
                           sort_by: str = "value", **kw) -> SortCondition: ...
```

### 2.2 Worksheet integration

```python
ws.auto_filter.ref = "A1:D100"
ws.auto_filter.add_filter_column(
    col_id=0,
    filter=NumberFilter(filters=[100, 200, 300]),
)
ws.auto_filter.add_filter_column(
    col_id=2,
    filter=StringFilter(values=["red", "blue"]),
)
ws.auto_filter.add_sort_condition(ref="A2:A100", descending=True)
```

## 3. OOXML output

### 3.1 `<autoFilter>` block

```xml
<autoFilter ref="A1:D100">
  <filterColumn colId="0">
    <filters>
      <filter val="100"/>
      <filter val="200"/>
      <filter val="300"/>
    </filters>
  </filterColumn>
  <filterColumn colId="2">
    <filters>
      <filter val="red"/>
      <filter val="blue"/>
    </filters>
  </filterColumn>
  <sortState ref="A2:A100">
    <sortCondition ref="A2:A100" descending="1"/>
  </sortState>
</autoFilter>
```

### 3.2 `<row hidden="1">` markers (the evaluation result)

After filter evaluation determines which rows to hide, emit
`<row r="N" hidden="1">` markers on the failing rows. Existing
visible rows stay unchanged. Wolfxl-merger composes hidden-state
patches into existing `<row>` elements (idempotent: re-running
save with the same filter produces identical bytes).

## 4. Filter evaluation

New crate `crates/wolfxl-autofilter/` (PyO3-free, ~800 LOC):

```rust
pub fn evaluate(
    rows: &[Row],
    filter_columns: &[FilterColumn],
    sort_state: Option<&SortState>,
) -> EvaluationResult {
    EvaluationResult {
        hidden_row_indices: Vec<u32>,
        sort_order: Option<Vec<u32>>,
    }
}
```

### 4.1 Per-filter semantics

| Filter | Pass iff |
|---|---|
| `BlankFilter` | cell.is_empty() |
| `ColorFilter(dxf_id, cell_color)` | cell's dxf_id matches; cell_color==true → check fill, else check font |
| `CustomFilter(op, val)` | binary operator applied to cell.coerced_value vs val |
| `CustomFilters(filters, and_=False)` | logical AND if and_, else OR over filters |
| `DateGroupItem(...)` | date components match the given grouping precision |
| `DynamicFilter("aboveAverage", val)` | cell.numeric_value > val (val pre-computed by patcher = avg(column)) |
| `DynamicFilter("today")` | date == today() at save-time |
| `DynamicFilter("Q1")` etc | date.quarter() matches |
| `IconFilter(icon_set, icon_id)` | cell's CF icon_id matches |
| `NumberFilter(filters, blank=...)` | cell.numeric_value ∈ filters, OR cell.is_empty() if blank |
| `StringFilter(values)` | cell.string_value ∈ values (case-insensitive per Excel) |
| `Top10(top, percent, val)` | rank within column ≤ val (top) or > N-val (bottom); percent flag scales |

### 4.2 Multi-column logic

Logical AND across columns. A row passes iff every column
filter accepts it. Within a single `CustomFilters`, the
`and_` flag controls AND/OR.

### 4.3 Dynamic filter "today" reproducibility

`DynamicFilter("today")` etc evaluate at **save time**, not at
construction time. For deterministic-bytes mode
(`WOLFXL_TEST_EPOCH=0`), the evaluator uses
`UNIX_EPOCH + WOLFXL_TEST_EPOCH` as the reference date.

### 4.4 Sort evaluation

`SortState` produces a permutation of row indices. Pod 1B
applies the permutation by rewriting `<row r="N">` indices in
sheet XML — a heavy operation. **Decision**: SortState
**XML-only** in v2.0; physical row reordering is a v2.1
follow-up. Document this divergence in RFC-056 §8 and the
release notes.

## 5. Modify mode

XlsxPatcher Phase 2.5o (after pivots, before cells):

1. For every sheet with `auto_filter.filter_columns` non-empty,
   collect existing cell values within `auto_filter.ref`.
2. Run `evaluate(rows, filter_columns, sort_state)`.
3. Emit `<row hidden="1">` markers for `hidden_row_indices`.
4. Splice the new `<autoFilter>` block via wolfxl-merger.

Phase 2.5o runs BEFORE CF Phase (existing 2.5g) so a row hidden
by AutoFilter stays hidden regardless of CF, but CF's
`<x:dxf>` rules apply to the visible rows.

## 6. RFC-035 deep-clone

AutoFilter is sheet-scoped. Deep-clone the `<autoFilter>` block
into the cloned sheet's XML. Re-evaluate on the cloned sheet's
data (which may differ if the user mutated cells in the same
save).

## 7. Testing

- `tests/test_autofilter_filters.py` — XML round-trip (~30 tests covering all 11 classes).
- `tests/test_autofilter_evaluation.py` — per-class evaluation (~30 tests).
- `tests/parity/test_autofilter_parity.py` — openpyxl XML byte equality (~10 tests).
- `tests/diffwriter/test_autofilter_*.py` — byte-stable golden files (~5 tests).
- `tests/test_autofilter_modify.py` — modify-mode round-trip (~10 tests).
- `tests/test_autofilter_copy_worksheet.py` — RFC-035 deep-clone (~5 tests).

## 8. Documented divergences from openpyxl

- **Filter evaluation** runs at save-time. openpyxl emits
  filters as XML and Excel evaluates on open. Wolfxl evaluates
  AND emits the `<row hidden>` markers. Most users don't
  notice; some advanced users who re-open in pandas
  (no Excel involved) will appreciate it.
- **Sort state**: XML round-trip only in v2.0. Physical row
  reordering deferred to v2.1.
- **Dynamic filter "today"** uses `WOLFXL_TEST_EPOCH` if set
  (deterministic mode); otherwise system clock.

## 9. References

- ECMA-376 Part 1 §18.3.2 (CT_AutoFilter)
- ECMA-376 Part 1 §18.3.2.4 (CT_FilterColumn)
- ECMA-376 Part 1 §18.3.2.6 (CT_CustomFilters)
- openpyxl 3.1.x `openpyxl.worksheet.filters` source (reference impl).

## 10. Dict contract

`Worksheet.to_rust_autofilter_dict()`:

```python
{
    "ref": str | None,
    "filter_columns": [
        {
            "col_id": int,
            "hidden_button": bool,
            "show_button": bool,
            "filter": {
                "kind": "blank" | "color" | "custom" | "dynamic"
                       | "icon" | "number" | "string" | "top10",
                # plus per-kind fields:
                "operator": str | None,
                "val": float | str | None,
                "filters": list | None,
                "and_": bool | None,
                "type": str | None,
                "val_iso": str | None,
                "max_val_iso": str | None,
                "icon_set": str | None,
                "icon_id": int | None,
                "dxf_id": int | None,
                "cell_color": bool | None,
                "blank": bool | None,
                "calendar_type": str | None,
                "values": list[str] | None,
                "top": bool | None,
                "percent": bool | None,
            },
            "date_group_items": [
                {"year": int, "month": int|None, ...,
                 "date_time_grouping": str},
                ...
            ],
        },
        ...
    ],
    "sort_state": {
        "sort_conditions": [
            {"ref": str, "descending": bool, "sort_by": str,
             "custom_list": str|None, "dxf_id": int|None,
             "icon_set": str|None, "icon_id": int|None},
            ...
        ],
        "column_sort": bool,
        "case_sensitive": bool,
        "ref": str | None,
    } | None,
}
```

PyO3 bindings:
- `serialize_autofilter_dict(d) -> bytes` (XML emit)
- `evaluate_autofilter(d, rows) -> dict` (returns `{"hidden": [int], "sort_order": [int]|None}`)

## 11. Acceptance

- 11 filter classes constructible.
- `add_filter_column` + `add_sort_condition` work.
- Patcher Phase 2.5o evaluates and emits `<row hidden>`
  markers; verified by openpyxl round-trip seeing the right
  rows hidden.
- LibreOffice fixture renders the same hidden-row pattern.
- 90+ tests green.
