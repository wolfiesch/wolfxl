# RFC-049: Pivot-chart linkage — `chart.pivot_source = pt`

Status: Approved (pre-dispatch contract spec authored; pods not yet dispatched)
Owner: Sprint Ν Pod-δ
Phase: 5 (2.0)
Estimate: M
Depends-on: RFC-046 (chart construction), **RFC-048 (pivot tables)**
Unblocks: v2.0.0 launch

## 1. Background — Problem Statement

In openpyxl, attaching a chart to a pivot table is:

```python
chart.pivot_source = openpyxl.pivot.table.PivotSource(
    name="MyPivot", fmtId=0
)
```

This makes Excel render the chart as a **PivotChart** — its
series are bound to the pivot's row/col/data axes; Excel's chart
toolbar shows pivot-specific actions; right-click → "Refresh"
re-pulls from the pivot cache.

In the OOXML chart XML, this is a `<c:pivotSource>` block at the
top of `<c:chart>`:

```xml
<c:chart>
  <c:pivotSource>
    <c:name>MyPivot</c:name>
    <c:fmtId val="0"/>
  </c:pivotSource>
  <c:plotArea>
    …
  </c:plotArea>
</c:chart>
```

Today wolfxl has no `pivot_source` attribute on its chart
classes; setting it raises `AttributeError`. Sprint Ν adds it,
plus a tiny extension to the v1.6 chart-emit pipeline to
emit `<c:pivotSource>` when set.

## 2. OOXML Spec Surface

ECMA-376 Part 1 §21.2.2.158 (`pivotSource`). Two children:
* `<c:name>` — fully-qualified name, format
  `[<sheet name>!&]<table name>`. We emit just the table name
  if the chart is on the same sheet as the pivot.
* `<c:fmtId val="N"/>` — format id, 0..2^16-1. Default 0.

If `<c:pivotSource>` is present, every series must also carry a
`<c:fmtId val="0"/>` element matching the pivot source's fmtId.
This is enforced at emit time.

## 3. openpyxl Reference

`openpyxl/chart/pivot.py:PivotSource` (~30 LOC).

## 4. WolfXL Surface Area

### 4.1 Python

* `python/wolfxl/chart/_chart.py` — `ChartBase.pivot_source`
  attribute (default `None`); accepts a `PivotTable` instance
  or a `(name, fmt_id)` tuple.
* `python/wolfxl/chart/_chart.py:ChartBase.to_rust_dict` —
  add `"pivot_source"` key (see §10).

### 4.2 Rust emit

`crates/wolfxl-writer/src/emit/charts.rs` — extend the chart
emitter to emit `<c:pivotSource>` if `pivot_source` is set.
Touches the start of `emit_chart_root` (between the
`<c:chart>` open and the `<c:plotArea>` open).

`crates/wolfxl-charts/src/model.rs` — add
`Chart::pivot_source: Option<PivotSource>` field.
`PivotSource { name: String, fmt_id: u32 }`.

### 4.3 Per-series `fmtId`

When a chart has `pivot_source`, each `<c:ser>` must carry a
`<c:fmtId val="0"/>` element. Pod-δ extends
`emit_series` to inject this when `chart.pivot_source.is_some()`.

## 5. Implementation Sketch

### 5.1 Python `pivot_source` attribute

```python
# In ChartBase
def __init__(self, ...):
    ...
    self._pivot_source: PivotSource | None = None

@property
def pivot_source(self) -> "PivotSource | None":
    return self._pivot_source

@pivot_source.setter
def pivot_source(self, value: "PivotTable | tuple[str, int] | None") -> None:
    if value is None:
        self._pivot_source = None
        return
    if isinstance(value, PivotTable):
        # Resolve to (table name, default fmt_id=0)
        self._pivot_source = PivotSource(name=value.name, fmt_id=0)
        return
    if isinstance(value, tuple) and len(value) == 2:
        name, fmt_id = value
        self._pivot_source = PivotSource(name=name, fmt_id=fmt_id)
        return
    raise TypeError(
        "Chart.pivot_source accepts a PivotTable, (name, fmt_id) "
        "tuple, or None"
    )
```

### 5.2 `to_rust_dict` extension

The chart-dict shape from RFC-046 §10 already has a `pivot_source`
slot reserved (see §10 below; Pod-δ adds the dict shape and
documents it as a v2.0.0 addition).

### 5.3 Rust emit

```rust
fn emit_chart_root(c: &Chart, w: &mut Writer) -> Result<()> {
    open(w, "c:chart", []);

    if let Some(ps) = &c.pivot_source {
        open(w, "c:pivotSource", []);
        open_text(w, "c:name", &ps.name);
        empty(w, "c:fmtId", [("val", &ps.fmt_id.to_string())]);
        close(w, "c:pivotSource");
    }

    if let Some(t) = &c.title {
        emit_title(w, t);
    }
    emit_plot_area(w, &c.plot_area, c.pivot_source.is_some())?;
    if let Some(l) = &c.legend {
        emit_legend(w, l);
    }
    close(w, "c:chart");
    Ok(())
}
```

`emit_plot_area`'s second arg `is_pivot` causes
`emit_series` to inject `<c:fmtId val="0"/>` after the series
order.

## 6. Cross-RFC: RFC-035 deep-clone

When a sheet bearing a pivot chart is `copy_worksheet`'d, the
cloned chart's `pivot_source.name` is rewritten if the cloned
pivot table also got a new name (e.g. `MyPivot` → `MyPivot1`).
Mirrors the cell-range-rewrite pattern.

## 7. Verification Matrix

1. **Rust unit tests** (`cargo test -p wolfxl-charts`):
   `pivot_source` → emits `<c:pivotSource>` block; per-series
   `<c:fmtId>` injected.
2. **Golden round-trip**: `tests/diffwriter/test_pivot_chart.py`.
3. **openpyxl parity**: construct via wolfxl, load via openpyxl,
   assert `chart.pivotSource.name == "MyPivot"`.
4. **LibreOffice cross-renderer**: open
   `tests/fixtures/pivots/pivot_chart.xlsx`; chart renders as
   a true PivotChart (right-click shows pivot actions).
5. **Cross-mode**: identical bytes write+modify.
6. **Regression fixture**: `tests/fixtures/pivots/pivot_chart.xlsx`.

## 8. Risks

| # | Risk | Likelihood | Impact | Mitigation |
|---|---|---|---|---|
| 1 | Per-series `<c:fmtId>` placement is at end-of-series; misordering is silent (Excel reads OK but openpyxl complains on round-trip). | low | low | Pin order in `emit_series`: order-control before number-format before pivot-fmtId. |
| 2 | Pivot-source name with `!` (cross-sheet) confuses the chart's plot-area cache strings. | low | med | Pin emit: if chart and pivot are on the same sheet, emit just the name; otherwise prefix with the sheet name escaped. |
| 3 | `pivot_source` set but `pivot_table.cache.records` empty → Excel renders a blank chart. | low | low | Documented at §10.4 — empty cache is the user's responsibility, mirrors openpyxl. |

## 9. Effort Breakdown

| Slice | Estimate |
|---|---|
| Pod-δ: Python attribute + setter + dict | 1d |
| Pod-δ: Rust emit extension | 2d |
| Tests | 1d |
| **Total** | ~3-4 days |

## 10. Pivot-source-dict contract (Sprint Ν, v2.0.0)

> **Status**: Authoritative. Extension to RFC-046 §10.
> The chart-dict's existing `pivot_source` slot (RFC-046 §10.1
> reserved a key for this; Sprint Ν fills it in).

### 10.1 `pivot_source` shape (within chart_dict)

```python
{
    # Within ChartBase.to_rust_dict() output, at top level:
    "pivot_source": {
        "name": str,                          # "MyPivot" or "Sheet1!MyPivot"
        "fmt_id": int,                        # default 0; max 65535
    } | None,
}
```

`None` → no `<c:pivotSource>` block emitted; chart is a
standard chart.

### 10.2 Validation

* `name` must match the regex `^([A-Za-z_][A-Za-z0-9_]*!)?[A-Za-z_][A-Za-z0-9_ ]*$` —
  optional sheet prefix + table name.
* `fmt_id` must be in `[0, 65535]`.

### 10.3 PyO3 boundary

No new helper needed. The chart's existing
`serialize_chart_dict` (RFC-046 §10.10) handles `pivot_source`
as a top-level optional key. Pod-α (RFC-046) extends the Rust
parser to read this key; Pod-δ extends the emitter to emit it.

### 10.4 Empty-cache caveat

If the chart's `pivot_source` points at a `PivotTable` whose
`cache.records` is empty, the chart will render blank in Excel
(no series have data). This is the same behaviour as openpyxl
and is documented in `docs/migration/openpyxl-migration.md`.

### 10.5 Versioning

`pivot_source` is a v2.0.0 addition to the chart-dict contract.
Pre-2.0 chart dicts have no `pivot_source` key and behave as if
it were `None`. This is a non-breaking extension.

## 11. Out of Scope

* **Calculated chart series** referencing pivot calculated fields. v2.1.
* **Pivot-chart-specific filtering** (chart-level field filters
  separate from pivot-level). v2.1.
* **`<c:formatId>`** on individual data points. v2.1.
* **Cross-sheet pivot charts** with the chart on its own
  chartsheet. v2.1.

## Acceptance

(Filled in after Shipped.)

- Commit: `<TBD: SHA>` — Sprint Ν Pod-δ merge
- Verification: `python scripts/verify_rfc.py --rfc 049` GREEN at `<TBD: SHA>`
- Date: <TBD>
