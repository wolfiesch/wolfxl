# RFC-047: Pivot caches — `wolfxl.pivot.PivotCache` (definition + records emit)

Status: Approved (pre-dispatch contract spec authored; pods not yet dispatched)
Owner: Sprint Ν Pods α / β / γ
Phase: 5 (2.0)
Estimate: XL
Depends-on: RFC-010 (rels graph), RFC-013 (patcher infra: file_adds, content-types ops, two-phase flush), RFC-035 (copy_worksheet — pivot deep-clone lift, see §6), RFC-046 (chart construction — `Reference` is shared)
Unblocks: RFC-048 (pivot tables), RFC-049 (pivot charts), v2.0.0 launch

> **S** = ≤2 days; **M** = 3-5 days; **L** = 1-2 weeks; **XL** = 2+ weeks
> (calendar, with parallel subagent dispatch + review).

## 1. Background — Problem Statement

`python/wolfxl/pivot/__init__.py` is currently:

```python
"""Shim for ``openpyxl.pivot``."""
from wolfxl._compat import _make_stub
PivotTable = _make_stub(
    "PivotTable",
    "Pivot tables are preserved on modify-mode round-trip but cannot be constructed.",
)
__all__ = ["PivotTable"]
```

Modify-mode workbooks already preserve existing pivot caches and
pivot tables (the rels graph and content-types graph carry them
through; `crates/wolfxl-structural/src/sheet_copy.rs:498` aliases
the existing parts when copying a sheet). What's missing is the
**construction** path. Every code sample like:

```python
from wolfxl.pivot import PivotTable, PivotCache, Reference
src = Reference(ws, min_col=1, min_row=1, max_col=4, max_row=100)
cache = PivotCache(source=src)
pt = PivotTable(cache=cache, location="F2",
                rows=["region"], cols=["quarter"],
                data=[("revenue", "sum")])
ws.add_pivot_table(pt)
```

raises `NotImplementedError` immediately at the `PivotTable()`
constructor.

This RFC closes the **pivot cache** half: typed Python class
(`wolfxl.pivot.PivotCache`), per-class `to_rust_dict()` returning
the §10 contract shape, Rust crate (`wolfxl-pivot`) that consumes
the contract and emits `xl/pivotCache/pivotCacheDefinition{N}.xml`
+ `xl/pivotCache/pivotCacheRecords{N}.xml`, patcher integration
(`Workbook.add_pivot_cache(pc)`), and the cross-RFC bits in
RFC-035 (`copy_worksheet` deep-clone) and RFC-013 (content-types).

RFC-048 closes the **pivot table** half (table definition emit,
`Worksheet.add_pivot_table(pt)`).

The two RFCs are bundled into a single sprint because the cache
and table cross-reference (cache fields → table pivotFields by
index); splitting them across non-collaborating pods would
re-introduce the Sprint Μ contract-gap problem.

**Target behaviour**: `PivotCache(source=Reference(...))`
constructs without `NotImplementedError`. `wb.add_pivot_cache(pc)`
returns the cache id; the workbook's `xl/workbook.xml`
`<pivotCaches>` collection lists it; the cache is emitted as
both `pivotCacheDefinitionN.xml` (schema) and
`pivotCacheRecordsN.xml` (denormalised data snapshot). Excel
opens the workbook without "PivotTable references invalid data"
warnings; the pivot's data is populated **without requiring
refresh-on-open**. openpyxl reads the cache via
`load_workbook(...).pivots[0].cache` and gets a `CacheDefinition`
back.

## 2. OOXML Spec Surface

ECMA-376 Part 1 §18.10 (PivotCache).

### 2.1 Parts

| Path | Content type | Rels type | Purpose |
|---|---|---|---|
| `xl/pivotCache/pivotCacheDefinition{N}.xml` | `application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml` | `…/relationships/pivotCacheDefinition` | Schema + SharedItems |
| `xl/pivotCache/pivotCacheRecords{N}.xml` | `application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml` | `…/relationships/pivotCacheRecords` | Denormalised data snapshot |

The records part is owned by the cache definition (cache.rels
points to records). The definition is owned by the workbook
(`xl/_rels/workbook.xml.rels` adds a rel of type
`pivotCacheDefinition`; `xl/workbook.xml`'s `<pivotCaches>`
collection enumerates them with `cacheId` + `r:id`).

### 2.2 `pivotCacheDefinition.xml` skeleton

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       r:id="rId1"
                       refreshOnLoad="0"
                       refreshedBy="wolfxl"
                       refreshedDate="45000.0"
                       createdVersion="6"
                       refreshedVersion="6"
                       minRefreshableVersion="3"
                       recordCount="100">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:D100" sheet="Sheet1"/>
  </cacheSource>
  <cacheFields count="4">
    <cacheField name="region" numFmtId="0">
      <sharedItems count="4">
        <s v="North"/>
        <s v="South"/>
        <s v="East"/>
        <s v="West"/>
      </sharedItems>
    </cacheField>
    <cacheField name="revenue" numFmtId="0">
      <sharedItems containsSemiMixedTypes="0" containsString="0"
                   containsNumber="1" minValue="100" maxValue="9999"/>
    </cacheField>
    <!-- ... -->
  </cacheFields>
</pivotCacheDefinition>
```

### 2.3 `pivotCacheRecords.xml` skeleton

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    count="100">
  <r>
    <x v="0"/>            <!-- region: index into sharedItems -->
    <n v="2400"/>          <!-- revenue: inline numeric -->
    <x v="0"/>            <!-- quarter -->
    <s v="Acme"/>          <!-- customer: inline string (not shared) -->
  </r>
  <r>
    <x v="1"/><n v="3100"/><x v="1"/><s v="Globex"/>
  </r>
  <!-- ... 100 records -->
</pivotCacheRecords>
```

Per-cell child element kinds: `<x v="N"/>` (index into the
field's `sharedItems`), `<n v=N/>`, `<s v=…/>`, `<b v="0"/`,
`<d v="2026-01-15T00:00:00"/>`, `<m/>` (missing), `<e v="…"/>`
(error string).

### 2.4 Workbook wiring

`xl/workbook.xml` after first `<sheets>` and before any
`<definedNames>`:

```xml
<pivotCaches>
  <pivotCache cacheId="0" r:id="rId7"/>
  <pivotCache cacheId="1" r:id="rId8"/>
</pivotCaches>
```

`cacheId` is a 0-based index used as a foreign key by pivot
tables (`<pivotTableDefinition cacheId="0">`). `r:id` is a rels
id pointing at the cache-definition part.

## 3. openpyxl Reference

`openpyxl/pivot/cache.py` (~28 kB). Key classes:

* `CacheDefinition` (18 attrs) — top-level schema.
* `CacheField` — per-source-column.
* `CacheSource` + `WorksheetSource` — source range pointer.
* `SharedItems` — enumerated unique values per field.
* `Record` (in `record.py`) — single row; child element list.

What we copy: the public `__attrs__` shape, the `WorksheetSource`
shape, the SharedItems enumeration structure.

What we explicitly do NOT copy:
- `xs:` (external) cache sources — OLAP / ODBC. Out of scope (§10).
- `consolidation` cache sources — multi-range pivots. Out of scope.
- `tupleCache` / `pcdkpis` (PowerPivot KPIs). Out of scope.
- `groupItems` (date / range grouping). Out of scope; v2.1.
- openpyxl's lazy-record reading helpers (`Record.__call__`).
  Records are written, not read, by Sprint Ν.

## 4. WolfXL Surface Area

### 4.1 Python coordinator

* `python/wolfxl/pivot/__init__.py` — replace `_make_stub` with
  real exports: `PivotTable`, `PivotCache`, `Reference`,
  `PivotField`, `DataField`, `RowField`, `ColumnField`,
  `PageField`, `WorksheetSource`.
* `python/wolfxl/pivot/_cache.py` — `PivotCache`,
  `CacheField`, `WorksheetSource`, `SharedItems`. Per-class
  `to_rust_dict()` returning §10.
* `python/wolfxl/pivot/_table.py` — see RFC-048.
* `python/wolfxl/_workbook.py` — new
  `Workbook.add_pivot_cache(cache)` returning the allocated
  `cache_id`; new `_pending_pivot_cache_adds` queue.
* `python/wolfxl/_worksheet.py` — see RFC-048 for
  `add_pivot_table`.

### 4.2 Patcher (modify mode)

* `crates/wolfxl-pivot/` — new crate. Pod-α owner.
  * `model::PivotCache`, `model::CacheField`,
    `model::SharedItems`, `model::CacheRecord`,
    `model::WorksheetSource`, `model::CacheValue` (variant enum
    matching the `<x>` / `<n>` / `<s>` / `<b>` / `<d>` / `<m>` /
    `<e>` element kinds).
  * `emit::pivot_cache_definition_xml(pc: &PivotCache, w: &mut
    Writer)`.
  * `emit::pivot_cache_records_xml(pc: &PivotCache, w: &mut
    Writer)`.
  * `parse::pivot_cache_dict(dict: &PyDict) -> Result<PivotCache>`.
* `src/wolfxl/mod.rs` — Pod-γ adds:
  * PyO3 binding `wolfxl._rust.serialize_pivot_cache_dict(d) ->
    bytes` (definition).
  * PyO3 binding `wolfxl._rust.serialize_pivot_records_dict(d) ->
    bytes` (records).
  * `XlsxPatcher.queue_pivot_cache_add(cache_def_xml,
    cache_records_xml)`.
  * `XlsxPatcher` Phase 2.5m: drain queue, allocate fresh
    `pivotCache{N}.xml` numbers via `PartIdAllocator`, splice
    `<pivotCaches>` into `xl/workbook.xml`, add cache+records
    rels to `xl/_rels/workbook.xml.rels` and the cache's own
    rels.

### 4.3 Native writer (write mode)

`crates/wolfxl-writer/src/emit/pivot_cache.rs` (new) — wraps the
`wolfxl-pivot` crate's emit functions and handles ZIP entry
creation. Pod-α emits the parts; the writer's part-id allocator
hands out numbers. The writer's `Workbook` holds a
`Vec<PivotCache>` populated by Python `to_rust_dict` calls.

## 5. Implementation Sketch

### 5.1 `wolfxl-pivot` crate creation (Pod-α, day 1)

```toml
# crates/wolfxl-pivot/Cargo.toml
[package]
name = "wolfxl-pivot"
version = "0.1.0"
edition = "2021"

[dependencies]
quick-xml = { workspace = true }
wolfxl-rels = { path = "../wolfxl-rels" }
```

Add to root `Cargo.toml` `[workspace.members]` immediately after
`wolfxl-charts`. PyO3-free (mirrors `wolfxl-formula`,
`wolfxl-structural`).

### 5.2 Cache-definition emit (Pod-α, days 1-3)

Pseudocode for `emit::pivot_cache_definition_xml`:

```rust
fn emit(pc: &PivotCache, w: &mut Writer) -> Result<()> {
    write_xml_decl(w);
    open_root(w, "pivotCacheDefinition", PIVOT_NS, R_NS,
              [("r:id", &pc.records_rid),
               ("refreshOnLoad", "0"),
               ("refreshedBy", "wolfxl"),
               ("createdVersion", "6"),
               ("refreshedVersion", "6"),
               ("minRefreshableVersion", "3"),
               ("recordCount", &pc.records.len().to_string())]);

    open(w, "cacheSource", [("type", "worksheet")]);
    empty(w, "worksheetSource",
          [("ref", &pc.source.range), ("sheet", &pc.source.sheet)]);
    close(w, "cacheSource");

    open(w, "cacheFields", [("count", &pc.fields.len().to_string())]);
    for f in &pc.fields {
        emit_cache_field(w, f);
    }
    close(w, "cacheFields");

    close(w, "pivotCacheDefinition");
    Ok(())
}
```

`emit_cache_field` walks the field's `SharedItems`, choosing
inline `<s>`/`<n>`/`<d>`/`<b>` children for shared values, or
the `containsSemiMixedTypes`/`containsNumber`/`minValue`/
`maxValue` attribute pattern for numeric-only fields with no
shared enumeration.

### 5.3 Cache-records emit (Pod-α, days 3-5)

```rust
fn emit_records(pc: &PivotCache, w: &mut Writer) -> Result<()> {
    write_xml_decl(w);
    open_root(w, "pivotCacheRecords", PIVOT_NS,
              [("count", &pc.records.len().to_string())]);
    for record in &pc.records {
        open(w, "r", []);
        for cell in &record.cells {
            match cell {
                CacheValue::Index(i) => empty(w, "x", [("v", &i.to_string())]),
                CacheValue::Number(n) => empty(w, "n", [("v", &fmt_number(*n))]),
                CacheValue::String(s) => empty(w, "s", [("v", s)]),
                CacheValue::Bool(b) => empty(w, "b", [("v", if *b { "1" } else { "0" })]),
                CacheValue::Date(d) => empty(w, "d", [("v", &d.iso_format())]),
                CacheValue::Missing => empty(w, "m", []),
                CacheValue::Error(s) => empty(w, "e", [("v", s)]),
            }
        }
        close(w, "r");
    }
    close(w, "pivotCacheRecords");
    Ok(())
}
```

### 5.4 Python `PivotCache` (Pod-β, days 1-4)

```python
class PivotCache:
    def __init__(self, source: Reference, *, refresh_on_load: bool = False):
        if not isinstance(source, Reference):
            raise TypeError("PivotCache(source=...) must be a Reference")
        self.source = source
        self.refresh_on_load = refresh_on_load
        self._cache_id: int | None = None     # set by add_pivot_cache
        self._fields: list[CacheField] | None = None  # built from source
        self._records: list[list] | None = None       # ditto

    def _materialize(self, ws: "Worksheet") -> None:
        """Walk source range; build fields + records.

        Inferred per column:
          - all-numeric → numeric field with min/max (no shared items)
          - low-cardinality → shared items enumeration
          - mixed → containsSemiMixedTypes
          - dates → date field
        """
        # Walk ws[source.range], dispatch per-column type inference,
        # populate self._fields and self._records.

    def to_rust_dict(self) -> dict:
        if self._fields is None:
            raise RuntimeError("PivotCache._materialize not yet called")
        return {
            "cache_id": self._cache_id,
            "source": self.source._materialized_dict(),
            "fields": [f.to_rust_dict() for f in self._fields],
            "refresh_on_load": self.refresh_on_load,
        }

    def to_rust_records_dict(self) -> dict:
        return {
            "field_count": len(self._fields),
            "records": [
                [_cell_to_rust_dict(c) for c in row]
                for row in self._records
            ],
        }
```

Materialization happens lazily, when `add_pivot_cache(cache)` is
called against a workbook (the workbook holds the source's
worksheet, so we can resolve the source range to actual values).

### 5.5 Patcher integration (Pod-γ, days 5-9)

Phase 2.5m additions:

```python
# In XlsxPatcher.flush() Phase 2.5m, after charts:
for cache_def_xml, cache_records_xml in self._pivot_cache_adds:
    cache_n = self._part_ids.allocate("pivotCache")
    def_path = f"xl/pivotCache/pivotCacheDefinition{cache_n}.xml"
    rec_path = f"xl/pivotCache/pivotCacheRecords{cache_n}.xml"
    self._file_adds[def_path] = cache_def_xml
    self._file_adds[rec_path] = cache_records_xml

    cache_rels = RelsGraph::new()
    cache_rels.add(rt::PIVOT_CACHE_RECORDS,
                   f"pivotCacheRecords{cache_n}.xml",
                   TargetMode::Internal)
    self._file_adds[
        f"xl/pivotCache/_rels/pivotCacheDefinition{cache_n}.xml.rels"
    ] = cache_rels.serialize()

    workbook_rid = self._workbook_rels.add(
        rt::PIVOT_CACHE_DEF,
        f"pivotCache/pivotCacheDefinition{cache_n}.xml",
        TargetMode::Internal,
    )
    self._workbook_pivot_caches.push((self._next_cache_id, workbook_rid))
    self._next_cache_id += 1

# Splice <pivotCaches> into xl/workbook.xml
self._splice_pivot_caches()

# Add content types
self._content_types.add(def_path, CT_PIVOT_CACHE_DEFINITION)
self._content_types.add(rec_path, CT_PIVOT_CACHE_RECORDS)
```

### 5.6 Native writer integration (Pod-α, days 7-9)

`Workbook` holds `Vec<PivotCache>`. `Workbook::save_to_writer`
emits the cache definition + records before pivot tables (which
reference cache by id), and after charts. Part-id allocator
hands out cache numbers; rels graph patches in the workbook.xml
`<pivotCaches>` collection.

## 6. Cross-RFC: RFC-035 deep-clone extension

`copy_worksheet` of a sheet referenced by a pivot table currently
either errors or copies by-reference (the cache, definition, and
table parts are aliased — see `sheet_copy.rs:498`). For v2.0, we
keep aliasing at the cache level (one cache can serve N pivot
tables — same pattern as image-media reuse) but **deep-clone** the
pivot-table part (since each pivot table has its own location and
optional source-range re-pointing).

Cell-range re-pointing: when the source sheet of the pivot is
itself being copied, the cloned pivot table's
`<worksheetSource sheet="…">` attribute is rewritten to the new
sheet's name. This mirrors the chart's `<c:f>` rewrite from
RFC-035 §5.4.

The `sheet_copy.rs:498` `t if t == rt::PIVOT_TABLE` branch
becomes a deep-clone path; the existing `rt::PIVOT_CACHE_DEF`
branch keeps aliasing (caches are workbook-scoped).

## 7. Verification Matrix

1. **Rust unit tests** (`cargo test -p wolfxl-pivot`):
   round-trip of cache + records on synthetic 4-field × 100-record
   fixture; SharedItems numeric-only path; SharedItems string path;
   date path; mixed-type path.
2. **Golden round-trip** (`tests/diffwriter/test_pivot_cache.py`):
   write+read, byte-stable with `WOLFXL_TEST_EPOCH=0`.
3. **openpyxl parity** (`tests/parity/`): construct via wolfxl,
   load via openpyxl, assert `wb.pivots[0].cache.recordCount` and
   `wb.pivots[0].cache.cacheFields[0].name` match.
4. **LibreOffice cross-renderer**: open
   `tests/fixtures/pivots/single_cache.xlsx` produced by wolfxl;
   pivot table shows region/revenue/quarter populated. Manual.
5. **Cross-mode**: write-mode `Workbook().add_pivot_cache(pc)`
   produces equivalent bytes to modify-mode
   `load_workbook("blank.xlsx").add_pivot_cache(pc)`.
6. **Regression fixture**: `tests/fixtures/pivots/single_cache.xlsx`.

## 8. Risks

| # | Risk | Likelihood | Impact | Mitigation |
|---|---|---|---|---|
| 1 | `pivotCacheRecords` malformed → Excel "PivotTable references invalid data" warning. | med | high | Pin §10b records-dict shape from a known-working openpyxl fixture; require `recordCount` attr matches `<r>` element count. |
| 2 | `SharedItems` `containsSemiMixedTypes` flag confusion → Excel rejects. | med | high | Run §10.4 SharedItems-flags computation through openpyxl-emitted fixtures and assert equivalence. |
| 3 | Workbook.xml `<pivotCaches>` insertion order matters (must come after `<sheets>` and before `<definedNames>`). | low | med | RFC-013 splice helper already orders by canonical CT_Worksheet child grammar. |
| 4 | Date serialization in records — Excel prefers OLE-automation serial floats, not ISO strings. | med | med | Emit dates as `<x v="N"/>` indices into a date-typed `<sharedItems>` block, not inline. Avoids the serial-float ambiguity. |

## 9. Effort Breakdown

| Slice | Estimate | Notes |
|---|---|---|
| Pod-α: `wolfxl-pivot` crate + emit | 5d | Cache definition + records + unit tests |
| Pod-β: Python `PivotCache` + `CacheField` + `Reference` re-export + `to_rust_dict` | 4d | Materialization from source range; type inference |
| Pod-γ: Patcher Phase 2.5m + workbook.xml splice + content-types | 4d | Builds on RFC-013 + RFC-046 patcher precedent |
| Tests | 3d | Layers 1-3 + 5 + 6 |
| RFC-035 deep-clone extension | 2d | Pod-γ side; mirrors chart-deep-clone pattern |
| **Total (parallel-pod)** | ~10 calendar days | |

## 10. Pivot-cache-dict contract (Sprint Ν, v2.0.0)

> **Status**: Authoritative. Both Pod-α (Rust `parse::pivot_cache_dict`)
> and Pod-β (Python `PivotCache.to_rust_dict`) MUST produce/consume
> this exact shape. Lesson #12 from Sprint Μ-prime: write the
> contract BEFORE pod dispatch.

### 10.1 Top-level keys (cache definition)

```python
{
    # Required
    "cache_id": int,                          # 0-based; allocated by Workbook.add_pivot_cache
    "source": <worksheet_source_dict>,        # see §10.2
    "fields": [<cache_field_dict>, ...],      # see §10.3 — empty list raises ValueError

    # Optional (None or missing → Excel default)
    "refresh_on_load": bool,                  # default False
    "refreshed_version": int,                 # default 6
    "created_version": int,                   # default 6
    "min_refreshable_version": int,           # default 3
    "refreshed_by": str,                      # default "wolfxl"

    # Records reference (set by patcher; Pod-β leaves None)
    "records_part_path": str | None,
}
```

### 10.2 `worksheet_source` shape

```python
{
    "sheet": str,                             # "Sheet1"
    "ref": str,                               # "A1:D100" — A1-style, absolute or relative
    "name": str | None,                       # Named range alternative — sheet+ref OR name (mutually exclusive)
}
```

If `name` is set, the emitter writes
`<worksheetSource name="MyRange"/>` instead of
`<worksheetSource ref="…" sheet="…"/>`. Validation in
`PivotCache.__init__` enforces exclusivity.

### 10.3 `cache_field` shape

```python
{
    "name": str,                              # "region" — column header in source
    "num_fmt_id": int,                        # default 0
    "data_type": "string" | "number" | "date" | "bool" | "mixed",
    "shared_items": <shared_items_dict>,      # see §10.4
    "formula": str | None,                    # calculated field (v2.1 — pin to None for v2.0)
    "hierarchy": int | None,                  # OLAP — pin to None for v2.0
}
```

### 10.4 `shared_items` shape

```python
{
    "count": int | None,                      # number of items in `items`; None → suppress count attr
    "items": [<shared_value>, ...] | None,    # enumeration; None → no <sharedItems> children, only attrs

    # SharedItems flags (omit attr if False)
    "contains_blank": bool,                   # default False
    "contains_mixed_types": bool,             # default False
    "contains_semi_mixed_types": bool,        # default True for string fields
    "contains_string": bool,                  # default True for string fields
    "contains_number": bool,                  # default False
    "contains_integer": bool,                 # default False
    "contains_date": bool,                    # default False
    "contains_non_date": bool,                # default True for non-date fields
    "min_value": float | None,                # numeric-only fields
    "max_value": float | None,                # numeric-only fields
    "min_date": str | None,                   # ISO format, date-only fields
    "max_date": str | None,                   # ISO format, date-only fields
    "long_text": bool,                        # default False
}
```

A field with `items=None` and numeric attrs (`contains_number=True`,
`min_value`, `max_value`) emits the
"numeric-only no-enumeration" form:
`<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="100" maxValue="9999"/>`.

A field with `items=[...]` emits each item as a child:
`<s v="…"/>` for strings, `<n v="…"/>` for numbers, `<d v="…"/>`
for dates, `<b v="0|1"/>` for booleans, `<m/>` for missing.

### 10.5 `shared_value` shape

A `shared_value` is a tagged variant:

```python
{"kind": "string", "value": str}
{"kind": "number", "value": float}
{"kind": "boolean", "value": bool}
{"kind": "date", "value": str}              # ISO 8601, "2026-01-15T00:00:00"
{"kind": "missing"}
{"kind": "error", "value": str}              # "#REF!", "#NAME?", etc.
```

### 10.6 `cache_records` (records part — separate dict, separate PyO3 binding)

```python
{
    "field_count": int,                       # must equal len(fields) of the cache def
    "record_count": int,                      # must equal len(records)
    "records": [<record_row>, ...],
}
```

`record_row` is a list of `record_cell` dicts, one per cache
field, in the same order as `fields[]` from §10.1.

### 10.7 `record_cell` shape

```python
{"kind": "index", "value": int}              # index into the field's shared_items.items
{"kind": "number", "value": float}            # inline numeric (field doesn't enumerate)
{"kind": "string", "value": str}              # inline string
{"kind": "boolean", "value": bool}             # inline bool
{"kind": "date", "value": str}                # ISO 8601 inline
{"kind": "missing"}                            # <m/>
{"kind": "error", "value": str}                # <e v="…"/>
```

The Pod-β materializer chooses `index` vs inline based on the
field's `shared_items.items`. Indices are 0-based.

### 10.8 Validation rules (raise `ValueError` at constructor)

* `fields` empty → "PivotCache requires ≥1 source field"
* `source.sheet` empty AND `source.name` empty → "PivotCache.source requires sheet+ref or name"
* `source.sheet` non-empty AND `source.name` non-empty → "sheet+ref and name are mutually exclusive"
* Field name not unique → "duplicate cache field name"
* `shared_items.items[i].kind == "index"` → forbidden (only valid in records, not in cache field shared items)
* `cache_id` set explicitly by user → forbidden ("set by add_pivot_cache")

### 10.9 Default value derivation (Pod-β materializer)

When walking the source range:

| Column observation | `data_type` | shared_items strategy |
|---|---|---|
| All numeric, ≤200 unique | `number` | enumerate items |
| All numeric, >200 unique | `number` | min/max attrs only |
| All string, ≤2000 unique | `string` | enumerate items |
| All string, >2000 unique | `string` | min/max NOT applicable; just attrs |
| All ISO date | `date` | enumerate dates as `<d>` items |
| Mixed types | `mixed` | contains_semi_mixed_types=True; enumerate items |
| All `None` | `string` | contains_blank=True, count=0 |

The 200/2000 thresholds are tunable in
`python/wolfxl/pivot/_cache.py:_INFER_THRESHOLDS`; documented in
the docstring.

### 10.10 PyO3 boundary

```python
# wolfxl._rust (PyO3 module)
def serialize_pivot_cache_dict(d: dict) -> bytes: ...
def serialize_pivot_records_dict(d: dict) -> bytes: ...
```

Both helpers are internal-only (not exported under
`wolfxl.__all__`). The Python coordinator calls them once per
cache during patcher Phase 2.5m.

### 10.11 Versioning & legacy keys

This is the v2.0.0 introduction of the contract. There is no
prior shape to be backward-compatible with. The Sprint Μ-prime
lesson #13 ("legacy key sunset documentation") does not apply —
no sunset because nothing pre-existed.

If v2.1 adds calculated fields or GroupItems, those slot in via
new optional keys (`formula` is already reserved at §10.3 with a
v2.0 pin to `None`).

## 11. Out of Scope

* **OLAP / external pivot caches** (cacheSource type=`external`).
  Out permanently.
* **Consolidation** sources (multi-range pivots). Deferred.
* **Calculated fields** (`<calculatedField>`). v2.1.
* **GroupItems** (date / range grouping). v2.1.
* **Pivot-cache styling** beyond `numFmtId`. v2.1.
* **In-place edits** to existing pivot caches beyond
  `add_pivot_cache`. v2.2.
* **Slicer caches**. v2.1 (separate RFC).

## Acceptance

(Filled in after Shipped.)

- Commit: `<TBD: SHA>` — Sprint Ν Pod-α/β/γ merge
- Verification: `python scripts/verify_rfc.py --rfc 047` GREEN at `<TBD: SHA>`
- Date: <TBD>
