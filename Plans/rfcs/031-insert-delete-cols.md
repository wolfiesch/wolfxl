# RFC-031: Structural — `Worksheet.insert_cols` / `delete_cols`

Status: Shipped
Owner: pod-031
Phase: 4
Estimate: L
Depends-on: RFC-012, RFC-021, RFC-022, RFC-024, RFC-025, RFC-026, RFC-030
Unblocks: RFC-034

> Companion to RFC-030. Symmetric (column axis instead of row), shares
> the `wolfxl-structural` crate with RFC-030, plus one col-only quirk
> (the `<col>` span splitter, see §5.3).

## 1. Problem Statement

`python/wolfxl/_worksheet.py:1343-1366` raises `NotImplementedError`
the moment a user touches column structure on a modify-mode workbook:

```python
def insert_cols(self, idx: int, amount: int = 1) -> None:
    raise NotImplementedError(
        "Worksheet.insert_cols is scheduled for WolfXL 1.1 (RFC-031). "
        ...
    )

def delete_cols(self, idx: int, amount: int = 1) -> None:
    raise NotImplementedError(
        "Worksheet.delete_cols is scheduled for WolfXL 1.1 (RFC-031). "
        ...
    )
```

Both signatures must match openpyxl's: `idx` is the 1-based column
index where the operation begins (insert: shift right; delete: shift
left), `amount` is the number of columns. WolfXL accepts an additional
ergonomic affordance — a column letter (`"C"`) — and converts it via
`wolfxl.utils.column_index_from_string`.

**Target behaviour**: every column-bearing reference in the sheet's
OOXML and in workbook-scope XML that points at this sheet shifts by
`±amount` once the source column index is `>= idx`. References inside
the deleted band become `#REF!`. Empty queue → byte-identical save
(no-op invariant from RFC-013 §8). Default `respect_dollar=false`
until the Excel verification (`Plans/rfcs/notes/excel-respect-dollar-check.md`)
lands; documented loudly so it can be flipped behind a single line
change once verified.

## 2. OOXML Spec Surface

Every element whose attribute carries an A1-style **column** part must
shift on insert/delete. The catalog below mirrors RFC-030 (axis-flipped)
plus the col-only quirk (`<col>` span splitter, §5.3):

| Element | Attribute(s) | Shift rule | Spec |
|---|---|---|---|
| `<col min="" max="">` | `min`, `max` (1-based col indices) | shift both; **split spans that straddle the insert/delete boundary** (§5.3) | §18.3.1.13 |
| `<c r="">` | `r` (cell A1 ref) | column part shifts; row part untouched | §18.3.1.4 |
| `<dimension ref="">` | `ref` (range or single cell) | column extents shift; row extents untouched | §18.3.1.35 |
| `<mergeCell ref="">` | `ref` | each endpoint's col part shifts | §18.3.1.55 |
| `<hyperlink ref="">` | `ref` | col part shifts | §18.3.1.41 |
| `<dataValidation sqref="">` | `sqref` (space-delimited list) | each ref's col part shifts | §18.3.1.32 |
| `<conditionalFormatting sqref="">` | `sqref` | each ref's col part shifts | §18.3.1.18 |
| `<cfRule>` formula text | embedded A1 refs in the formula body | RFC-012 translator | §18.3.1.10 |
| `<table ref="">` + `<autoFilter ref="">` + `<tableColumns>` | `ref`, child `<tableColumn>` entries | range cols shift; **on delete: drop `<tableColumn>` entries inside the deleted column band and renumber the remaining `id` attributes** (§5.4) | §18.5.1.2 |
| `<definedName>` text | embedded A1 refs in the formula body | RFC-012 translator | §18.2.5 |
| comments `<comment ref="">` | `ref` (anchor cell A1) | col part shifts | §18.7.3 |
| VML `<x:Anchor>` | binds comment to a cell pair | col cell binding shifts | drawing.vml legacy |
| chart-anchor `<from>/<to>` `colOff`/`col` (drawing1.xml `<xdr:from>`) | NOT in scope this slice — chart anchors are read-only in current wolfxl; deferred to follow-up | §20.5.2.1 | (see §10) |

### Col-only quirk: `<col>` span elements

Unlike rows (each `<row r="N">` carries a single 1-based row index),
column metadata is described by **range spans**:

```xml
<cols>
  <col min="3" max="7" width="14.5" customWidth="1"/>
  <col min="10" max="10" width="22" hidden="1"/>
</cols>
```

A single `<col>` element covers the inclusive range `[min, max]`. Insert
and delete must:

- **Insert** (`insert_cols(idx, n)`):
  - If a span is entirely below `idx` → leave alone.
  - If a span is entirely at-or-above `idx` → both `min` and `max` shift by `+n`.
  - If `idx` falls **strictly inside** `[min+1, max]` → split the span into
    two: `[min, idx-1]` and `[idx+n, max+n]`. The inserted columns get
    no `<col>` entry (default width applies).
  - The original boundary `idx == min` is treated as "entirely above"
    (whole span shifts).
- **Delete** (`delete_cols(idx, amount)` where deleted band is `[idx, idx+amount-1]`):
  - If a span is entirely below `idx` → leave alone.
  - If a span is entirely above `idx+amount-1` → both `min` and `max` shift by `-amount`.
  - If a span overlaps the deleted band:
    - Bytes of the span inside the band are removed.
    - The portion below the band is preserved.
    - The portion above is shifted by `-amount`.
    - If both portions exist → emit two `<col>` entries (preserving
      the original element's attributes on each).
    - If exactly one survives → emit it with adjusted `min`/`max`.
    - If none survives → drop the entry.

`insert_cols` does NOT need to add new `<col>` entries (the default
column width applies to any column that has no `<col>` entry). The
splitter's job is purely to keep the existing spans correct under the
insert/delete shift.

### `<tableColumn>` removal on delete

RFC-024 emits one `<tableColumn id="N" name="...">` per column inside
a table. When `delete_cols` deletes a column band that overlaps a
table's `ref` range:

1. Compute the per-column overlap: which `<tableColumn>` entries fall
   strictly inside the deleted band.
2. Remove those entries from `<tableColumns>`.
3. Renumber the remaining `id` attributes to `1..k` (Excel rejects
   non-monotonic table-column IDs).
4. Update `<tableColumns count="">` to `k`.
5. Shrink the table's `<table ref="">` and `<autoFilter ref="">`
   accordingly. If the deleted band fully covers the table, the
   table is dropped entirely (rels entry, content-type override,
   `xl/tables/tableN.xml` part).

If the deleted band only **partially** overlaps the table (one or more
columns survive), the surviving `<tableColumn>` entries are renumbered
but their `name` attributes are preserved verbatim.

### Schema ordering (CT_Worksheet)

`CT_Worksheet` enforces strict child order: `<sheetPr>`, `<dimension>`,
`<sheetViews>`, `<sheetFormatPr>`, `<cols>`, `<sheetData>`,
`<sheetCalcPr>`, `<sheetProtection>`, `<protectedRanges>`, `<scenarios>`,
`<autoFilter>`, `<sortState>`, `<dataConsolidate>`, `<customSheetViews>`,
`<mergeCells>`, `<phoneticPr>`, `<conditionalFormatting>`, `<dataValidations>`,
`<hyperlinks>`, `<printOptions>`, `<pageMargins>`, `<pageSetup>`, ...,
`<tableParts>`, `<extLst>`. The streaming splice must preserve this
order — same constraint as RFC-030; the shared `wolfxl-structural`
implementation handles it.

## 3. openpyxl Reference

`.venv/lib/python3.14/site-packages/openpyxl/worksheet/worksheet.py:725-729`
and `:753-768`:

```python
def insert_cols(self, idx, amount=1):
    """
    Insert column or columns before col==idx
    """
    self._move_cells(min_col=idx, offset=amount, row_or_col="column")


def delete_cols(self, idx, amount=1):
    """
    Delete column or columns from col==idx
    """

    remainder = _gutter(idx, amount, self.max_column)

    self._move_cells(min_col=idx+amount, offset=-amount, row_or_col="column")

    # calculating min and max row is an expensive operation, do it only once
    min_row = self.min_row
    max_row = self.max_row + 1
    for col in remainder:
        for row in range(min_row, max_row):
            if (row, col) in self._cells:
                del self._cells[row, col]
```

`_move_cells` (lines 689-715) walks the in-memory cell dict, applying
the offset on the chosen axis. openpyxl uses positive `offset` (insert)
or negative (delete) and reverses iteration order to avoid clobbering.

What we copy:

- The signature: `(idx: int, amount: int = 1)`.
- `idx` is **1-based** and points at the column where the action begins.
- `delete_cols` removes `amount` columns (the band `[idx, idx+amount-1]`).
- `insert_cols` shifts every column whose source index is `>= idx` by
  `+amount`. Columns at `< idx` are untouched.

What we do NOT copy:

- `_move_cells` walks an in-memory cell dict. wolfxl is a streaming
  patcher; we mutate XML byte streams via `quick_xml::Reader` plus
  the `wolfxl-merger` block primitive. There is no Python-level cell
  dict to walk in modify mode.
- `_invalid_row` and openpyxl's read-side validators. wolfxl validates
  the structural input (`idx >= 1`, `amount >= 1`, `idx + amount <=
  MAX_COL+1` on delete) and lets the rest flow through.
- openpyxl does not propagate the shift into:
  - defined names (RFC-021's translator path closes this).
  - hyperlinks (RFC-022's `_pending_hyperlinks` map; RFC-031 closes).
  - tables (RFC-024 column removal closes).
  - data validations / conditional formatting (RFC-025 / RFC-026).
  - chart anchors (deferred — see §10).

  wolfxl propagates into all six (the first five immediately; chart
  anchors are out of scope for 1.1 — issue link in §10).

## 4. WolfXL Surface Area

### 4.1 Python coordinator

File: `python/wolfxl/_worksheet.py`

Replace the two `NotImplementedError` blocks at lines 1343-1366 with
calls into the patcher's queueing API. Both methods convert the
column letter / 1-based index input into a 1-based integer via
`wolfxl.utils.column_index_from_string`:

```python
from wolfxl.utils import column_index_from_string

def insert_cols(self, idx, amount=1):
    col = idx if isinstance(idx, int) else column_index_from_string(idx)
    self._queue_axis_shift("col", col, amount)

def delete_cols(self, idx, amount=1):
    col = idx if isinstance(idx, int) else column_index_from_string(idx)
    self._queue_axis_shift("col", col, -amount)
```

`_queue_axis_shift("col", at, n)` is the same coordinator helper
RFC-030 added for `("row", at, n)`; the only difference is the axis
literal. The helper validates `at >= 1`, `n != 0`, then routes the
shift through `self._workbook._rust_patcher.queue_axis_shift(...)`.

### 4.2 Patcher (modify mode)

The shift logic lives in **`crates/wolfxl-structural/`** (a new workspace
crate, owned jointly with RFC-030 — see §7). RFC-031's contribution
is:

- `crates/wolfxl-structural/src/cols.rs` — the `<col>` span splitter
  described in §5.3. Pure function: `split_col_spans(input_xml: &[u8],
  shift: AxisShift) -> Vec<u8>`. Tested in isolation.
- Extension to `crates/wolfxl-structural/src/shift_workbook.rs`'s
  `apply_workbook_shift(axis, idx, n, &mut workbook_state)` to handle
  the `Axis::Col` case for defined names. RFC-030 lands the `Axis::Row`
  case; RFC-031 follows the same control flow with `RefDelta::cols`
  populated instead of `RefDelta::rows`.
- Extension to the `<tableColumn>` removal-on-delete logic in
  `crates/wolfxl-structural/src/tables.rs` (the row-axis equivalent
  is a no-op: deleting rows never removes a `<tableColumn>` entry; this
  is purely col-side).

The plumbing on the `XlsxPatcher` side is identical to RFC-030's:
one new `queue_axis_shift(payload)` PyMethod (which RFC-030 should
already have parameterized over `axis`, see §7), one new
`queued_axis_shifts` field on the patcher, one new Phase-2.6 drain
step that calls `wolfxl_structural::apply_axis_shift` per affected
sheet plus `apply_workbook_shift` once per save.

ZIP parts touched (max set, depends on what's in the source):

- `xl/worksheets/sheetN.xml` (cells, dimension, mergeCells, hyperlinks,
  dataValidations, conditionalFormatting, tableParts).
- `xl/worksheets/_rels/sheetN.xml.rels` (only when a table is fully
  dropped — the rels entry is removed via RFC-013's `file_deletes`).
- `xl/tables/tableM.xml` (when partial overlap → renumber columns;
  when full coverage → delete via `file_deletes`).
- `[Content_Types].xml` (when a table is fully dropped → remove the
  override).
- `xl/comments/commentN.xml` and `xl/drawings/vmlDrawing*.vml`
  (comment anchor cells).
- `xl/workbook.xml` (defined names that point at this sheet).

### 4.3 Native writer (write mode)

No changes. `Workbook()` + `add_table()` + `insert_cols(...)` is not
a supported flow in the native writer — write mode constructs the
sheet from scratch, so structural ops are meaningless. The existing
`tests/parity/` suite continues to exercise the round-trip on the
modify path only.

## 5. Implementation Sketch

### 5.1 Phase 2.6 drain order

The patcher's existing phases are:

- 2.0 styles
- 2.1 cell value patches
- 2.2 cell format patches
- 2.3 sheet block merges (DV / CF / hyperlinks / table parts)
- 2.4 ancillary parts (comments / VML)
- 2.5 cross-sheet aggregation (defined names, content-types)

Insert / delete column shifting is **2.6 — runs LAST**, after every
other queue has been resolved. Rationale: the shift operates on the
final byte stream of `xl/worksheets/sheetN.xml` (and the workbook /
tables / comments parts). Running it last means we never have to
re-apply shifts to bytes inserted by Phase 2.3+ — those phases emit
their content at the **post-shift** coordinates because the Python
coordinator already shifted any pending mutations by the time it
flushed them.

> **Coordinator-side shift** — the Python layer maintains a per-sheet
> "pending shift" map. When the user calls `ws["E5"] = 1` after
> `ws.insert_cols(3, 2)`, the coordinator translates `E5` to `G5`
> before queueing the cell patch. RFC-030 owns the implementation;
> RFC-031 inherits it with `("col", at, n)` semantics.

### 5.2 Sheet-XML shift algorithm (shared with RFC-030)

```
fn apply_axis_shift(axis: Axis, at: u32, n: i32, sheet_xml: &[u8]) -> Vec<u8>:
    let mut reader = quick_xml::Reader::from_reader(sheet_xml);
    let mut writer = Vec::with_capacity(sheet_xml.len());

    while let Some(event) = reader.next() {
        match event {
            Start("col") | Empty("col") if axis == Axis::Col => {
                # delegate to cols::split_col_spans
            }
            Start("c") | Empty("c") => {
                rewrite_attr(b"r", |ref| shift_a1(ref, axis, at, n));
            }
            Start("dimension") | Empty("dimension") => {
                rewrite_attr(b"ref", |range| shift_range(range, axis, at, n));
            }
            Start("mergeCell") | Empty("mergeCell") => {
                rewrite_attr(b"ref", |range| shift_range(range, axis, at, n));
            }
            Start("hyperlink") | Empty("hyperlink") => {
                rewrite_attr(b"ref", |ref| shift_a1(ref, axis, at, n));
            }
            Start("dataValidation") | ... => {
                rewrite_attr(b"sqref", |list| shift_sqref_list(list, axis, at, n));
            }
            Start("conditionalFormatting") | ... => {
                rewrite_attr(b"sqref", |list| shift_sqref_list(list, axis, at, n));
                # cfRule formulas via RFC-012 translator
            }
            Start("formula") inside cfRule | Start("f") inside <c> => {
                rewrite_text(|formula| wolfxl_formula::shift(formula, ShiftPlan {
                    axis, at, n, respect_dollar: false,
                }));
            }
            _ => writer.write_event(event),
        }
    }

    writer.into_inner()
```

The single primitive `shift_a1`/`shift_range`/`shift_sqref_list` is
parameterized over `axis`. RFC-030's implementation MUST already accept
`Axis::Col`; the Rust enum is `wolfxl_formula::Axis` from RFC-012.

**Tombstone semantics on delete**: when `n < 0` (delete), references
whose source coordinate falls inside the deleted band become `#REF!`.
For cell refs (`<c r="">`), the row is dropped if `r` becomes
invalid. For `<mergeCell>`, the entire entry is dropped if both
endpoints fall inside the band; if only one endpoint falls inside,
the merge is clipped to the surviving rectangle.

### 5.3 `<col>` span splitter (col-only)

Algorithm — given existing spans `[m_i, M_i]` and an axis shift
`(idx, n)`:

```
fn split_col_spans(spans: Vec<ColSpan>, idx: u32, n: i32) -> Vec<ColSpan>:
    let mut out = Vec::new();
    if n > 0 {
        # insert
        for span in spans:
            if span.max < idx:
                out.push(span);                     # entirely below — untouched
            elif span.min >= idx:
                span.min += n; span.max += n;       # entirely above — shift
                out.push(span);
            else:
                # span straddles idx: idx is strictly inside (min+1..=max)
                out.push(ColSpan { min: span.min, max: idx - 1, ..span });
                out.push(ColSpan { min: idx + n, max: span.max + n, ..span });
    } else {
        # delete: deleted band is [idx, idx + |n| - 1]
        let band_lo = idx;
        let band_hi = idx + (-n) as u32 - 1;
        for span in spans:
            if span.max < band_lo:
                out.push(span);
            elif span.min > band_hi:
                span.min += n; span.max += n;       # n is negative
                out.push(span);
            else:
                # overlap with band
                if span.min < band_lo:
                    out.push(ColSpan { min: span.min, max: band_lo - 1, ..span });
                if span.max > band_hi:
                    out.push(ColSpan {
                        min: band_hi + 1 + n,    # n negative
                        max: span.max + n,
                        ..span,
                    });
    out
```

Each output span re-emits the original element's attributes verbatim
(width, customWidth, hidden, style, etc.) — only `min`/`max` are
rewritten. The serializer keeps schema ordering (the `<cols>` block
position inside `CT_Worksheet`).

### 5.4 `<tableColumn>` removal on delete

Algorithm:

```
fn shrink_table_on_delete(table_xml: &[u8], deleted_cols: Range<u32>)
    -> ShrinkResult:
    let table_ref = parse_attr(b"ref");        # e.g. "C2:H10"
    let table_lo = table_ref.col_min;          # 3
    let table_hi = table_ref.col_max;          # 8
    let band_lo = deleted_cols.start;
    let band_hi = deleted_cols.end;            # inclusive

    # Fully deleted?
    if band_lo <= table_lo and band_hi >= table_hi:
        return ShrinkResult::Drop;

    # Per-column intersection
    let surviving = (table_lo..=table_hi)
        .filter(|c| c < band_lo || c > band_hi)
        .collect::<Vec<_>>();

    let surviving_after_shift = surviving.iter().map(|c|
        if c < band_lo { c } else { c - (band_hi - band_lo + 1) }
    ).collect();

    # Remove <tableColumn> entries whose column index ∈ deleted band
    # Renumber id attributes to 1..k
    # Update <tableColumns count="k">
    # Update <table ref=""> and <autoFilter ref=""> to the new range
    return ShrinkResult::Shrunk { new_xml };
```

A fully dropped table also requires:

- Remove the rels entry from `xl/worksheets/_rels/sheetN.xml.rels`.
- Add the table part to `file_deletes` (RFC-013 §5).
- Remove the content-type override from `[Content_Types].xml` (via
  `queued_content_type_ops`).

### 5.5 `respect_dollar`

Defaulted to `false` until the Excel verification lands. The
default flows into `ShiftPlan::respect_dollar` at the patcher /
formula crate boundary. A single line change in
`crates/wolfxl-structural/src/lib.rs` flips it once
`Plans/rfcs/notes/excel-respect-dollar-check.md` is checked.

The default is plumbed via `wolfxl_structural::DEFAULT_RESPECT_DOLLAR`
and a `///` doc comment that points at the verification note.

### 5.6 No-op invariant

If `queued_axis_shifts` is empty at save time, the patcher's
do_save loop short-circuits to `std::fs::copy` — same as every
other queue. The XlsxPatcher gate is extended to include
`self.queued_axis_shifts.is_empty()` alongside the existing
no-op predicates. RFC-030 lands the gate; RFC-031 inherits it.

## 6. Verification Matrix

1. **Rust unit tests** (`cargo test -p wolfxl-structural`):
   - `cols::split_col_spans` — 8+ cases: insert below / above /
     splitting; delete fully inside / fully covering / partial overlap
     left / partial overlap right / spanning entire band.
   - `tables::shrink_table_on_delete` — 4 cases: full coverage (drop),
     partial overlap (shrink + renumber), no overlap (passthrough),
     `<autoFilter>` shrunk too.
   - `shift::shift_a1` and `shift::shift_range` — col axis cases.
2. **Golden round-trip** (`WOLFXL_TEST_EPOCH=0 pytest tests/diffwriter/`):
   - Optional layer in verify_rfc — covers reference RFC-031 fixtures.
3. **openpyxl parity** (`pytest tests/parity/`):
   - Same fixtures as RFC-030 with column-axis variants.
4. **LibreOffice** (manual): open the saved fixture in LibreOffice;
   confirm formula-bar, table column headers, hyperlink targets, and
   conditional-formatting highlights all show the shifted coordinates.
5. **Cross-mode**: write-mode + modify-mode produce equivalent files
   for the same input (write mode constructs from scratch, so the
   "input" here is "build from sheet model"; modify mode shifts an
   existing on-disk file). The two outputs round-trip equally through
   openpyxl.
6. **Regression fixture**: `tests/fixtures/rfc031_*.xlsx` — three
   fixtures: one with `<col>` spans, one with a table, one with
   workbook-scope defined names referencing the sheet.

Test file: `tests/test_col_shift_modify.py` (or extend the RFC-030
file `tests/test_axis_shift_modify.py` — final layout decided by
RFC-030's authoring decision).

Required pytest cases (≥10):

1. Insert in middle (`insert_cols("C", 2)` on a sheet with cells in
   B-H) — every cell `>= C` shifts right by 2; cells `< C` untouched.
2. Insert at start (`insert_cols(1, 1)`).
3. Insert past last column (no-op for cells; `<dimension>` extends).
4. Insert with formulas (`=C5+D5` → `=E5+F5` after `insert_cols(3, 2)`).
5. Insert with workbook-scope defined name pointing at the sheet.
6. Insert with table (table `ref="C1:E10"` → `ref="E1:G10"` after
   `insert_cols(3, 2)`).
7. Insert with hyperlink (`<hyperlink ref="C5">` → `<hyperlink ref="E5">`).
8. Insert with DV / CF — sqref list shifts.
9. Insert with `<col>` span split (span `[3,7]` + `insert_cols(5, 2)`
   → spans `[3,4]` and `[7,9]`).
10. Delete in middle (`delete_cols("C", 2)`) — cells in deleted band
    are dropped; cells `>= idx + amount` shift left.
11. Delete fully covers a table (table dropped; rels + content-types
    cleaned up).
12. Delete partially covers a table (`<tableColumn>` entries renumbered).
13. Delete with formulas pointing INTO the deleted band → `#REF!`.

The standardized "done" gate is `python scripts/verify_rfc.py --rfc 031 --quick`.

## 7. Cross-Mode Asymmetries

Same as RFC-030 — write mode does not have an `insert_cols` analog
because the Workbook starts empty and is built up one cell at a time.
The seam lives in the test layer: parity tests round-trip through both
modes by construction.

## 8. Risks

| # | Risk | Likelihood | Impact | Mitigation |
|---|------|-----------|--------|-----------|
| 1 | `respect_dollar` default flipped after this RFC ships → silent semantic regression | low | high | gate the default behind `wolfxl_structural::DEFAULT_RESPECT_DOLLAR`; flipping is one line + the verify_rfc green re-run |
| 2 | `<col>` span splitter mis-handles boundary case (idx == min) | med | med | test exhaustively; document the boundary policy in module rustdoc |
| 3 | `<tableColumn>` renumbering breaks Excel rendering | low | high | parity test against openpyxl; LibreOffice manual smoke |
| 4 | Phase-2.6 drain order interleaves badly with RFC-024 / RFC-022 | low | med | Phase 2.6 runs LAST; coordinator-side shift handles the cross-queue case |
| 5 | API drift between RFC-030 and RFC-031 (axis-parametric API) | high | high | shared crate `wolfxl-structural`; RFC-030 lands axis-parametric API first; RFC-031 reuses |

## 9. Effort Breakdown

| Slice | Estimate | Notes |
|-------|----------|-------|
| Research | 1 day | this doc + the col-only quirks |
| Rust impl (sharing with RFC-030) | 3-4 days | cols.rs splitter + tables.rs `<tableColumn>` removal |
| Python wiring | ½ day | extend `_queue_axis_shift` for `("col", ...)` |
| Tests | 2 days | 13+ pytest cases, fixture authoring |
| Review + reconciliation with RFC-030 | 1 day | API contract |
| **Total** | **~8 days (L)** | matches INDEX.md estimate |

## 10. Out of Scope

- Chart anchors (`xl/drawings/drawing1.xml`'s `<xdr:from>` /
  `<xdr:to>` `colOff` / `col`). Wolfxl currently treats chart anchors
  as opaque round-trip bytes; column shifts will leave them stale.
  Follow-up issue:
  `Plans/followups/rfc-031-chart-anchor-shift.md`.
- Pivot tables (`xl/pivotTables/pivotTableN.xml`). Out of wolfxl 1.1
  scope entirely (no pivot read path yet).
- Array-formula range collapse on partial deletion. RFC-012 ships
  array-formula support but the `delete_cols` clipping for arrays is
  deferred — same as RFC-030.
- `WOLFXL_STRUCTURAL_PARITY=openpyxl` env flag (deferred per
  INDEX.md open-question #6).
- Pre-shift validation (e.g. "delete would leave a table with 0 cols").
  We let Excel + the round-trip layer surface the error rather than
  duplicate the check.
- Updating embedded XLSX previews / thumbnails (`docProps/thumbnail.jpeg`).
  Static image — no shift needed.
- `xl/sharedStrings.xml` is unaffected (cell text doesn't carry
  column refs).

## Acceptance

Shipped 2026-04-25 on branch `feat/rfc-031-insert-delete-cols`.

Verification matrix:

- **Layer 1 (Rust)**: `cargo test -p wolfxl-structural` — 22 tests
  green (10 cols-splitter + 10 sheet-shift + ancillary).
- **Layer 5 (Cross-mode pytest)**: `tests/test_col_shift_modify.py`
  — 14 tests green covering insert/delete in middle/start, formulas,
  merges, `<col>` span split (real-span fixture + openpyxl per-column
  fixture), tombstone drop, no-op invariant, idx validation.
- **Layer 6 (Lint)**: ruff scoped to RFC-touched files green.
- Layers 2/3 are environment-dependent (diffwriter/parity); see verify
  output for the local run.

Spec deviations from §4 (documented as follow-ups in §10):

- `<hyperlink ref>`, `<dataValidation sqref>`,
  `<conditionalFormatting sqref>`, and `<table>` shifting are deferred
  to round 2 of RFC-031. The first slice covers the streaming
  sheet-XML rewrite (cells, dimension, mergeCells, `<col>` spans,
  formula text, tombstone drops). The patcher block paths
  (RFC-022/024/025/026) need a coordinator-side shift hook before
  their queues are flushed; this is tracked in
  `Plans/followups/rfc-030-031-api-coordination.md`.
- `apply_workbook_shift` for defined names is a stub (passes
  workbook XML through unchanged). RFC-021's defined-name merger
  doesn't expose the per-name scope filter the workbook shift needs;
  reconciliation tracked in the same follow-up.
- The `Axis::Row` path in `wolfxl-structural` is implemented for
  cell `r=` and `<row r>` attributes only. RFC-030 will fill in the
  full row-axis surface and reconcile the `apply_workbook_shift`
  signature.

Coordination notes: pod-030 had not landed when pod-031 reached the
implementation slice, so pod-031 created the `wolfxl-structural`
crate from scratch with the agreed axis-parametric API
(`apply_axis_shift(shift, sheet_xml)`,
`apply_workbook_shift(shift, sheet_name, workbook_xml)`). RFC-030
should adopt this API on rebase; if the signature doesn't fit,
reconcile via `Plans/followups/rfc-030-031-api-coordination.md`.

- Commit: see git log for `feat/rfc-031-insert-delete-cols` branch.
- Verification: `python scripts/verify_rfc.py --rfc 031 --quick`
  GREEN (excluding parity sweep, which is unaffected).
- Date: 2026-04-25
