# RFC-034: Structural — `Worksheet.move_range`

Status: Shipped
Owner: pod-034
Phase: 4
Estimate: L
Depends-on: RFC-012, RFC-030, RFC-031
Unblocks: —

## 1. Problem Statement

`python/wolfxl/_worksheet.py:1422-1440` raises `NotImplementedError`
the moment a user calls `Worksheet.move_range(...)` on a modify-mode
workbook:

```python
def move_range(
    self,
    cell_range: Any,
    rows: int = 0,
    cols: int = 0,
    translate: bool = False,
) -> None:
    raise NotImplementedError(
        "Worksheet.move_range is scheduled for WolfXL 1.1 (RFC-034). "
        ...
    )
```

User code that does the canonical openpyxl pattern hits the stub:

```python
wb = wolfxl.load_workbook("input.xlsx", modify=True)
ws = wb.active
# Move the rectangular block C3:E10 down by 5 rows and right by 2.
ws.move_range("C3:E10", rows=5, cols=2)
wb.save("output.xlsx")
```

After this RFC, the same call physically relocates every cell whose
A1 coordinate falls inside the source rectangle to the destination
rectangle (`(r + rows, c + cols)`). Formulas inside the moved block
are paste-style translated: relative references shift by
`(rows, cols)`; absolute references (`$A$1`) DO NOT shift (i.e.
`respect_dollar=true`, in line with RFC-012 §5.5). Formulas
*outside* the moved block that reference cells *inside* the source
range stay pointing at the OLD coordinate by default, matching
openpyxl's `move_range(translate=False)`. When the caller passes
`translate=True`, those external formulas are rewritten so they
re-anchor to the new location of the moved cells.

Validation:
- `cell_range` must be a non-empty A1-style range string (`"C3:E10"`)
  or a `(min_col, min_row, max_col, max_row)` tuple — anything else
  raises `ValueError`.
- `rows` and `cols` must be `int`; non-int raises `TypeError`.
- The destination rectangle (`src + (rows, cols)`) must lie inside
  Excel's coordinate space (`1..=MAX_ROW` and `1..=MAX_COL` per
  axis, inclusive). Out-of-bounds raises `ValueError`.
- `rows == 0 and cols == 0` is a no-op (matches openpyxl's
  early-return at `worksheet.py:781`). Empty queue → byte-identical
  save (no-op identity from RFC-013 §8).

Empty queue invariant: a modify-mode workbook with no `move_range`
calls produces byte-identical `xl/worksheets/sheet*.xml`.
Multiple `move_range` calls on the same or different sheets are
applied in append order, mirroring the existing axis-shift queue.

## 2. OOXML Spec Surface

Every part the operation rewrites. ECMA-376 Part 1 references in
parens; column-by-column treatment for parity with RFC-030/031.

| Part / part-set | Element | Attribute / location | Treatment |
|---|---|---|---|
| `xl/worksheets/sheet*.xml` | `<c>` | `r` (cell A1 ref) | Cells in `src` are relocated to `r + (rows, cols)`. Cells already at the destination band are silently overwritten (matches openpyxl `worksheet.py:780`: "Existing cells will be overwritten"). |
| `xl/worksheets/sheet*.xml` | `<c>` | inline-string / formula / value | Carried with the cell; relocated as a unit. |
| `xl/worksheets/sheet*.xml` | `<f>` | text content (formula) | If the cell is inside `src`, the formula is paste-translated: `respect_dollar=true`, `(d_row, d_col) = (rows, cols)`. If `translate=True`, formulas in cells *outside* `src` whose tokens point INTO `src` are also re-anchored. |
| `xl/worksheets/sheet*.xml` | `<row>` | `r` | Affected only indirectly: rows whose entire content moves out get re-emitted by sorting cells to their new addresses. We perform a parse + re-emit of `<sheetData>` rather than an in-place attribute splice (the granular row-by-row rewrite is more brittle than a tree rebuild for this op — see §5.1). |
| `xl/worksheets/sheet*.xml` | `<dimension>` | `ref` | Recomputed from the post-move cell set (see §5.6). |
| `xl/worksheets/sheet*.xml` | `<mergeCell>` | `ref` | A merge entirely inside `src` shifts by `(rows, cols)` (its anchor stays inside the moved block). A merge that straddles the source boundary is left in place — the merge's anchor cell is outside `src`, so the visible merge does NOT move with the operation. (Diverges from openpyxl, which doesn't touch merges; we explicitly shift fully-contained merges so the user-visible block stays merged in its new location.) |
| `xl/worksheets/sheet*.xml` | `<hyperlink>` | `ref` | Single-cell anchors inside `src` shift; range anchors fully inside `src` shift. Anchors that straddle the boundary stay put. |
| `xl/worksheets/sheet*.xml` | `<dataValidation>` | `sqref` | Per-piece treatment: pieces fully inside `src` shift; pieces outside or straddling the boundary stay put. (Matches the merge rule above.) |
| `xl/worksheets/sheet*.xml` | `<conditionalFormatting>` | `sqref` | Same per-piece rule as `<dataValidation>`. |
| `xl/worksheets/sheet*.xml` | `<dataValidation>/<formula1>`, `<formula2>` | text content | If `translate=True`, route through `wolfxl_formula::move_range(respect_dollar=true)` so refs into `src` re-anchor to `dst`. With `translate=False`, leave verbatim (mirrors openpyxl's "Formulae and references will not be updated" docstring). |
| `xl/worksheets/sheet*.xml` | `<cfRule>/<formula>` | text content | Same as DV formula text. |

**Out-of-scope OOXML surface** (preserved verbatim):
- `<col>` span elements — `move_range` is not an axis op; column
  widths are unaffected by relocating cell content.
- `xl/tables/tableN.xml` — moving cells inside a table's `ref`
  range is undefined in OOXML (Excel forbids it via the UI). We
  preserve the table part verbatim and document the limitation.
  If a user moves a region overlapping a table they get whatever
  Excel infers on next open; we explicitly do NOT try to be
  clever here. Tracked as a follow-up if a user reports a
  practical break.
- `xl/drawings/drawing*.xml` — chart anchors, image anchors. Same
  rationale as RFC-030/031: chart anchors are read-only in current
  wolfxl and round-trip verbatim.
- `xl/comments*.xml` and VML drawings. Comments anchored on
  `src`-cells move with their cells (the comment `<comment ref>`
  + the VML anchor row both shift). Implementation-wise this
  reuses the per-cell shift helper from RFC-030 (`shift_comments_xml`,
  `shift_vml_xml`), but bounded to the source rectangle. **Deferred
  to follow-up** — the v1 slice does not touch comments. Documented
  as a known gap.
- `xl/workbook.xml` `<definedName>` — defined names that quote refs
  inside `src` are NOT rewritten by v1 even when `translate=True`.
  (`translate=True` only rewrites *worksheet*-local formula bodies.)
  Tracked in §10. The `wolfxl_formula::move_range` call site is
  already structured so a future revision can plumb workbook.xml
  through the same pipeline.

**Schema-ordering constraint**: `CT_Worksheet` child order is
documented in RFC-011 (slot 6 = sheetData, 15 = mergeCells, 17 = CF,
18 = DV, 19 = hyperlinks). RFC-034 rewrites contents in place at
the same slots; ordering is preserved naturally by the
`shift_sheet_cells`-style streaming splicer that the implementation
shares with RFC-030/031.

## 3. openpyxl Reference

`/opt/homebrew/lib/python3.14/site-packages/openpyxl/worksheet/worksheet.py:771-799`:

```python
def move_range(self, cell_range, rows=0, cols=0, translate=False):
    """
    Move a cell range by the number of rows and/or columns:
    down if rows > 0 and up if rows < 0
    right if cols > 0 and left if cols < 0
    Existing cells will be overwritten.
    Formulae and references will not be updated.
    """
    if isinstance(cell_range, str):
        cell_range = CellRange(cell_range)
    if not isinstance(cell_range, CellRange):
        raise ValueError("Only CellRange objects can be moved")
    if not rows and not cols:
        return

    down = rows > 0
    right = cols > 0

    if rows:
        cells = sorted(cell_range.rows, reverse=down)
    else:
        cells = sorted(cell_range.cols, reverse=right)

    for row, col in chain.from_iterable(cells):
        self._move_cell(row, col, rows, cols, translate)

    # rebase moved range
    cell_range.shift(row_shift=rows, col_shift=cols)
```

`_move_cell` (worksheet.py:801-816):

```python
def _move_cell(self, row, column, row_offset, col_offset, translate=False):
    cell = self._get_cell(row, column)
    new_row = cell.row + row_offset
    new_col = cell.column + col_offset
    self._cells[new_row, new_col] = cell
    del self._cells[(cell.row, cell.column)]
    cell.row = new_row
    cell.column = new_col
    if translate and cell.data_type == "f":
        t = Translator(cell.value, cell.coordinate)
        cell.value = t.translate_formula(row_delta=row_offset, col_delta=col_offset)
```

What we copy verbatim:

- The signature: `(cell_range, rows=0, cols=0, translate=False)`.
- The `(rows, cols) == (0, 0)` early-return (no-op, no error).
- Source-range parsing via the equivalent of `CellRange(cell_range)`
  — wolfxl uses `wolfxl.utils.cell.range_boundaries` which already
  validates A1 syntax.
- "Existing cells will be overwritten" — destination cells outside
  `src` that conflict with the relocation are silently dropped and
  replaced.
- The iteration ordering (sort rows reverse on `down`, forward on
  `up`; sort cols reverse on `right`, forward on `left`) is purely
  to avoid clobbering during in-place mutation of the dict — it is
  irrelevant to wolfxl which builds a fresh post-move cell list and
  re-emits the sheet from it.
- `translate=True` semantics: only formula bodies stored *in cells*
  are rewritten. Defined names, table column formulas, and
  workbook-scope formulas are NOT touched (openpyxl scope). v1
  mirrors this scope; §10 lists the gap.

What we do NOT copy:

- The `Translator(cell.value, cell.coordinate)` import path. wolfxl
  uses the `wolfxl_formula::move_range` helper which is already
  set up with `respect_dollar=true` (see `crates/wolfxl-formula/
  src/translate.rs:466-482`). The two semantics agree per RFC-012
  §5.5.
- The in-memory dict-mutation pattern (`self._cells[(r, c)] = cell;
  del self._cells[(old_r, old_c)]`). wolfxl is a streaming patcher;
  we rewrite the on-disk XML directly via a parse + re-emit pass.
- openpyxl does NOT propagate `translate=True` into hyperlink
  anchors, merge cells, DV/CF sqref, or comment refs. wolfxl
  explicitly DOES move merge / hyperlink / DV / CF anchors that are
  fully inside `src`, regardless of `translate` — these are
  coordinate-bearing structures, not formula text, and leaving them
  pointing at empty cells would silently corrupt the merge / DV
  rules. Loud-divergence note: the docstring on `move_range`
  documents this departure, mirroring INDEX-decision-#6's
  "loud-divergence over silent-divergence" stance.
- openpyxl's `cell_range.shift(...)` line at the end of `move_range`
  rebases the in-memory `CellRange` object so subsequent code can
  use the post-move range. wolfxl's API takes the range string
  by value and does not return a rebased range; callers can
  construct one themselves if they need to.

## 4. WolfXL Surface Area

### 4.1 Python coordinator

File: `python/wolfxl/_worksheet.py`

Replace the `NotImplementedError` block with a working body that:

1. Validates input types:
   - `cell_range`: must be a string or have a 4-tuple representation
     that `wolfxl.utils.cell.range_boundaries` accepts. Anything else
     → `ValueError("move_range: cell_range must be a range string ...")`.
   - `rows`, `cols`: must both be `int`; otherwise `TypeError`.
   - `translate`: must be `bool` (Python's looser truthiness is fine
     here — we coerce via `bool(translate)`).
2. Parses `cell_range` via `range_boundaries` to get
   `(min_col, min_row, max_col, max_row)`. Raises `ValueError` if any
   coordinate is `None` (whole-row / whole-col refs not supported by
   `move_range`).
3. Validates the destination rectangle is in-bounds: every corner of
   `(min_col + cols, min_row + rows, max_col + cols, max_row + rows)`
   must satisfy `1 <= row <= MAX_ROW` and `1 <= col <= MAX_COL`. Out-
   of-bounds → `ValueError("move_range: destination ... is out of bounds")`.
4. If `(rows, cols) == (0, 0)` → return immediately (no-op).
5. Appends a tuple to a NEW `wb._pending_range_moves` list:
   `(sheet_title, src_min_col, src_min_row, src_max_col, src_max_row,
   d_row, d_col, translate)`.

File: `python/wolfxl/_workbook.py`

- Initialise `self._pending_range_moves: list = []` in all three
  `__init__` paths (write / `_from_reader` / `_from_patcher`),
  alongside the existing `_pending_axis_shifts`.
- Add a flush helper:
  ```python
  def _flush_pending_range_moves_to_patcher(self) -> None:
      patcher = self._rust_patcher
      if patcher is None or not self._pending_range_moves:
          return
      for (sheet, c0, r0, c1, r1, d_row, d_col, translate) in self._pending_range_moves:
          patcher.queue_range_move(sheet, c0, r0, c1, r1, d_row, d_col, translate)
      self._pending_range_moves.clear()
  ```
- Call it from `save()` between `_flush_pending_axis_shifts_to_patcher`
  and `_rust_patcher.save(filename)`. Order: range-moves run AFTER
  axis-shifts so a user that does `insert_rows(2, 3)` then
  `move_range("C3:E10", rows=5)` sees the move applied to the
  post-shift coordinate space. This mirrors the openpyxl mental
  model (each call mutates the workbook in place).

### 4.2 Patcher (modify mode)

File: `crates/wolfxl-structural/src/range_move.rs` — NEW module.
Public API:

```rust
/// Plan for one paste-style range move (RFC-034).
///
/// `src_lo` and `src_hi` are 1-based inclusive corners of the source
/// rectangle: `src_lo = (min_row, min_col)`, `src_hi = (max_row,
/// max_col)`. `d_row` / `d_col` are signed offsets — positive shifts
/// down / right, negative shifts up / left. `translate` controls
/// whether *external* formulas (cells outside the source rectangle)
/// also rewrite their refs that point INTO the moved block.
pub struct RangeMovePlan {
    pub src_lo: (u32, u32),
    pub src_hi: (u32, u32),
    pub d_row: i32,
    pub d_col: i32,
    pub translate: bool,
}

/// Apply one range-move plan to a worksheet XML. Returns the new
/// bytes. No-op `(d_row, d_col) == (0, 0)` returns the input
/// verbatim. Out-of-bounds destinations are the caller's problem
/// (the Python layer validates).
pub fn apply_range_move(sheet_xml: &[u8], plan: &RangeMovePlan) -> Vec<u8>;
```

Re-exported from `crates/wolfxl-structural/src/lib.rs`:

```rust
pub mod range_move;
pub use range_move::{apply_range_move, RangeMovePlan};
```

`Cargo.toml` already depends on `wolfxl-formula` and `quick-xml`; no
new deps.

File: `src/wolfxl/mod.rs`

Add a new field on `XlsxPatcher` mirroring `queued_axis_shifts`:

```rust
queued_range_moves: Vec<RangeMove>,

#[derive(Debug, Clone)]
pub struct RangeMove {
    pub sheet: String,
    pub src_min_col: u32,
    pub src_min_row: u32,
    pub src_max_col: u32,
    pub src_max_row: u32,
    pub d_row: i32,
    pub d_col: i32,
    pub translate: bool,
}
```

Add a new PyMethod:

```rust
fn queue_range_move(
    &mut self,
    sheet: &str,
    src_min_col: u32,
    src_min_row: u32,
    src_max_col: u32,
    src_max_row: u32,
    d_row: i32,
    d_col: i32,
    translate: bool,
) -> PyResult<()> { ... }
```

Validation: `src_min_col >= 1`, `src_min_row >= 1`,
`src_min_col <= src_max_col`, `src_min_row <= src_max_row`. Append-
only. Python validates destination bounds before calling.

Add a no-op gate alongside the existing
`queued_axis_shifts.is_empty()` check at line 1021.

Add Phase-2.5j drain block after Phase-2.5i (axis shifts), before
the `drop(zip)` that precedes Phase 4 (zip rewrite):

```rust
// --- Phase 2.5j: Range moves (RFC-034) ---
if !self.queued_range_moves.is_empty() {
    self.apply_range_moves_phase(&mut file_patches, &mut zip)?;
}
```

The `apply_range_moves_phase` helper mirrors `apply_axis_shifts_phase`:
read each affected sheet XML (from `file_patches` if mutated, else
source ZIP), call `wolfxl_structural::apply_range_move`, fold the
new bytes back into `file_patches`. Each queued op runs on the
post-previous-op bytes (mirrors RFC-030's multi-op sequencing).

### 4.3 Native writer (write mode)

**No write-mode surface.** Same rationale as RFC-030/031 §4.3:
write-mode workbooks are constructed cell-by-cell with no pre-
existing data; structural moves of an empty workbook are a no-op.
A future RFC that needs in-memory bookkeeping (e.g. RFC-035
copy_worksheet) can lift this; until then write-mode `move_range`
on an in-memory workbook is undefined.

## 5. Implementation Sketch

### 5.1 Sheet-XML rewrite algorithm

The implementation does NOT use a streaming attribute splice (the
shape of RFC-030 `shift_sheet_cells`). Reason: a range-move can
relocate cells across `<row>` boundaries — a cell at `B3` may end
up at `B8`, requiring it to migrate from one `<row>` element to
another. A streaming splice that only patches `r` attributes would
leave the cell as the first child of the wrong row, producing
malformed OOXML.

Instead `apply_range_move` does a two-phase rewrite of `<sheetData>`:

1. **Parse**: walk `<row>` and `<c>` events into a flat
   `Vec<(row, col, BytesStart, child_events)>` keyed by
   1-based `(row, col)`. Carry every child event verbatim (`<v>`,
   `<f>`, inline string `<is>`, etc.). For a `<row>` element,
   capture its attributes (style, height, customHeight, hidden) so
   we can re-emit them on the post-move row.
2. **Plan**: for each captured cell:
   - If `(row, col)` is inside `src` (the rectangle defined by
     `src_lo..=src_hi`): relocate to `(row + d_row, col + d_col)`.
     The cell's `<f>` payload (if any) goes through
     `wolfxl_formula::move_range(formula, &src_range, &dst_range,
     respect_dollar=true)` — this is the paste-style rewrite that
     SHIFTS relative refs by `(d_row, d_col)` and LEAVES `$`-marked
     refs put.
   - Else (cell outside `src`):
     - If `translate=False`: leave the cell verbatim.
     - If `translate=True` AND the cell has an `<f>` payload:
       route the formula through
       `wolfxl_formula::move_range(formula, &src_range, &dst_range,
       respect_dollar=true)`. This rewrites only the *refs* that
       point into `src`; refs to other locations are left verbatim
       (the translator's `RefDelta::move_src` field handles the
       scoping — refs outside `src` don't match and pass through).
3. **Conflict handling**: relocated cells overwrite any existing
   cells at the destination. We dedup by post-move `(row, col)`
   keeping the moved-cell entry (matches openpyxl's "Existing cells
   will be overwritten" docstring). Source cells whose old slot is
   not overwritten by something else are dropped (the slot becomes
   blank).
4. **Re-emit**: group cells by their new row, sort cells inside
   each row by column, sort rows by row index, and serialise back
   into `<sheetData>`. Row-element attributes (style, height) carry
   over from whichever pre-move row contributed cells to the post-
   move row; on conflict we prefer the row whose `r` matched the
   post-move row in the source (i.e. preserve existing row
   styling). If the post-move row didn't exist before, synthesize
   a bare `<row r="N">`.

For everything OUTSIDE `<sheetData>` (mergeCells, hyperlinks, DV,
CF, dimension), use a streaming splice that mirrors
`shift_sheet_cells`:

- `<dimension ref="">`: recompute from the post-move cell set
  (§5.6).
- `<mergeCell ref="">`: per-merge check — if the merge is fully
  inside `src`, shift its endpoints by `(d_row, d_col)` (delegate
  to a small `shift_range_by` helper). If the merge straddles the
  boundary, leave it. If the merge is fully outside `src`, leave
  it.
- `<hyperlink ref="">`: same per-anchor rule. Single-cell anchors
  fully inside `src` shift; range anchors fully inside shift; else
  leave.
- `<dataValidation sqref>` and `<conditionalFormatting sqref>`:
  per-piece rule — split the sqref on whitespace; for each piece
  check fully-in-src / boundary-straddle / fully-out and emit the
  shifted or original piece. Empty result drops the parent
  element (matches the drop-when-empty semantics from
  `shift_sheet_cells`'s `rewrite_ref_attr`).
- `<dataValidation>/<formula1>`, `<formula2>`, and
  `<cfRule>/<formula>`: text content. With `translate=True`, route
  through `wolfxl_formula::move_range`. With `translate=False`,
  leave verbatim.

### 5.2 Formula handling inside the moved block

Per RFC-012 §5.5 and the `wolfxl_formula::move_range` signature at
`crates/wolfxl-formula/src/translate.rs:466-482`, the call is:

```rust
let new_formula = wolfxl_formula::move_range(
    &cell.formula_text,
    &Range {
        min_row: src_lo.0,
        max_row: src_hi.0,
        min_col: src_lo.1,
        max_col: src_hi.1,
    },
    &Range {
        min_row: src_lo.0 as i64 as i64 + d_row as i64,
        // ...
    },
    /* respect_dollar */ true,
);
```

`respect_dollar=true` is the paste-style flag. With it set:

- `=A1` (relative) → `=A6` after a 5-row down move.
- `=$A$1` (absolute) → `=$A$1` (NOT shifted).
- `=$A1` (mixed col-abs / row-rel) → `=$A6` (only the row-rel part
  shifts).
- `=A$1` → `=A$1` (col-rel shifts in cols only; here `cols=0` so
  unchanged).

The translator's `RefDelta::move_src` field is what actually drives
the rewrite — refs that fall inside `move_src` re-anchor by
`move_dst - src.min`. Refs outside `move_src` are left alone.

### 5.3 Formula handling outside the moved block (`translate=True`)

For cells OUTSIDE `src` whose formulas reference cells INSIDE
`src`, the `translate=True` flag rewrites those refs to point at
the new (post-move) location.

The `wolfxl_formula::move_range` helper handles both cases with the
same call: refs outside `src` pass through unchanged, refs inside
`src` get re-anchored. So calling it on every formula in the sheet
(under `translate=True`) is correct — cells inside `src` see their
own refs re-anchored, cells outside `src` see refs that point INTO
`src` re-anchored, and cells whose formulas don't touch `src` are
no-ops (the translator returns the input string unchanged when no
ref matches).

`translate=False` skips the call for cells outside `src`. Cells
inside `src` still get the call (a relocation always re-anchors
the formula's relative refs to keep it semantically equivalent at
its new location — matching openpyxl's `_move_cell(translate=…)`
which only differs from `translate=True` in whether the formula's
ABSOLUTE refs that point OUTSIDE the cell's own paste-anchor are
treated as paste-style or coord-remap).

> Wait — the openpyxl semantics are subtler. Read the docstring
> again: "Formulae and references will not be updated." That sentence
> applies to `translate=False`. With `translate=True`, the formula
> body of each MOVED cell is paste-translated by `(rows, cols)`, and
> formulas in cells OUTSIDE the moved block are NOT touched.

The wolfxl spec departs from openpyxl on this last point: when
`translate=True`, we ALSO rewrite external formulas pointing into
`src`. Loud-divergence rationale: a user who passes
`translate=True` is signalling "I want my formulas to follow my
data" — propagating into external formulas is the strict superset
of openpyxl's behaviour and matches the user's stated intent. The
docstring documents the divergence, with the openpyxl-only path
available via `translate=False`.

If a user wants the openpyxl-narrow `translate=True` semantics
(internal-only), they can do
`ws.move_range(src, rows, cols, translate=False)` and then walk
the moved cells manually. We ship the broader semantics by
default.

### 5.4 Merge cell handling

Three cases per merge:

1. **Fully inside `src`**: every corner is in `[src_lo, src_hi]`.
   Shift both corners by `(d_row, d_col)`. The merge moves with
   the block.
2. **Straddling the boundary**: at least one corner is inside `src`
   and at least one is outside. Leave the merge unchanged. The
   user's call has split the merge anchor from its movement; the
   safest behaviour is to keep the merge in its original location.
   (Excel's UI forbids this case; we don't try to repair it.)
3. **Fully outside `src`**: leave alone.

Implementation: `range_in_src(merge_range, src_range) -> bool` is
a 4-int comparison. If true, emit `shift_range_by(merge_range,
d_row, d_col)`; else emit verbatim.

### 5.5 Hyperlink / DV / CF anchor handling

Same 3-case rule as merges. Single-cell anchors collapse to "fully
inside" or "fully outside" (no straddle case).

For DV and CF, the `sqref` attribute is a multi-range list; we
apply the rule per piece. An empty post-move list drops the parent
element (matches `shift_sheet_cells`).

### 5.6 Dimension recompute

After the parse + re-emit, walk the post-move cell set and compute
`(min_row, min_col)` and `(max_row, max_col)`. Emit
`<dimension ref="A1:..."`/> at the start of `<worksheet>`. If the
post-move cell set is empty, emit `<dimension ref="A1"/>` (matches
the openpyxl convention for empty sheets).

The implementation reuses the dimension-recompute logic that the
patcher already runs after RFC-030/031 shifts, so we don't add new
machinery.

### 5.7 Multi-op sequencing

`queued_range_moves` drains in append order. Each op reads the
sheet XML via `get_bytes(file_patches, zip, sheet_path)` which
returns the most-recent rewrite (if any) or the source bytes.
After each op, the new bytes go back into `file_patches` so the
next op sees them.

### 5.8 No-op invariant

`apply_range_move(xml, &plan)` with `(d_row, d_col) == (0, 0)`
returns the input bytes verbatim. The Python layer also short-
circuits on the `(0, 0)` case before queueing, so this is belt-
and-braces.

## 6. Verification Matrix

Six layers per the template:

| Layer | Coverage |
|---|---|
| 1. Rust unit tests (`cargo test -p wolfxl-structural`) | ≥10 tests in `range_move.rs::tests`: simple in-bounds move (+rows/+cols), negative deltas, fully-inside merge shift, straddling-merge leave-alone, hyperlink anchor shift, DV sqref multi-piece per-piece treatment, formula inside src with `$` (must NOT shift), formula inside src with relative ref (must shift), `translate=True` external-formula rewrite, no-op identity. |
| 2. Golden round-trip (diffwriter) | Best-effort; this RFC operates exclusively in modify mode. |
| 3. openpyxl parity (`pytest tests/parity/`) | The 97/97 baseline must remain green. `move_range` is NOT in the parity sweep (openpyxl semantics diverge per §3); regression coverage is in the new RFC-034 file. |
| 4. LibreOffice cross-renderer | Manual: open RFC-034 fixture in LibreOffice, confirm the moved block renders at the destination, formulas evaluate correctly, merges still merge in the new location. |
| 5. Cross-mode | N/A — write-mode has no `move_range` semantics (§4.3). |
| 6. Regression fixture | `tests/test_move_range_modify.py` — ≥10 end-to-end tests (see RFC §3 plan). Mirrors `test_axis_shift_modify.py` shape. Fixture authoring uses `openpyxl.Workbook` to construct the source xlsx. |

The standardized "done" gate is
`python scripts/verify_rfc.py --rfc 034 --quick`.

## 7. Cross-Mode Asymmetries

**None — patcher-only.** Same rationale as RFC-030/031 §7. Write-
mode workbooks have no `<sheetData>` content to relocate at
construction time.

If a future RFC introduces a write-mode cell-dict (e.g. RFC-035
copy_worksheet's splice-target machinery), the write-mode path can
land then. Documented as a follow-up below.

## 8. Risks

| # | Risk | Likelihood | Impact | Mitigation |
|---|------|-----------|--------|-----------|
| 1 | Parse + re-emit of `<sheetData>` corrupts unfamiliar sheet shapes (e.g. shared strings, inline-string cells, formula array `<f t="array">`). | med | high | The implementation parses the `<row>` / `<c>` event stream and treats each `<c>` as an opaque BytesStart + child events. Inline strings (`<is><t>...</t></is>`), shared-string indices (`<v>`), and array formulas (`<f t="array" ref="...">`) are carried through verbatim. Test coverage in `range_move.rs::tests` includes inline-string and shared-string cases. |
| 2 | Multi-op sequencing produces wrong results when two `move_range` calls overlap. | low | med | Each op runs against the post-previous-op bytes (the same machinery RFC-030 uses). Test: `test_two_moves_compose` in `tests/test_move_range_modify.py`. |
| 3 | `translate=True` external-formula rewrite re-anchors refs that the user actually wanted left at the OLD coordinate. | med | low | The flag defaults to `False` (matches openpyxl). When the user explicitly passes `True`, the broader rewrite is the requested behaviour. Loud-divergence note in the docstring. |
| 4 | Merge cell that straddles the source boundary is left in place even though the user might expect it to clip. | low | low | Documented in §5.4. Excel's UI forbids creating this case in the first place; we preserve verbatim rather than guess. |
| 5 | Defined names that quote refs into `src` are NOT rewritten under `translate=True`. | low | low | Documented as a known gap in §10. Tracked as a follow-up. |
| 6 | Range-move on a sheet that contains a table whose `ref` overlaps the moved block leaves the table ref pointing at empty cells. | med | low | Documented in §10. Excel re-derives table dimensions on next open from cell content; the table will likely report fewer rows. Acceptable for v1; tracked as a follow-up if reported. |
| 7 | Move that lands at a destination overlapping the source rectangle (e.g. `move_range("A1:E5", rows=2)` overlaps with original A3:E5) silently overwrites cells inside the overlap before they've been read. | high | med | The implementation reads ALL cells into a flat vector before computing relocations, so overlap is safe — the source-cell read happens before the destination write. Test: `test_overlapping_move_preserves_cells`. |

## 9. Effort Breakdown

| Slice | Estimate | Notes |
|---|---|---|
| Research | ½d | This document. Heavy reuse of RFC-030 §1-3 reduced authoring time. |
| `range_move` module | 2-3d | Parse + re-emit `<sheetData>` is the only new logic; merge / hyperlink / DV / CF anchor handling is a thin wrapper over the existing `shift_anchor` plus an in-src predicate. |
| `wolfxl_formula::move_range` integration | ½d | Already shipped under RFC-012. Just call it. |
| Patcher PyMethod + Phase 2.5j | ½d | Mirror `apply_axis_shifts_phase` — straightforward. |
| Python wiring | ½d | `_worksheet.py` stub replacement + `_workbook.py` flush + queue init in three `__init__` paths. |
| Tests (`tests/test_move_range_modify.py`) | 1-2d | ≥10 end-to-end tests. |
| Verification + iteration | 1d | Run `verify_rfc.py --rfc 034 --quick` until green. |

Total: ~7 days. Matches L estimate from INDEX.md.

## 10. Out of Scope

- **Defined names that reference cells inside `src`**: under
  `translate=True`, the v1 slice does NOT rewrite
  `xl/workbook.xml`'s `<definedName>` text. Tracked as a follow-up
  (`Plans/followups/rfc-034-defined-names.md` if a user reports it).
- **Tables whose `ref` overlaps `src`**: the v1 slice does NOT
  rewrite `xl/tables/tableN.xml`. Excel re-derives table extent
  on next open; documented breakage acceptable.
- **Comments and VML drawings anchored on `src`-cells**: deferred
  to follow-up. The single-cell `<comment ref>` and the VML
  `<x:Anchor>` integers can be shifted using the existing
  `shift_comments_xml` and `shift_vml_xml` helpers — but bounded
  to a rectangle rather than an axis. Tracked as
  `rfc-034-comments-followup` if reported.
- **Pivot tables, charts**: same rationale as RFC-030/031.
- **`INDIRECT(...)` and other text-arg formulas**: preserved
  verbatim. Detection via `has_volatile_indirect`; warning-emission
  is a follow-up.
- **Write-mode `move_range`**: see §7. Empty-workbook range moves
  are a no-op; if a future RFC needs them, follow-up.
- **`WOLFXL_STRUCTURAL_PARITY=openpyxl` env flag**: deferred to
  post-1.1 per INDEX open-question #6.
- **`cell_range` as an `openpyxl.worksheet.cell_range.CellRange`
  object**: v1 accepts only A1 range strings or 4-tuples. If a
  user has a `CellRange` object they pass `str(cr)` — documented
  in the docstring.

## Acceptance

- Commit: see git log for `feat/rfc-034-move-range` branch.
- Verification: `python scripts/verify_rfc.py --rfc 034 --quick`
  GREEN.
- Tests: `tests/test_move_range_modify.py` 12/12 pass;
  `tests/parity/` 97/97 pass (unaffected); `tests/test_axis_shift_modify.py`
  + `tests/test_col_shift_modify.py` + `tests/test_move_sheet_modify.py`
  + `tests/test_structural_op_stubs.py` all green.
- Date: 2026-04-26.
