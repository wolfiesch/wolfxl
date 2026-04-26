# RFC-030: Structural — `Worksheet.insert_rows` / `delete_rows`

Status: Shipped
Owner: pod-030
Phase: 4
Estimate: L
Depends-on: RFC-012, RFC-021, RFC-022, RFC-024, RFC-025, RFC-026
Unblocks: RFC-031, RFC-034

## 1. Problem Statement

The two structural-row methods at
`python/wolfxl/_worksheet.py:1316-1340` raise `NotImplementedError`
today:

```python
def insert_rows(self, idx: int, amount: int = 1) -> None:
    raise NotImplementedError(
        "Worksheet.insert_rows is scheduled for WolfXL 1.1 (RFC-030). "
        ...
    )

def delete_rows(self, idx: int, amount: int = 1) -> None:
    raise NotImplementedError(
        "Worksheet.delete_rows is scheduled for WolfXL 1.1 (RFC-030). "
        ...
    )
```

User code that does the canonical openpyxl pattern hits the stub:

```python
wb = wolfxl.load_workbook("input.xlsx", modify=True)
ws = wb.active
ws.insert_rows(5, 3)   # NotImplementedError
ws.delete_rows(2)
wb.save("output.xlsx")
```

After this RFC, the same call shifts every coordinate-bearing
construct in the file: cell coordinates, formulas (already shipped via
`wolfxl_formula::shift`), hyperlink anchors, table refs and
autofilter, data-validation `sqref`, conditional-formatting `sqref`,
defined-name expressions, comment refs, and VML drawing anchor cells.
References that point INTO the deleted band become `#REF!` per OOXML
semantics. References that would shift past row `1_048_576` also
become `#REF!`. Out-of-range and invalid arguments (`idx < 1`,
`amount < 1`) raise `ValueError`. Empty pending queue is the no-op
identity path — workbook bytes are unchanged.

## 2. OOXML Spec Surface

Every element whose attribute (or text content) carries a row reference
that this RFC must rewrite. ECMA-376 Part 1 references in parens.

| Part / part-set | Element | Attribute / location | Notes |
|---|---|---|---|
| `xl/worksheets/sheet*.xml` | `<row>` | `r` | 1-based row index. |
| `xl/worksheets/sheet*.xml` | `<c>` | `r` | A1 ref like `B5`. |
| `xl/worksheets/sheet*.xml` | `<f>` | text content (formula) | Route through `wolfxl_formula::shift`. |
| `xl/worksheets/sheet*.xml` | `<dimension>` | `ref` | Sheet bounding range, e.g. `A1:E10`. |
| `xl/worksheets/sheet*.xml` | `<mergeCell>` | `ref` | Inside `<mergeCells>`. |
| `xl/worksheets/sheet*.xml` | `<hyperlink>` | `ref` | Inside `<hyperlinks>`. |
| `xl/worksheets/sheet*.xml` | `<dataValidation>` | `sqref` | Multi-range, space-separated. |
| `xl/worksheets/sheet*.xml` | `<conditionalFormatting>` | `sqref` | Multi-range, space-separated. |
| `xl/worksheets/sheet*.xml` | `<cfRule>/<formula>` | text content | Route through `wolfxl_formula::shift`. |
| `xl/worksheets/sheet*.xml` | `<dataValidation>/<formula1>` and `<formula2>` | text content | Route through `wolfxl_formula::shift`. |
| `xl/tables/table*.xml` | `<table>` | `ref` | Range. |
| `xl/tables/table*.xml` | `<autoFilter>` | `ref` | Range; matches table.ref's row span minus header? In our writer, equals table.ref. |
| `xl/workbook.xml` | `<definedName>` | text content | Route through `wolfxl_formula::shift`; honor `localSheetId` when shifting (only shift on the sheet being mutated). |
| `xl/comments*.xml` | `<comment>` | `ref` | Single-cell anchor. |
| `xl/drawings/vmlDrawing*.vml` | `<x:Anchor>` | text content (8 ints) | First two ints are `(col, row)` 0-based of the upper-left anchor — must shift the row index when row >= idx. The other six are sub-cell offsets / lower-right anchor. |

**Out-of-scope OOXML surface** (this RFC preserves the bytes verbatim
and may emit a warning):
- `<chartSpace>` (drawings) — chart-data refs must remain pointing at
  pre-shift cells if the chart data is supposed to slide; LibreOffice
  and Excel each do something different. Defer to RFC follow-up.
- `<pivotCacheDefinition>` and pivot tables.
- `<extLst>` extensions inside any sheet block.
- `INDIRECT(...)` formulas: detected via
  `TranslateResult::has_volatile_indirect`. We leave the formula text
  verbatim and (in a follow-up) surface a warning to the caller.
- Calculated-column table formulas: the `<calculatedColumnFormula>`
  inside a table column gets the shift treatment via formula
  translator. Plain table-column metadata (name, totalsRowLabel) is
  left alone.

**Schema-ordering constraint**: `CT_Worksheet` child order is
documented in RFC-011 (slot 15 = mergeCells, 17 = CF, 18 = DV, 19 =
hyperlinks, 31 = legacyDrawing, 37 = tableParts). RFC-030 does NOT add
new sibling blocks — it rewrites existing block contents in place — so
ordering is preserved naturally.

## 3. openpyxl Reference

`.venv/lib/python3.14/site-packages/openpyxl/worksheet/worksheet.py:717-738`:

```python
def insert_rows(self, idx, amount=1):
    """
    Insert row or rows before row==idx
    """
    self._move_cells(min_row=idx, offset=amount, row_or_col="row")
    self._current_row = self.max_row


def delete_rows(self, idx, amount=1):
    """
    Delete row or rows from row==idx
    """

    remainder = _gutter(idx, amount, self.max_row)

    self._move_cells(min_row=idx+amount, offset=-amount, row_or_col="row")

    # calculating min and max col is an expensive operation, do it only once
    min_col = self.min_column
    max_col = self.max_column + 1
    for row in remainder:
        for col in range(min_col, max_col):
            if (row, col) in self._cells:
                del self._cells[row, col]
    self._current_row = self.max_row
    if not self._cells:
        self._current_row = 0
```

`_move_cells` operates on the in-memory cell dict. openpyxl's cell
dict is the source of truth; on save, the workbook serializer walks
the dict. wolfxl is patcher-mode: there is no in-memory dict to
mutate — we rewrite the on-disk XML directly.

**openpyxl edge cases NOT replicated**:
- The `_current_row` bookkeeping is a writer-side cursor for the
  append-row API and has no analog in patcher mode.
- openpyxl's `_move_cells` does NOT touch formulas, hyperlinks,
  tables, defined names, or DV/CF refs. Users on openpyxl manually
  call `move_range(translate=True)` for formulas — wolfxl shifts
  them automatically. This is loud-divergence #6 from
  `Plans/rfcs/INDEX.md` (deferred opt-out flag).
- openpyxl's `_gutter` simply returns the rows that fall inside the
  deletion band so it can drop them. We implement the equivalent
  in Rust by passing the `DeletedRange` tombstone into the formula
  translator.

## 4. WolfXL Surface Area

### 4.1 Python coordinator

- `python/wolfxl/_worksheet.py:1316-1340` — replace the two
  `NotImplementedError` blocks with working bodies that:
  1. Validate `idx >= 1` and `amount >= 1`; raise `ValueError` otherwise.
  2. Append a tuple to `wb._pending_axis_shifts`:
     `(sheet_title, axis="row", idx, amount)` for insert,
     `(sheet_title, axis="row", idx, -amount)` for delete.
  3. Modify-mode and write-mode both queue. (Write-mode flush is
     out of scope — write-mode workbooks are constructed cell-by-cell
     and structural shifts of an empty workbook are a no-op; if anyone
     hits this in practice we add a follow-up.)

- `python/wolfxl/_workbook.py` — add `_pending_axis_shifts: list`
  initialization in three `__init__` paths, and add
  `_flush_pending_axis_shifts_to_patcher` that drains the list into
  `XlsxPatcher.queue_axis_shift(sheet, axis, idx, n)`. Hook into
  `save()` between Phase 2.5h (CF) and `_rust_patcher.save()`.

### 4.2 Patcher (modify mode)

- **NEW workspace crate** `crates/wolfxl-structural/` (no PyO3 — like
  `wolfxl-formula`). Modules:
  - `lib.rs` — re-exports.
  - `axis.rs` — `Axis::{Row, Col}` enum (already in
    `wolfxl-formula`, but we re-export ours so callers don't pull
    in the formula crate just to name the enum), the unified
    `ShiftPlan` and helper conversions.
  - `shift_cells.rs` — rewrites `<sheetData>` cell coords + `<row r="">`.
    Streams the XML using `quick-xml`, copies bytes verbatim except
    when patching a `<c r="">` or `<row r="">` attribute.
  - `shift_anchors.rs` — pure-string rewrite of `ref` / `sqref`
    attribute values. Covers single cell (`B5`), range (`A1:E10`),
    and multi-range (`A1:E10 G1:H5`). Returns `Some("#REF!")` when
    a single-cell anchor falls inside a delete band; for multi-range
    `sqref` we drop tombstoned ranges (matching openpyxl's `clip` of
    DV/CF).
  - `shift_formulas.rs` — thin wrapper around `wolfxl_formula::shift`
    plus a helper to construct a `ShiftPlan` from `(axis, idx, n)`
    with `respect_dollar=false` (see §10).
  - `shift_workbook.rs` — top-level orchestrator
    `apply_workbook_shift(zip_bytes_in, ops) -> SheetMutations` that
    walks every part and returns a `HashMap<String, Vec<u8>>` of
    rewritten files.

- `src/wolfxl/mod.rs` — new fields:
  - `queued_axis_shifts: Vec<AxisShift>` where
    `AxisShift { sheet: String, axis: AxisKind, idx: u32, n: i32 }`.
  - new PyMethod `queue_axis_shift(sheet, axis, idx, n)` that
    appends.
  - new Phase 2.5i flush block in `do_save` that iterates
    `queued_axis_shifts` in order and applies each shift to:
    - the per-sheet XML (reads from `file_patches` if already
      mutated by an earlier phase, else from the source ZIP);
    - all `xl/tables/table*.xml` parts touched by tables that live
      on the affected sheet;
    - `xl/workbook.xml` defined names (ones with no `localSheetId`,
      OR with `localSheetId` matching the affected sheet's index);
    - all `xl/comments*.xml` parts on the affected sheet;
    - the affected sheet's `xl/drawings/vmlDrawing*.vml` part.

- The flush block runs AFTER all other Phase-2.5 blocks so it picks
  up patcher-side rewrites (e.g. a hyperlink the user just queued
  on row 6 should also shift after `insert_rows(5, 3)`).

### 4.3 Native writer (write mode)

**No write-mode surface.** `Workbook()` write-mode workbooks are
constructed cell-by-cell with no pre-existing data; structural shift
of an empty workbook is a no-op. If a future caller needs it (after
RFC-035 copy_worksheet, perhaps), we revisit. Documented in §7.

## 5. Implementation Sketch

### 5.1 `crates/wolfxl-structural/` API

```rust
// axis.rs
pub enum Axis { Row, Col }
pub struct ShiftPlan { pub axis: Axis, pub idx: u32, pub n: i32 }

// shift_cells.rs
pub fn shift_sheet_cells(xml: &[u8], plan: &ShiftPlan) -> Vec<u8>;

// shift_anchors.rs
pub fn shift_anchor(s: &str, plan: &ShiftPlan) -> String;
pub fn shift_sqref(s: &str, plan: &ShiftPlan) -> String;

// shift_formulas.rs
pub fn shift_formula(formula: &str, plan: &ShiftPlan) -> String;
pub fn shift_formula_with_meta(formula: &str, plan: &ShiftPlan) -> wolfxl_formula::TranslateResult;

// shift_workbook.rs
pub struct AxisShiftOp {
    pub sheet: String,
    pub axis: Axis,
    pub idx: u32,
    pub n: i32,
}

pub struct SheetXmlInputs<'a> {
    /// All sheet XMLs in the workbook, keyed by sheet name (NOT path).
    pub sheets: BTreeMap<String, &'a [u8]>,
    /// Workbook XML bytes (for defined-name shifts).
    pub workbook_xml: &'a [u8],
    /// Per-sheet table parts (sheet name → vec of (path, bytes)).
    pub tables: BTreeMap<String, Vec<(String, &'a [u8])>>,
    /// Per-sheet comments part (sheet name → (path, bytes)).
    pub comments: BTreeMap<String, (String, &'a [u8])>,
    /// Per-sheet vmlDrawing part (sheet name → (path, bytes)).
    pub vml: BTreeMap<String, (String, &'a [u8])>,
    /// Sheet name → 0-based position (for definedName localSheetId).
    pub sheet_positions: BTreeMap<String, u32>,
}

pub struct WorkbookMutations {
    /// Path → new bytes (UTF-8 XML or VML bytes).
    pub file_patches: BTreeMap<String, Vec<u8>>,
}

pub fn apply_workbook_shift(
    inputs: SheetXmlInputs<'_>,
    ops: &[AxisShiftOp],
) -> WorkbookMutations;
```

### 5.2 Per-part shift table

For a `Row` shift `(idx=5, n=3)` on sheet `S1`:

| Part | What we rewrite |
|---|---|
| `S1` sheet XML — `<row r="5">` | `r` attribute (5 → 8 / `<row r="5">` deletion if n<0 and inside band). Children's `r` attrs are also rewritten. |
| `S1` sheet XML — `<c r="A5">` | `r` attribute via `shift_cells`. |
| `S1` sheet XML — `<f>` text | via `shift_formulas`. |
| `S1` sheet XML — `<dimension ref>` | via `shift_anchor` (range). |
| `S1` sheet XML — `<mergeCell ref>` | via `shift_anchor`. Tombstoned merges drop. |
| `S1` sheet XML — `<hyperlink ref>` | via `shift_anchor`. |
| `S1` sheet XML — `<dataValidation sqref>` + nested `<formula1>` / `<formula2>` | sqref via `shift_sqref`; formulas via `shift_formulas`. |
| `S1` sheet XML — `<conditionalFormatting sqref>` + nested `<cfRule>/<formula>` | sqref via `shift_sqref`; formulas via `shift_formulas`. |
| Tables on `S1` — `<table ref>` + `<autoFilter ref>` | both via `shift_anchor`. |
| Tables on `S1` — `<calculatedColumnFormula>` | via `shift_formulas`. |
| `xl/workbook.xml` — `<definedName>` | text content via `shift_formulas`. Scope rule: a name with `localSheetId="N"` only shifts when `N == sheet_positions[op.sheet]`; a name with no `localSheetId` (workbook-scope) shifts iff the formula references `S1` (the formula translator handles this naturally — refs to other sheets are left alone because the shift only applies to the sheet our `ShiftPlan` is anchored on; we use a `formula_sheet`-aware path). |
| `xl/comments*.xml` on `S1` — `<comment ref>` | via `shift_anchor`. |
| `xl/drawings/vmlDrawing*.vml` on `S1` — `<x:Anchor>` text | parse the 8 ints, rewrite int #1 (row index, 0-based) when `row + 1 >= idx`. Bound to MAX_ROW; out-of-range tombstones to "delete this shape" — we drop the shape from the VML in that case. |

### 5.3 No-op invariant

`apply_workbook_shift(inputs, &[])` returns an empty
`WorkbookMutations`. Callers that drain an empty queue MUST short-
circuit BEFORE calling into the structural crate. This matches every
other Phase-2.5 setter and ensures `wb.save()` after no edits is
byte-identical to the source.

### 5.4 `respect_dollar` decision

Per the BLOCKER in `Plans/rfcs/notes/excel-respect-dollar-check.md`,
the `ShiftPlan::respect_dollar` flag on `wolfxl_formula::shift` is
required-no-default precisely because the Excel-side verification
hasn't been done. RFC-030 picks **`respect_dollar=false`** at the
patcher boundary and documents the choice loudly:

- It matches the current spec text of `Plans/rfcs/INDEX.md` open
  question #3.
- It matches the "Excel insert/delete is a coordinate-space remap"
  semantic that we've been operating under.
- The verification result MAY require a 1-line patch to flip the
  flag once the human runs the 5-minute Excel test. Until then,
  `crates/wolfxl-structural/src/shift_formulas.rs` carries a
  top-of-file `// TODO(RFC-012 BLOCKER):` comment and §10 below
  enumerates the follow-up.

The decision is intentionally NOT exposed at the `queue_axis_shift`
or `Worksheet.insert_rows` PyMethod level: openpyxl signature parity
demands `(idx, amount=1)` on `insert_rows`, and adding an extra arg
would break drop-in usage. The single `respect_dollar=false` choice
is hard-coded in `crates/wolfxl-structural/src/shift_formulas.rs`.

### 5.5 Order of operations in Phase 2.5i

1. Drain `queued_axis_shifts` in append order.
2. For each op `(sheet, axis, idx, n)`:
   1. Resolve `sheet_path` from `self.sheet_paths`.
   2. Read sheet XML from `file_patches` if already mutated, else
      from the source ZIP.
   3. Read table parts attached to the sheet (from rels patches
      where mutated, else source).
   4. Read comments + vml parts (likewise).
   5. Read `xl/workbook.xml` once per save (cached).
   6. Build `SheetXmlInputs` and call
      `wolfxl_structural::apply_workbook_shift` with this single op.
   7. Merge the returned `file_patches` back into
      `self.file_patches`.
3. After all ops: write `xl/workbook.xml` back if any defined-name
   changes were emitted.

Multi-op sequencing matters: applying `insert(5, 3)` followed by
`delete(2, 1)` must produce the same result as Excel does (apply
each in source order; coordinates after the first op are in the new
coordinate space). We achieve this by iterating ops and re-reading
from `file_patches` between iterations.

## 6. Verification Matrix

| Layer | Coverage |
|---|---|
| 1. Rust unit tests (`-p wolfxl-structural`) | ≥30 tests covering: shift_anchor (single, range, multi-range, tombstone, MAX_ROW overflow); shift_sqref; shift_sheet_cells (cell + row attr + edge cases); shift_formula (insert + delete + INDIRECT detection); shift_workbook (full pipeline). |
| 2. Golden round-trip (diffwriter) | Best-effort; this RFC operates exclusively in modify mode. |
| 3. openpyxl parity | `tests/parity/` continues to pass (97/97 baseline). The structural ops are NOT in the parity suite (openpyxl semantics diverge — see Finding 6 in INDEX.md). |
| 4. LibreOffice cross-renderer | Manual: open RFC-030 fixture in LibreOffice, confirm formulas + hyperlinks remain pointing at the correct cells. |
| 5. Cross-mode | N/A — write-mode has no `insert_rows` semantics (constructing-from-empty). |
| 6. Regression fixture | `tests/test_axis_shift_modify.py` — ≥12 tests covering insert/delete in middle, at start, past last row; formulas pointing into and across the shift band; defined names; tables; hyperlinks; DV/CF; and the empty-queue no-op identity. |

The standardized "done" gate is
`python scripts/verify_rfc.py --rfc 030 --quick`.

## 7. Cross-Mode Asymmetries

**None — patcher-only.** Write mode has no `insert_rows`/`delete_rows`
because write-mode workbooks are constructed cell-by-cell starting
from an empty `<sheetData>`. Structural shift of an empty in-memory
sheet would only mutate the `_current_row` bookkeeping cursor on the
Python side, which has no analog yet (RFC-030 does not add it).

If a future RFC introduces the cell-dict bookkeeping (e.g. RFC-035
`copy_worksheet` may need it for splicing copied cells), the
write-mode path can be added then. Documented as a follow-up below.

## 8. Risks

| # | Risk | Likelihood | Impact | Mitigation |
|---|------|-----------|--------|-----------|
| 1 | `respect_dollar` default is wrong (Excel actually keeps `$`-refs put). | med | high | Hard-coded `false` per BLOCKER doc. 1-line flip if verification disagrees. TODO comment in `shift_formulas.rs`. |
| 2 | VML drawing anchor parsing breaks on non-Microsoft-emitted VML (e.g. produced by a non-Excel tool). | low | low | Drop the shape on parse failure rather than panic; surface a debug log. |
| 3 | A defined name with `localSheetId` referring to the wrong sheet's row gets shifted. | low | med | Scope check via `sheet_positions` lookup. Workbook-scope names are routed to the formula translator with `formula_sheet=Some(op.sheet)` so the per-sheet shift only applies to refs that target our sheet. |
| 4 | Multi-op sequencing produces wrong results because of coordinate-space drift. | med | med | Iterate ops one at a time; re-read XML from `file_patches` between iterations. Test coverage in `tests/test_axis_shift_modify.py::test_multi_op_sequence`. |
| 5 | `INDIRECT(...)` formulas don't shift even though the user expects them to. | med | low | Detected via `has_volatile_indirect`; verbatim pass-through; documented warning. |
| 6 | Pivot tables, charts contain row refs that we don't shift. | high | low | Documented in §10. The pivot/chart definitions remain pointing at the OLD coordinates; user-visible bug only when they re-open in Excel and the chart series go stale. Acceptable for 1.1; tracked as a follow-up. |
| 7 | Comment VML anchor misalignment after shift causes the comment popup to appear on the wrong row. | low | med | We rewrite the row index in the `<x:Anchor>` text; col index is untouched. Comments crate (RFC-023) computes anchor margins from col widths, not row heights — out of scope here. |

## 9. Effort Breakdown

| Slice | Estimate | Notes |
|---|---|---|
| Research | 1d | This document. |
| `wolfxl-structural` crate | 2-3d | axis + shift_cells + shift_anchors + shift_formulas + shift_workbook. ~30 unit tests. |
| Python wiring | ½d | `_worksheet.py` stub replacement + `_workbook.py` flush + `mod.rs` PyMethod + Phase 2.5i. |
| Tests (`tests/test_axis_shift_modify.py`) | 1d | ≥12 end-to-end tests. |
| Verification + iteration | 1d | Run `verify_rfc.py --rfc 030 --quick` until green. |

## 10. Out of Scope

- **Pivot tables** (`xl/pivotCache/`, `xl/pivotTables/`): preserved
  verbatim. Pivot row refs remain pointing at the OLD coordinates;
  re-refreshing in Excel is the user's recourse.
- **Charts** (`xl/charts/`): preserved verbatim. Chart-data ranges
  remain pointing at the OLD coordinates. Tracked as a follow-up.
- **`INDIRECT(...)` formulas**: preserved verbatim. We surface
  detection via `has_volatile_indirect`; warning-emission is a
  follow-up.
- **`OFFSET / ADDRESS / INDEX / CHOOSE / HYPERLINK` text-arg refs**:
  same treatment as `INDIRECT` — preserved verbatim.
- **External book references** (`[Book2.xlsx]Sheet1!A1`): pass-through
  (the formula translator's `RefKind::ExternalBook` is leave-alone).
- **Calculated-column table formulas**: actually IN SCOPE (the
  `<calculatedColumnFormula>` text inside table.xml is shifted via
  `shift_formulas`). Plain table-column metadata is NOT.
- **`respect_dollar` verification follow-up**: a 1-line flip in
  `crates/wolfxl-structural/src/shift_formulas.rs` if the BLOCKER
  doc resolves to "true". Tracked in
  `Plans/rfcs/notes/excel-respect-dollar-check.md`.
- **Write mode `insert_rows`**: see §7. Empty-workbook structural
  shifts are a no-op; if a future RFC needs them they're a follow-up.
- **`WOLFXL_STRUCTURAL_PARITY=openpyxl` env flag**: deferred to
  post-1.1 per INDEX open-question #6.

## Acceptance

- Commit `c4011d0` — `docs(rfc-030): recreate insert/delete-rows spec from template`
- Commit `da8f158` — `feat(rfc-030): add wolfxl-structural workspace crate`
- Commit `9e42342` — `feat(rfc-030): wire insert_rows / delete_rows end-to-end`
- Verification: `python scripts/verify_rfc.py --rfc 030 --quick` GREEN at `9e42342`
- Tests: `tests/test_axis_shift_modify.py` 16/16 pass; `tests/parity/` 97/97 pass
- Date: 2026-04-25
