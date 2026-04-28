# WolfXL 1.1 — Full openpyxl modify-mode parity for structural operations

WolfXL 1.1 closes the long-standing structural-ops gap with openpyxl.
After 1.1, every `Worksheet` / `Workbook` mutation an openpyxl program
typically issues against an existing xlsx — inserting / deleting rows,
inserting / deleting columns, moving a rectangular range, reordering
sheets, copying a sheet — has a byte-fidelity-preserving, pure-Rust
backend in modify mode. No more "load with openpyxl just to call
`ws.insert_rows`".

The 1.1 cycle is also a foundation cleanup: the unused
`rust_xlsxwriter` workspace dependency is finally gone, leaving the
native writer as the sole xlsx-write path it has been since Phase 2.

## What's new

### Row and column insert / delete (RFC-030, RFC-031, RFC-031 round 2)

```python
from wolfxl import load_workbook

wb = load_workbook("budget.xlsx", modify=True)
ws = wb["Sheet1"]

ws.insert_rows(5, amount=3)          # shift everything ≥ row 5 down by 3
ws.delete_rows(10)                   # delete row 10
ws.insert_cols("D", amount=2)        # accepts column letters or 1-based ints
ws.delete_cols(1, amount=4)

wb.save("budget.xlsx")
```

`Worksheet.insert_rows`, `delete_rows`, `insert_cols`, and
`delete_cols` rewrite the underlying OOXML in-place. Every cell
coordinate, `<dimension>` ref, merge anchor, hyperlink anchor, table
range, data-validation `sqref`, conditional-formatting `sqref`,
defined-name reference, and formula reference touched by the shift
band is remapped. Per OOXML semantics, a delete that consumes a
referenced cell emits `#REF!` tombstones in dependent formulas.

`insert_cols` / `delete_cols` accept either a 1-based integer or an
Excel column letter (`"D"`, `"AB"`).

A subtle point worth flagging for openpyxl migrants: insert / delete
runs with `respect_dollar=False` because the operation is
fundamentally a coordinate-space remap. If you have a formula
`=SUM($A$1:$A$10)` in a cell and you `insert_rows(5, amount=3)`, the
formula becomes `=SUM($A$1:$A$13)` — the absolute reference shifts
along with everything else, because the cells it points at have
moved. This matches openpyxl. The opposite convention (`respect_dollar=True`,
the paste-style behavior used by `move_range`) is documented in
*Known divergences* below.

`tableN.xml` is rewritten end-to-end — RFC-031 round 2 closed the
gap where `<tableColumns count="N">` and the `<tableColumn>` element
list stayed at the pre-shift size when an insert / delete overlapped a
table's column band. Pre-round-2 outputs were rejected by Excel and
openpyxl; 1.1 round-trips cleanly through both.

The per-column `<col min="..." max="..." width="..." style="...">`
metadata is preserved across column inserts and deletes via the
`<col>` span splitter (`crates/wolfxl-structural/src/cols.rs`) — width
and style information for each surviving column lands on the right
output column.

### Range moves (RFC-034)

```python
ws.move_range("B2:D10", rows=5, cols=2)              # paste-style move
ws.move_range("B2:D10", rows=5, cols=2, translate=True)  # also re-anchor outside refs
```

`Worksheet.move_range(cell_range, rows=0, cols=0, translate=False)`
relocates a rectangular block of cells. Formulas inside the moved
block are paste-translated using `respect_dollar=True`: relative
references shift by `(rows, cols)`, while `$`-anchored references
stay put. With `translate=True`, formulas in cells outside the moved
block that reference cells inside the source rectangle are re-anchored
to the new destination.

`MergeCells`, hyperlinks, and data-validation / conditional-format
`sqref` regions that lie fully inside the source rectangle are shifted
with the block; pieces straddling the source boundary or fully outside
are left in place. This matches openpyxl's documented behavior.

### Sheet operations (RFC-036, RFC-035)

```python
wb.move_sheet("Q4", offset=-2)                   # reorder tabs in-place
new_ws = wb.copy_worksheet(wb["Template"])       # modify-mode duplicate
new_ws = wb.copy_worksheet(wb["Template"], name="January")
```

`Workbook.move_sheet(sheet, offset)` reorders sheets in `<sheets>`
and updates every defined-name `localSheetId` that depended on tab
index.

`Workbook.copy_worksheet(source, name=None)` (modify-mode only) clones
an existing sheet with all of its cell data, formulas, styles,
merges, hyperlinks, tables, data validations, and conditional
formatting rules. The new sheet receives a unique sheet ID, a unique
relationship ID, and either an auto-numbered name (`"Sheet (2)"`) or
the caller's chosen name.

Copies are not perfect clones — see *Known divergences* below for the
remaining 1.1 limitations on images / drawings and openpyxl
re-saves of wolfxl-generated copies.

### Foundation: `rust_xlsxwriter` stripped (RFC-001)

The native writer (RFC W5 replacement) has been the sole xlsx-write
path since Phase 2. The `rust_xlsxwriter` workspace dependency
remained in `Cargo.toml` as dead weight; 1.1 removes it. No public
API change; build times shrink and the dependency tree no longer
advertises a code path that hasn't shipped a byte in months.

In real numbers: the workspace `cargo build --release` on a clean
target directory drops one transitive dep tree (`rust_xlsxwriter`
and its serde / regex / chrono pulls), trimming a measurable chunk
off cold-build wall time and shrinking the Cargo.lock surface that
downstream auditors have to scan. None of the wolfxl crates referenced
the symbol; the strip is purely manifest-level.

## Known divergences from openpyxl

| Surface | wolfxl 1.1 behavior | openpyxl behavior | Workaround |
|---|---|---|---|
| `copy_worksheet` re-saved by openpyxl | wolfxl writes a structurally-correct duplicate. If you then `openpyxl.load_workbook(...)` the file and re-save it, openpyxl drops the copy's tables, data validations, conditional formatting, and sheet-scoped defined names from the duplicate. | openpyxl's own `copy_worksheet` deep-copies these as Python objects, so its re-save preserves them. | Do not openpyxl-save in the middle of a wolfxl pipeline. Either keep the file in wolfxl until the final save, or accept the loss on the duplicate. The source sheet is unaffected either way. |
| `copy_worksheet` images / drawings | The copy's drawing part is aliased to the source's `xl/media/imageN.png`. Mutating the copy's image bytes mutates the source's. | openpyxl deep-copies image bytes per sheet. | wolfxl 1.1 has no `replace_image`-style API. Tracked for 1.2. If you need an isolated image, edit the source before copying. |
| `calcChain.xml` | Stale after structural ops. Excel rebuilds `calcChain.xml` on next open, so end users see correct calculation order. | openpyxl drops `calcChain.xml` entirely on re-save. | None needed for the Excel reader path. Programmatic readers that consume `calcChain.xml` directly will see incomplete data; have them ignore the file or recompute the chain themselves. |
| `respect_dollar` semantic split | insert / delete: `respect_dollar=False` (coordinate remap — `$A$1` shifts). move_range: `respect_dollar=True` (paste-style — `$A$1` does NOT shift). | openpyxl: same split in practice (insert/delete remaps absolute refs; `Worksheet.move_range` paste-translates). | None — the semantics match openpyxl. The split is documented here so library authors building on top of the wolfxl-formula crate know which mode to pick for their own ops. |

## Migration from 1.0

None required. Every new method is additive. Existing code that
calls `load_workbook`, `Workbook()`, `iter_rows`, `append`,
`write_rows`, `cell_records`, or any of the styling / formula APIs
continues to work with no source change. The `rust_xlsxwriter` dep
removal is invisible to library users — the binary cdylib is the
only consumer that ever saw the symbol.

If you previously worked around the missing structural ops by
opening the file with openpyxl just to call `insert_rows` / `move_range`
/ `move_sheet` / `copy_worksheet` and saving back through wolfxl,
you can now stay inside wolfxl modify mode end to end. The detour was
expensive both in latency and in fidelity (every openpyxl save round-
trips through the openpyxl object model and can drop preserved
parts) — eliminating it is the primary 1.1 win.

## Performance characteristics

Structural operations in 1.1 are pure-Rust XML rewrites, executed
directly against the serialized OOXML bytes. There is no roundtrip
through a Python model, no ZipFile read-and-rewrite of the entire
archive, and no openpyxl-style "rebuild every cell from scratch"
pass. Cost scales with the number of cells, anchors, and formulas
*touched by the shift band*, not with workbook size.

For the typical workload — insert a row near the top of a 100K-cell
sheet, save — the structural ops add a sub-millisecond bump on top
of the existing modify-mode flush. Larger ops (delete 50 columns
across a 5-million-cell sheet) stay linear in the band footprint
because the rewriter never scans cells outside the affected range.

`crates/wolfxl-structural` ships 99 tests covering coordinate shift,
anchor rewrite, formula re-emission, table column splitting, multi-
range `sqref` handling, comment / VML repositioning, and defined-name
re-targeting. The Python integration suite adds a further
~50 modify-mode end-to-end cases across the new methods. The
fuzz / property test for `apply_workbook_shift` (added in 1.1)
exercises 1000+ random `(axis, idx, n)` combinations against a
realistic fixture and asserts no panics + well-formed XML output.

A reasonable mental model for cost: each structural op is one pass
through the affected sheet's `<sheetData>` (cell-coord rewrite), one
pass through any tables on that sheet, one pass through comments /
VML if present, and one pass through the workbook-level
`<definedNames>` block. The XML reader is `quick_xml`'s streaming
parser, so memory stays bounded by the sheet's working-set rather
than the workbook's total size. There is no all-pairs scan, no
recursive descent over unrelated sheets, and no formula
re-evaluation — formula *strings* are rewritten in place; the
calculation cache (`<v>` values) is left untouched, matching the
Excel-on-open recalc convention.

## Acknowledgments

Sprint Δ + Sprint Ε pods that landed 1.1:

- **RFC-030 — `insert_rows` / `delete_rows`** — Phase-4a row pod.
- **RFC-031 — `insert_cols` / `delete_cols`** — Phase-4a column pod
  + Sprint Ε round-2 pod (`<tableColumns>` rewrite).
- **RFC-034 — `move_range`** — Sprint Ε Pod-B.
- **RFC-035 — `copy_worksheet` (modify mode)** — Sprint Ε Pods α/β/γ/δ
  (spec, implementation, integration, release).
- **RFC-036 — `move_sheet`** — Phase-4a sheet pod.
- **RFC-001 — `rust_xlsxwriter` strip** — Phase-4a foundation cleanup.
- **Sprint Ε Pod-ε (this release)** — release notes, changelog
  consolidation, fuzz expansion, README feature matrix update.

Specs: see `Plans/rfcs/030-insert-delete-rows.md`,
`Plans/rfcs/031-insert-delete-cols.md`,
`Plans/rfcs/034-move-range.md`,
`Plans/rfcs/035-copy-worksheet.md`,
`Plans/rfcs/036-move-sheet.md`,
`Plans/rfcs/000-template.md` for the RFC process itself.

The `respect_dollar` semantic split was settled in
`Plans/rfcs/notes/excel-respect-dollar-check.md`. The "two pods on
one shared crate need to sequence, not parallelize" sprint retro
note out of Phase-4a's RFC-031 reconciliation is captured in
`Plans/followups/rfc-030-031-api-coordination.md`.

Thanks to everyone who filed structural-ops issues in 1.0 — the
shape of 1.1's API is largely your design.
