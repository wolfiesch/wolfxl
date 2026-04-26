# RFC-036: `Workbook.move_sheet` (reorder)

Status: Researched
Owner: pod-036
Phase: 4
Estimate: M
Depends-on: RFC-021
Unblocks: —

## 1. Problem Statement

`wolfxl.Workbook.move_sheet` is currently a hard stub. The Python
coordinator at `python/wolfxl/_workbook.py:297-310` reads:

```python
def move_sheet(self, sheet: Worksheet | str, offset: int = 0) -> None:
    """Move *sheet* by *offset* positions within the sheet order.

    Tracked by RFC-036 (Phase 4 / WolfXL 1.1). See
    ``Plans/rfcs/036-move-sheet.md`` for the implementation plan.
    """
    raise NotImplementedError(
        "Workbook.move_sheet is scheduled for WolfXL 1.1 (RFC-036). "
        "See Plans/rfcs/036-move-sheet.md for the implementation plan. "
        "Workaround: use openpyxl for structural ops, then load the result "
        "with wolfxl.load_workbook() to do the heavy reads."
    )
```

The openpyxl signature it shadows is
`Workbook.move_sheet(self, sheet, offset=0)` (see §3). User-visible
calls like:

```python
wb = wolfxl.load_workbook("budget.xlsx", modify=True)
wb.move_sheet("Summary", offset=-2)   # raise NotImplementedError today
wb.save("budget.xlsx")
```

…land on the stub. Target behaviour: re-order the sheet within the
workbook's tab list by `offset` positions, preserve every existing
macro/chart/pivot/comment/table byte-for-byte (modify-mode contract),
and re-point any sheet-scoped defined names whose `localSheetId`
indices shift as a result of the move.

## 2. OOXML Spec Surface

ECMA-376 Part 1, §18.2.27 (`CT_Workbook`) — child ordering of
`<workbook>` is fixed: `fileVersion`, `fileSharing`, `workbookPr`,
`workbookProtection`, `bookViews`, `sheets`, `functionGroups`,
`externalReferences`, `definedNames`, `calcPr`, `oleSize`,
`customWorkbookViews`, `pivotCaches`, `smartTagPr`, `smartTagTypes`,
`webPublishing`, `fileRecoveryPr`, `webPublishObjects`, `extLst`. The
**`<sheets>` block is the source of truth for tab order**: each child
`<sheet name="…" sheetId="…" r:id="…"/>` occupies one position; the
0-based index of a sheet in this block is its "local sheet id" for
the purposes of every other workbook-level reference. `sheetId`
values are stable across reorders (Excel does not renumber them on a
move); only the document order of the children changes.

ECMA-376 Part 1, §18.2.5 (`CT_DefinedName`) — the `localSheetId`
attribute on a `<definedName>` is "the sheet on which the name is
defined". Critically, it is **a 0-based position index into the
`<sheets>` block, NOT a sheet `r:id` and NOT a `sheetId`**. When the
position of a sheet in `<sheets>` changes, every defined name whose
`localSheetId` pointed at the old position must be re-pointed at the
new position; defined names whose `localSheetId` pointed at OTHER
sheets that got shifted by the move must also be re-pointed
accordingly.

Other workbook parts that reference sheets do so by `r:id` (`<sheet
r:id="rId3"/>` in `xl/_rels/workbook.xml.rels`) or by `sheetId`
(macros/VBA), neither of which changes on a reorder. The patcher
therefore does NOT need to touch `xl/_rels/workbook.xml.rels`,
`xl/calcChain.xml`, `xl/sharedStrings.xml`, `xl/worksheets/sheetN.xml`
parts, or any of the ancillary `xl/comments*.xml` /
`xl/tables/table*.xml` parts.

## 3. openpyxl Reference

`.venv/lib/python3.14/site-packages/openpyxl/workbook/workbook.py`,
lines 220-229:

```python
def move_sheet(self, sheet, offset=0):
    """
    Move a sheet or sheetname
    """
    if not isinstance(sheet, Worksheet):
        sheet = self[sheet]
    idx = self._sheets.index(sheet)
    del self._sheets[idx]
    new_pos = idx + offset
    self._sheets.insert(new_pos, sheet)
```

Behaviour notes:

- Accepts either a `Worksheet` instance or a sheet name string.
- `offset=0` is a no-op (the sheet is removed and reinserted at the
  same index).
- `offset` is integer-valued and may be negative.
- `new_pos` is NOT clamped explicitly, but Python's
  `list.insert(idx, x)` already clamps: indices `>= len(list)` append,
  indices `<= -len(list)` prepend. Net behaviour: the sheet always
  ends up somewhere in the list.
- openpyxl does NOT touch `localSheetId` on defined names — the read
  path stores the raw integer and writes it back unchanged. This
  means a user who calls openpyxl's `move_sheet` with sheet-scoped
  defined names ends up with broken references. WolfXL fixes this
  bug; see §7.

What we do NOT copy: the read-path index machinery
(`_sheets.index(sheet)`) which depends on openpyxl's mutable
`_sheets` list of `Worksheet` instances. WolfXL's tab order is
canonically tracked in `Workbook._sheet_names: list[str]` (with
`_sheets: dict[str, Worksheet]` keyed by name); we operate on the
list, not on identity-comparison of objects.

## 4. WolfXL Surface Area

### 4.1 Python coordinator

`python/wolfxl/_workbook.py:297-310` — replace the
`NotImplementedError` body with:

1. Type-check / resolve `sheet`: a `Worksheet` instance gets its
   `.title`; a string is validated against `self._sheet_names`.
2. Locate the source index `idx = self._sheet_names.index(name)`.
3. Compute the target index using the same clamping rule as
   openpyxl's `list.insert` (see §5).
4. Update the in-memory `self._sheet_names` list so subsequent
   reads (`wb.sheetnames`, `wb.worksheets`) see the new order
   immediately.
5. If `_rust_patcher is not None`, queue the move on the patcher via
   the new PyMethod `queue_sheet_move(sheet, offset)`. The patcher
   defers the workbook.xml splice until `save()`.
6. If `_rust_writer is not None` (write mode), the in-memory
   `_sheet_names` change is sufficient on its own — the writer emits
   `xl/workbook.xml` from the model at save time, and tab order falls
   out of `_sheet_names` iteration.

A second flush method, `_flush_pending_sheet_moves_to_patcher`, is
not needed: the queueing happens directly inside `move_sheet`. The
existing `save()` orchestration just runs the patcher's drain loop
which already covers `queued_sheet_moves` via the new Phase 2.5h.

### 4.2 Patcher (modify mode)

New module `src/wolfxl/sheet_order.rs` — a streaming-splice rewriter
for `xl/workbook.xml`:

- `merge_sheet_moves(workbook_xml: &[u8], moves: &[(String, i32)]) ->
  Result<Vec<u8>, String>`. Parses `<sheets>` order, resolves each
  pending `(name, offset)` against the running tab list, computes the
  position remap table, rewrites the `<sheets>` children in place,
  and rewrites every `<definedName localSheetId="N">` whose `N`
  appears in the remap.
- Pure-Rust streaming splice via `quick_xml`. **NOT** a full
  `<workbook>` rewrite — every other child of `<workbook>`
  (`fileVersion`, `workbookPr`, `bookViews`, `calcPr`, `extLst`, …)
  flows through byte-for-byte. Same RFC-021 streaming-splice idiom.
- Empty `moves` slice → identity transform (return source bytes
  verbatim) — preserves the modify-mode no-op invariant.

`XlsxPatcher` (PyO3 class in `src/wolfxl/mod.rs`) gains:

- Field `queued_sheet_moves: Vec<(String, i32)>` — insertion-ordered
  list of pending moves. A single save can queue multiple moves;
  they are applied in queue order against the running tab list.
- PyMethod `queue_sheet_move(sheet: &str, offset: i32)`.
- The empty-queue short-circuit guard in `do_save` adds
  `&& self.queued_sheet_moves.is_empty()` so a workbook with no
  moves still produces a byte-identical save.
- New phase Phase 2.5h, sequenced **before** Phase 2.5f
  (defined-names) since both phases mutate `xl/workbook.xml`. When
  either queue is non-empty we read the source workbook.xml ONCE,
  apply Phase 2.5h's reorder + localSheetId remap, then hand the
  result to Phase 2.5f's defined-names merger. The two phases
  compose without re-parsing workbook.xml.
- A side-effect of Phase 2.5h is to update the in-memory
  `self.sheet_order: Vec<String>` field on the patcher so subsequent
  RFC-020 (`docProps/app.xml`) and RFC-026 (CF aggregation) phases
  see the reordered tab list when they iterate sheet names.

### 4.3 Native writer (write mode)

The native writer at `crates/wolfxl-writer/src/emit/workbook_xml.rs`
emits `xl/workbook.xml` from a structured workbook model. Tab order
in that model is driven by the order in which sheets were
`add_sheet`ed, which mirrors `Workbook._sheet_names`. The Python
coordinator already updates `_sheet_names` in step 4 above; once the
list reflects the post-move order, the next writer flush emits the
correct order. No writer-side patches are needed for this RFC.

There IS one cross-mode asymmetry worth noting: the writer does not
currently emit sheet-scoped defined names with `localSheetId`
attributes (T2 / construction work, post-1.1). Until that lands, the
"defined names with `localSheetId` survive a move" guarantee is a
modify-mode-only feature. See §7.

## 5. Implementation Sketch

### 5.1 Position remap table

Suppose the source workbook has sheets `[A, B, C, D]` at positions
`[0, 1, 2, 3]`, with three defined names referencing positions `0`,
`2`, and `3` via `localSheetId`. The user calls `move_sheet("A",
offset=2)`. The target position is `0 + 2 = 2`.

Apply the move: pop A from index 0, insert it at index 2. New tab
list is `[B, C, A, D]` at positions `[0, 1, 2, 3]`.

The position remap table — old index → new index — is:

| sheet | old pos | new pos |
|------:|--------:|--------:|
| A     | 0       | 2       |
| B     | 1       | 0       |
| C     | 2       | 1       |
| D     | 3       | 3       |

Defined-name rewrites:

- `localSheetId="0"` (was scoped to A) → `localSheetId="2"`.
- `localSheetId="2"` (was scoped to C) → `localSheetId="1"`.
- `localSheetId="3"` (was scoped to D) → `localSheetId="3"` (unchanged).

The remap is computed from a single pass over the new tab list:
`for (new_pos, name) in enumerate(new_order): old_pos =
old_order.index(name); remap[old_pos] = new_pos`. Multiple queued
moves compose by replaying the remap on top of the previous result.

### 5.2 Streaming splice shape

Phase 2.5h reads `xl/workbook.xml`, walks it with `quick_xml`, and
locates two byte ranges:

- The inner content of `<sheets>` (between `<sheets>` and
  `</sheets>`). Each `<sheet …/>` empty-element gets its byte slice
  recorded along with its parsed `name` attribute.
- The inner content of `<definedNames>` (if present). Each
  `<definedName>` start-tag's `localSheetId` attribute is captured
  with its byte position.

The output assembly is:

```text
[ source bytes pre-<sheets> ]
<sheets>
  [ original sheet[i] bytes, in NEW tab-list order ]
</sheets>
[ source bytes between </sheets> and <definedNames> ]
<definedNames>
  [ for each <definedName>: copy verbatim, but rewrite the
    localSheetId attribute value when the integer maps in the remap ]
</definedNames>
[ remaining source bytes ]
```

This preserves attribute order, whitespace, and any unrelated
attributes on `<sheet>` and `<definedName>` elements. Only the
`localSheetId` attribute value (an ASCII integer) is rewritten,
in-place, when its old value is in the remap key set.

### 5.3 No-op invariant

If `queued_sheet_moves` is empty, the `do_save` guard returns the
source bytes verbatim (no `xl/workbook.xml` patch). If it is
non-empty but every queued move resolves to `offset=0` after
clamping (e.g. moving an already-first sheet by `-5`), the merger
still re-emits the block; this is acceptable because the user
explicitly asked for a move. The byte-identity guarantee is that
**no queued move** ⇒ **no change** — not that **a no-op move** ⇒
**no change**. The empty-queue path is what the modify-mode
contract requires.

### 5.4 Index clamping

The patcher clamps the new position to `[0, n-1]` where `n` is the
current tab count. This matches openpyxl's effective behaviour
(`list.insert` clamps), but does it explicitly so the patcher's
remap table is always well-formed. Negative offsets that would
otherwise produce an index `< 0` clamp to `0`; positive offsets that
would push past the end clamp to `n-1`. Same rule applied on both
the Python and Rust sides.

### 5.5 Validation

The Python coordinator validates:

- `sheet` is either a `Worksheet` (resolved to its `.title`) or a
  `str`. Other types raise `TypeError`.
- The resolved sheet name exists in `self._sheet_names`. Otherwise
  raise `KeyError`.
- `offset` is integer-typed (Python's `int` only — `bool` is rejected
  because `isinstance(True, int)` is `True` in Python and we want a
  hard error rather than treating `True` as `1`).

Wrong-type or missing-sheet errors fire at `move_sheet` call time
(eager), not at `save()` time. This matches openpyxl's behaviour for
`self[sheet]` raising `KeyError` on unknown names.

## 6. Verification Matrix

1. **Rust unit tests** (`cargo test`) — inline `#[cfg(test)]` module
   in `src/wolfxl/sheet_order.rs` covering: forward move, backward
   move, no-op move, clamp-high, clamp-low, single-sheet workbook,
   workbook with no `<definedNames>`, workbook with `<definedNames>`
   that reference unaffected sheets, multi-move composition.

2. **Golden round-trip (diffwriter)** — N/A. This RFC's surface is
   modify-mode only; the writer requires no changes (§4.3). The
   patcher's byte-stable output is verified by direct round-trip
   tests.

3. **openpyxl parity** — the `tests/parity/` sweep continues to pass
   97/97. No `move_sheet` parity tests are added because openpyxl's
   `move_sheet` does NOT remap `localSheetId` (see §3 / §7). A
   `localSheetId` parity test would either spuriously fail (we are
   correct, openpyxl is buggy) or require a bug-for-bug feature flag
   which §10 defers post-1.1.

4. **LibreOffice cross-renderer** — manual: open a moved-sheet xlsx
   in LibreOffice and confirm that (a) the tab order matches Excel,
   (b) any sheet-scoped print area `localSheetId="N"` still resolves
   to the correct sheet. The fixture `tests/fixtures/` does not need
   a new permanent file; a tester can produce one with the new test
   helpers in `tests/test_move_sheet_modify.py`.

5. **Cross-mode** — covered by the round-trip tests: write-mode
   path works because `_sheet_names` is the writer's source of
   truth, and the in-memory update in §4.1 step 4 happens for both
   modes.

6. **Regression fixture** — `tests/test_move_sheet_modify.py` builds
   fixtures inline with openpyxl (matching the precedent of
   `test_defined_names_modify.py` and `test_modify_hyperlinks.py`)
   so the tests are hermetic and don't add a permanent
   `tests/fixtures/` artifact.

The standardized "done" gate is `python scripts/verify_rfc.py --rfc
036 --quick`.

## 7. Cross-Mode Asymmetries

- **modify mode (RFC-036, this slice)**: `move_sheet` re-points
  `<definedName localSheetId="…">` indices correctly. Sheet-scoped
  print areas, named ranges, and other defined names survive the
  move and continue to resolve to their original sheets.
- **write mode**: the writer does not yet emit `localSheetId` on
  defined names (T2 construction work, post-1.1). Until that lands,
  a workbook constructed in write mode has no sheet-scoped defined
  names to remap, and `move_sheet` reduces to "update
  `_sheet_names`". The asymmetry is invisible to users today; it
  becomes a divergence only after T2 ships sheet-scoped defined-name
  construction.
- **openpyxl parity**: openpyxl's `move_sheet` does not remap
  `localSheetId`; we diverge here. WolfXL's behaviour is the
  Excel-correct one (see §3). The divergence will be flagged in the
  1.1 release notes (`KNOWN_GAPS.md`) the same way we flag every
  better-than-openpyxl behaviour.

## 8. Risks

| # | Risk | Likelihood | Impact | Mitigation |
|---|---|---|---|---|
| 1 | The `<sheets>` element is wrapped in a default XML namespace declaration; `quick_xml` reads `local_name() == "sheet"` correctly, but a buggy match on `name() == "sheet"` would miss it. | low | high | Use `local_name()` exclusively (matching RFC-021 §4 / `defined_names.rs`). Add a Rust unit test fixture using the namespaced form. |
| 2 | The user mutates defined names in the same `save()` that does a `move_sheet`. Two phases racing on `xl/workbook.xml`. | medium | high | Phase 2.5h reads workbook.xml once, applies the reorder + localSheetId remap, and feeds the result to Phase 2.5f. The defined-names merger is unaware of the move; it operates on the post-move bytes. Tested by `test_rfc036_compose_with_defined_names`. |
| 3 | A `<sheet>` element carries an attribute like `state="hidden"`; rewriting must preserve it. | low | medium | The splice copies each `<sheet …/>` element's full byte slice and only reorders the slices. No attribute rewriting on `<sheet>`. |
| 4 | A user queues many moves; the cumulative remap is wrong. | low | medium | The implementation replays each move against a running tab list, recomputing the remap after each apply. Tested by `test_rfc036_multiple_moves_compose`. |
| 5 | Sheet name with XML-special characters (`&`, `<`) in `<sheet name="…">`. | low | low | The reorder operates on byte slices of the `<sheet>` element, never on the parsed `name` attribute. The parsed name is used only for lookup against the user's `move_sheet` argument; user-supplied names are compared as Rust `String` values. |

## 9. Effort Breakdown

| Slice | Estimate | Notes |
|---|---|---|
| Research | ½d | This document. |
| Rust impl | 1d | `sheet_order.rs` + `mod.rs` wiring. |
| Python wiring | ½d | Replace stub; queue, validate, update `_sheet_names`. |
| Tests | 1d | 6+ pytest cases + Rust units. |
| Review | ½d | |

Total: ~3-4 days, well within the M bucket.

## 10. Out of Scope

- **Sheet rename** (`ws.title = "NewName"`): a separate post-1.1 RFC.
  Renaming touches `<sheet name="…">`, every cross-sheet formula
  reference (RFC-012's translator at the patcher edge), every
  `<definedName>` whose formula text references the sheet, and every
  external part that names the sheet (chart titles, pivot caches).
  RFC-036 is REORDER-only.
- **Sheet add / remove**: out of scope. Add lives under T2
  construction (`Workbook.create_sheet` already works in write mode);
  remove via `del wb["Sheet"]` is tracked as a separate stub.
- **Re-pointing formula text** that uses `Sheet1!$A$1` notation:
  `move_sheet` does NOT change sheet names, so formulas referencing
  `Sheet1!…` continue to resolve correctly without translation. RFC-
  012 is unused by this RFC.
- **`bookViews` `activeTab` re-pointing**: out of scope. The
  `activeTab` attribute on `<workbookView>` is a 0-based index into
  the post-reorder `<sheets>` list, and Excel/LibreOffice both
  recompute it on next open. We could update it for fidelity; we
  defer to the broader "view state" RFC where workbook views are
  modelled end-to-end.
- **Calc-chain invalidation**: `xl/calcChain.xml` references cells
  by sheet `r:id` (which does not change on a move) and by formula
  position (which also does not change). No mutation needed.
- **Custom workbook views**: `<customWorkbookViews>` is preserved
  byte-for-byte by the streaming splice. If a custom view stored a
  position-based active-tab pointer, it lives outside this RFC's
  scope.

## Acceptance

(Filled in after Shipped.)
