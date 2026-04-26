# Known parity gaps — WolfXL vs openpyxl

This file enumerates every openpyxl symbol that SynthGL relies on but WolfXL
0.3.2 does not yet expose (or exposes under a different name). Each gap is
tied to a phase in the rollout plan.

Gaps are also encoded in `openpyxl_surface.py` via `wolfxl_supported=False`
— the parity smoke test keeps the two in sync.

## Gate

- Every gap must have a phase owner.
- Closing a gap: flip `wolfxl_supported=True`, remove the entry here, expect
  `test_known_gap_still_gaps` to fail red (which is the signal to also
  commit the ratchet-baseline update).

## Gaps by category

### Sheet access (type-hint imports — SHIPPED)

`wolfxl.Worksheet` and `wolfxl.Cell` are now re-exported at the top level
(see `python/wolfxl/__init__.py`). SynthGL's type-hint imports work as a
drop-in.

### Range / layout API shape (Phase 0 cleanup — SHIPPED)

`Worksheet.max_row`, `Worksheet.max_column`, and `Worksheet.merged_cells`
are now public properties. `merged_cells` returns a `_MergedCellsProxy`
backed by the Rust `read_merged_ranges` call in read mode (closes the
"merged_cells empty on read" per-fixture gap below as a side-effect).

### Utils (Phase 0 cleanup — SHIPPED)

All seven utility symbols ship through `python/wolfxl/utils/`:

- `wolfxl.utils.cell.get_column_letter`, `column_index_from_string`,
  `range_boundaries`, `coordinate_to_tuple`
- `wolfxl.utils.numbers.is_date_format`
- `wolfxl.utils.datetime.from_excel`, `CALENDAR_WINDOWS_1900`

Behavior is bug-for-bug compatible with openpyxl 3.1.x and pinned by
`test_utils_parity.py`. Bound checks (`get_column_letter` capped at 18278
= ZZZ) and the 1900 leap-year correction (`from_excel`) match openpyxl
verbatim.

### Phase 1 — T1 DefinedName WRITE

| openpyxl path | phase | note |
|---|---|---|
| `Workbook.defined_names["X"] = DefinedName(...)` | Phase 1 | Rust side (`add_named_range`) already exists; just expose `__setitem__` in the Python proxy. |

### Phase 2 — T0 Password-protected reads

| openpyxl path | phase | note |
|---|---|---|
| `openpyxl.load_workbook(path, ...)` on encrypted file | Phase 2 | Add `password=` kwarg; dispatch through `msoffcrypto-tool` → `CalamineStyledBook.open_bytes()`. |

### Phase 3 — T2 Rich-text reads

| openpyxl path | phase | note |
|---|---|---|
| `Cell.value` when backing is `CellRichText` | Phase 3 | Currently wolfxl flattens rich text to plain. Add `Cell.rich_text` property (iter-compatible with `openpyxl.cell.rich_text.CellRichText`). |

### Phase 4 — T2 Streaming reads

| openpyxl path | phase | note |
|---|---|---|
| `openpyxl.load_workbook(path, read_only=True)` + `ws.iter_rows(values_only=True)` on 1M-cell sheets | Phase 4 | WolfXL accepts the kwarg but reads the full sheet into memory. Add a SAX fast path for `read_only=True` or sheets > 50k rows. |

### Phase 5 — T1 .xls / .xlsb

| openpyxl path | phase | note |
|---|---|---|
| `openpyxl.load_workbook('foo.xlsb')` | Phase 5 | openpyxl itself doesn't read xlsb; parity target is pandas-style "same values came out". Migrate WolfXL from `calamine-styles` to upstream `calamine` (xlsb native). |
| `openpyxl.load_workbook('foo.xls')` | Phase 5 | openpyxl doesn't read xls either. Parity target is xlrd behavior. |

## Per-fixture read gaps (surfaced by Phase 0 baseline run)

Phase 0's read-parity xfail list is now empty. Any newly discovered
fixture-specific drift should be added to `test_read_parity.py::KNOWN_FIXTURE_GAPS`
and documented here before the ratchet baseline is updated.

## Out of scope (documented, not planned)

- Writing encrypted xlsx. Decision: T3 per plan — document in migration guide.
- Rich-text write. Decision: T3 — SynthGL has no current write use case.
- Pivot tables, charts, images, data validation — not in SynthGL's openpyxl surface.

## Modify mode — T1.5-deferred features (W4F audit)

Modify mode (`load_workbook(path, modify=True)`) is served by `XlsxPatcher`,
which surgically rewrites changed parts and copies everything else verbatim.
Several mutation paths are intentionally deferred to a post-Wave-5 T1.5
slice. Each raises `NotImplementedError` with "T1.5" in the message so
callers see a clear migration hint rather than silent data loss.

| Modify-mode mutation | Status |
|---|---|
| `wb.properties.title = ...` (any property mutation) on existing file | Raises — T1.5 |
| `wb.defined_names[name] = DefinedName(...)` on existing file | Raises — T1.5 |
| `cell.comment = Comment(...)` | Raises — T1.5 |
| `cell.hyperlink = Hyperlink(...)` | Raises — T1.5 |
| `ws.add_table(Table(...))` | Raises — T1.5 |
| `ws.data_validations.append(...)` | Raises — T1.5 |
| `ws.conditional_formatting.add(...)` | Raises — T1.5 |
| Sheet/column/row structural mutations | **SHIPPED in WolfXL 1.1** — `insert_rows`/`delete_rows` (RFC-030), `insert_cols`/`delete_cols` (RFC-031), `Worksheet.move_range` (RFC-034), `Workbook.move_sheet` (RFC-036). |
| `wb.copy_worksheet(...)` | **Modify-mode only — SHIPPED in WolfXL 1.1 (RFC-035)**. See divergence section below. Write-mode raises `NotImplementedError` per §3 OQ-a, tracked for 1.2. |

Supported in modify mode (round-trips cleanly via `_flush_to_patcher`):

- Cell values: string, number, boolean, formula, blank
- Font (bold/italic/underline/strikethrough/size/name/color)
- Fill (solid pattern bg color)
- Alignment (horizontal/vertical/wrap/indent/rotation)
- Number format
- Borders (left/right/top/bottom — style + color)

`tests/test_modify_mode_independence.py` encodes these contracts as
pre-rip-out invariants: any future change that breaks the patcher's
independence from the writer backend, or that silently falls through to
the writer for a T1.5-deferred feature, fails CI immediately.

## RFC-035 — `copy_worksheet` divergences from openpyxl (SHIPPED 1.1)

`Workbook.copy_worksheet` ships in modify mode in WolfXL 1.1
(RFC-035). The behaviour deliberately diverges from openpyxl's
`WorksheetCopy` in five places. Each divergence is asserted by
`tests/parity/test_copy_worksheet_parity.py`; this section is the
ratchet-tracked record of WHY wolfxl preserves what openpyxl drops.

| Feature | WolfXL behaviour | openpyxl behaviour | Rationale |
|---|---|---|---|
| Tables | Cloned with auto-renamed `name`/`displayName` (`{base}_{N}`, N starts at 2 per RFC-035 §3 OQ-b). New `<table id>`, new content-type `<Override>`, new rels entry. | `WorksheetCopy._copy_cells` walks the in-memory cell dict only; `<tableParts>` is silently dropped. | wolfxl operates on ZIP bytes, not an in-memory model — clone preserves the source's full feature surface. |
| Data validations | Cloned in-place inside the cloned sheet's XML. | Dropped — `WorksheetCopy` has no DV-handling branch. | DV is part of "what makes the sheet work"; dropping it silently degrades cloned templates. |
| Conditional formatting | Cloned in-place inside the cloned sheet's XML (with cross-sheet `dxfId` allocation). | Dropped — same reason as DV. | Same as DV. |
| Sheet-scoped defined names | Fresh entries emitted with `localSheetId == new_idx` (post-copy tab position) per §3 OQ-c. Source's sheet-scope names retained. | Dropped — `WorksheetCopy` does not touch `xl/workbook.xml`'s `<definedNames>`. | `_xlnm.Print_Area` is the canonical "make-this-sheet-printable" hint; dropping it silently breaks print-preview on the clone. |
| Image media | Aliased — cloned drawing rels point at the same `xl/media/imageN.png` as the source. | Deep-copied — pillow re-encodes the image binary on the clone. | Avoids 50× bloat on workbooks with logo images and many sheet copies. RFC-035 §5.3 documents the contract; future "modify a copy's image" RFC will deep-clone. |
| Calc chain (`xl/calcChain.xml`) | Not mutated — Excel rebuilds it on next open. | Same. | calcChain is a perf optimization, not a correctness contract. |

### RFC-035 cross-RFC composition gaps

Surfaced by Pod-γ's full harness (`tests/test_copy_worksheet_modify.py`,
originally six xfail cases). Pod-δ closed four of the six in
Sprint Ζ; two remain as documented 1.2 follow-ups (low real-world
impact — only reachable through synthesized fixtures or naive-splice
fakeouts).

#### Fixed in 1.1 (Sprint Ζ Pod-δ)

- ✅ **#1 `test_i_copy_and_edit_copy_in_same_save`** —
  fixed by Pod-δ commit `fix(rfc-035): patch cloned-sheet bytes through
  file_adds, not zip`. Phase 3 now reads cloned-sheet bytes from
  `file_adds` / `file_patches` first, falling back to the source ZIP
  only for genuine source-side sheets, and routes the rewrite back to
  `file_adds` for cloned paths so Phase 4's new-entry pass picks up the
  patched bytes. Test flips xfail (strict, OSError) → PASS.
- ✅ **#2 `test_j_copy_then_move_sheet_in_same_save`** —
  fixed by Pod-δ commit `fix(rfc-035): seed Phase 2.5h workbook.xml
  read from file_patches`. Phase 2.5h's reorder pass now prefers
  `file_patches["xl/workbook.xml"]` over the source-ZIP read so the
  Phase 2.7 → Phase 2.5h handoff happens through the shared
  `file_patches` map (the intended composition per RFC-035 §5.4).
  Test flips xfail (strict, AssertionError) → PASS.
- ✅ **#3 `test_k_copy_then_add_table_to_copy`** —
  fixed by the same commit as #1 (shared root cause). The Phase 2.5f
  rels-graph load also probes `file_adds` / `file_patches` first so a
  user `add_table` on a cloned sheet sees the cloned rels graph
  rather than an empty fallback. Test flips xfail (strict, OSError)
  → PASS.
- ✅ **#5 `test_q_defined_names_upsert_collision`** —
  fixed by Pod-δ commit `fix(rfc-035): route cloned defined names
  through RFC-021 merger`. Phase 2.7's `defined_names_to_add` push
  now scans `queued_defined_names` for a matching
  `(name, local_sheet_id)` key and skips on hit so the user's
  explicit upsert wins over the planner's default (per RFC-035 §5.4
  and Pod-β's last-write-wins-on-the-USER invariant). Test flips
  xfail (strict, AssertionError) → PASS.

#### Deferred to 1.2

| # | Failing case | Symptom | Why deferred |
|---|---|---|---|
| 4 | `test_p_self_closing_sheets_block` | wolfxl loader rejects synthesized `<sheets/>` workbook.xml fixtures, so the splice's self-closing branch is unreachable through the public API. | Real Excel never emits a self-closing `<sheets/>` for a non-empty workbook. Reachable only via direct ZIP edit. 1.2 follow-up: add a Rust-level Phase 2.7 unit test that exercises the splice on a self-closing input directly when the loader gains a `permissive=True` mode. |
| 6 | `test_r_cdata_pi_fuzz_fakeout` | Phase 2.7 splice is naive — a workbook.xml comment containing literal `</sheets>` may fool the byte-level locator. | Acknowledged in Pod-β's handoff note as "acceptable for 1.1 since no real Excel-emitted workbook contains it". 1.2 follow-up: promote the splice to a SAX/quick-xml-driven scan that respects element nesting. |

Tracked by: `tests/test_copy_worksheet_modify.py` — the two remaining
deferred cases are still pinned with `xfail(strict=True)`. When a
fix lands, remove the xfail marker and update this table.
