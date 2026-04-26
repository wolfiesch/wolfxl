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

### Phase 3 — T2 Rich-text reads (✅ SHIPPED in 1.3, Sprint Ι Pod-α)

Both reads and writes round-trip via the new
``wolfxl.cell.rich_text.{CellRichText, TextBlock, InlineFont}``
shims (matching openpyxl's iteration / equality protocol).

* ``Cell.rich_text`` always returns the structured runs (or ``None``
  for plain cells), regardless of how the workbook was opened.
* ``Cell.value`` keeps its prior contract by default — flattens
  rich-text to plain ``str`` so existing call sites are
  unaffected.  Pass ``load_workbook(..., rich_text=True)`` to
  flip ``Cell.value`` to return ``CellRichText`` for cells whose
  backing string carries `<r>` runs (matches openpyxl 3.x's own
  ``rich_text=True`` flag).
* Setting ``cell.value = CellRichText([...])`` round-trips in both
  write mode (native writer emits inline-string runs) and modify mode
  (patcher emits inline-string runs, SST left untouched).

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
- Pivot tables, charts, images, data validation — not in SynthGL's openpyxl surface.

## Closed in 1.3 (Sprint Ι Pod-α)

- ✅ **Rich-text read** — Phase 3 row above is now SHIPPED.
- ✅ **Rich-text write** — was previously listed as out-of-scope T3.
  Sprint Ι Pod-α shipped both write-mode (native writer) and
  modify-mode (patcher) inline-string emit paths, so user code that
  builds a workbook with rich-text cells round-trips end-to-end via
  wolfxl's writer.  The SST is intentionally left untouched — runs
  are emitted as inline strings (`t="inlineStr"` + `<is>`), matching
  openpyxl's own write path verbatim.

## Modify mode — T1.5 audit (now closed) and structural extensions

Modify mode (`load_workbook(path, modify=True)`) is served by `XlsxPatcher`,
which surgically rewrites changed parts and copies everything else verbatim.
The W4F audit originally enumerated seven mutation paths that were deferred
to a post-Wave-5 T1.5 slice. **All seven shipped in WolfXL 1.1's Phase 3**
(per `tests/test_modify_mode_independence.py` lines 14-21 and the per-RFC
modify-mode test files). Structural mutations and `copy_worksheet` followed.
This table is now a status snapshot, not a deferred-work list.

| Modify-mode mutation | Status |
|---|---|
| `wb.properties.title = ...` (any property mutation) on existing file | ✅ Shipped — RFC-020 (`tests/test_modify_properties.py`, `tests/test_workbook_properties_t1.py`) |
| `wb.defined_names[name] = DefinedName(...)` on existing file | ✅ Shipped — RFC-021 (`tests/test_defined_names_modify.py`). The `__setitem__` proxy is exposed; round-trip verified end-to-end. |
| `cell.comment = Comment(...)` | ✅ Shipped — RFC-023 (`tests/test_comments_modify.py`) |
| `cell.hyperlink = Hyperlink(...)` | ✅ Shipped — RFC-022 (`tests/test_modify_hyperlinks.py`, `tests/test_hyperlink_internal_flag.py`) |
| `ws.add_table(Table(...))` | ✅ Shipped — RFC-024 (`tests/test_tables_modify.py`) |
| `ws.data_validations.append(...)` | ✅ Shipped — RFC-025 (`tests/test_modify_data_validations.py`) |
| `ws.conditional_formatting.add(...)` | ✅ Shipped — RFC-026 (`tests/test_modify_conditional_formatting.py`) |
| Sheet/column/row structural mutations | ✅ Shipped — `insert_rows`/`delete_rows` (RFC-030), `insert_cols`/`delete_cols` (RFC-031), `Worksheet.move_range` (RFC-034), `Workbook.move_sheet` (RFC-036). |
| `wb.copy_worksheet(...)` (modify mode) | ✅ Shipped — RFC-035 in 1.1 (Sprint Ζ Pod-δ closed four of six composition gaps; two remain xfail per "RFC-035 cross-RFC composition gaps" below). See divergence section below. |
| `wb.copy_worksheet(...)` (write mode) | ✅ Shipped — Sprint Θ (1.2) Pod-C1 lifts the §3 OQ-a `NotImplementedError`. |

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
Sprint Ζ; Sprint Θ Pod-A closed bug #4 via the new
``permissive=True`` loader flag. One case (#6) remains as a
documented 1.2 follow-up.

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

#### Fixed in 1.2 (Sprint Θ Pods A + B)

- ✅ **#4 `test_p_self_closing_sheets_block`** —
  fixed by Sprint Θ Pod-A commit `fix(rfc-035): add permissive=True
  loader mode, close bug #4 self-closing <sheets/>`.
  `wolfxl.load_workbook(..., permissive=True)` now falls back to the
  workbook rels graph when `xl/workbook.xml`'s `<sheets>` block is
  empty / self-closing: each worksheet relationship target is
  registered under a synthesized title (`Sheet1`, `Sheet2`, ...) and
  the in-memory workbook.xml is rewritten to expose the synthesized
  `<sheet>` entries so downstream phases (Phase 2.7 splice, defined-
  names merger) see a well-formed workbook. The flag defaults to
  `False`; well-formed inputs are unaffected. Test flips xfail →
  PASS end-to-end (load → copy_worksheet → save → reload via
  openpyxl).

- ✅ **#6 `test_r_cdata_pi_fuzz_fakeout`** —
  fixed by Sprint Θ Pod-B commit `fix(rfc-035): replace naive splice
  with quick-xml SAX scan, close bug #6 CDATA fakeout`. The Phase 2.7
  `splice_into_sheets_block` helper now drives a `quick_xml::Reader`
  over `xl/workbook.xml` and locates the real `<sheets>` open/close
  by event-stream nesting depth rather than byte-substring search.
  Comments, CDATA sections, and processing instructions surface as
  separate quick-xml events and are ignored, so a workbook.xml
  comment containing the literal `</sheets>` token no longer
  perturbs the splice point. Five new Rust unit tests pin the
  invariant (normal, self-closing, comment fakeout, CDATA fakeout,
  malformed). Test flips xfail → PASS.

  As a side-fix, the test fixture helper
  `_inject_comment_with_sheets_token` was anchoring on `?>` (XML
  declaration close) — but openpyxl-saved workbooks omit the XML
  decl, so the injection was a silent no-op that masked the bug.
  The helper now anchors on the `<workbook ...>` opening tag via
  regex.

#### Deferred queue (post-1.2)

All RFC-035 cross-RFC composition gaps surfaced in Sprint Ζ have been
closed by Sprints Ζ (Pod-δ: #1, #2, #3, #5) and Θ (Pods A+B: #4, #6).
No further deferred items remain in this category.
