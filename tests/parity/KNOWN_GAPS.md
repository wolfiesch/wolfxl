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

The read-parity harness is down to a single fixture-specific xfail.

### `number_format` parity on styled-but-empty template cells

In `time_series/ilpa_pe_fund_reporting_v1.1.xlsx`, openpyxl materializes
missing coordinates inside the worksheet dimension as synthetic blank cells
with `number_format == 'General'`. WolfXL still surfaces the worksheet style
grid for some of those blank coordinates, so `Cell.number_format` reports the
template's currency/count format instead.

**Fix sketch:** resolve `number_format` from the sparse worksheet cell model,
not just the positional style grid. The exact openpyxl contract appears to be
"only workbook-backed cells carry style-derived number formats; synthetic blank
cells inside the used range stay General." That likely means reading style IDs
directly from worksheet XML and treating absent cells as style-less even when
the style grid has a positional format.

## Out of scope (documented, not planned)

- Writing encrypted xlsx. Decision: T3 per plan — document in migration guide.
- Rich-text write. Decision: T3 — SynthGL has no current write use case.
- Pivot tables, charts, images, data validation — not in SynthGL's openpyxl surface.
