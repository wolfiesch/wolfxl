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

The harness (read_parity test) discovered these on real SynthGL fixtures.
Each is xfailed in `test_read_parity.py::KNOWN_FIXTURE_GAPS` and will flip
green once the underlying wolfxl bug is fixed in Phase 0 cleanup or later.

### `Cell.value` type narrowing — date vs datetime

WolfXL returns `datetime.date(2007, 2, 25)` where openpyxl returns
`datetime.datetime(2007, 2, 25, 0, 0)` for cells whose date format has no
time component. Affected fixtures: `flat_register/excelx_*.xlsx`.

**Fix sketch:** in `python/wolfxl/_cell.py::_payload_to_python`, when payload
type is `"date"`, return a `datetime` (not `date`) to mirror openpyxl. If
preserving the date-vs-datetime distinction matters for some callers, expose
both `cell.value` (datetime) and `cell.date_value` (date).

### `data_only=True` is not honored — formulas return as strings

WolfXL ignores the `data_only` kwarg and returns formula strings (`'=P4'`)
where openpyxl returns the cached evaluated value. Affected fixture:
`time_series/ilpa_pe_fund_reporting_v1.1.xlsx`.

**Fix sketch:** in `python/wolfxl/__init__.py::load_workbook`, when
`data_only=True` is passed, set a flag on the workbook that causes
`_payload_to_python` to return `payload["value"]` (cached) instead of
`payload["formula"]` for formula cells. The cached value is already in the
calamine payload — wolfxl just isn't surfacing it.

### `merged_cells` empty on read

WolfXL's `Worksheet.merged_cells` is a private `_merged_ranges` set that
reads from local writer state, not from the underlying xlsx. When a
workbook is opened in read mode, merged ranges are never populated.
Affected: any fixture with merged cells.

**Fix sketch:** wire `CalamineStyledBook.read_merged_cells(sheet)` (the Rust
side already supports this — used by other read paths) into the Worksheet's
`merged_cells.ranges` accessor.

### `number_format` backslash escapes stripped

WolfXL returns `'([$-409]mmm-yy -'` where openpyxl returns
`'\\([$-409]mmm\\-yy\\ \\-'`. WolfXL is stripping the backslash escapes
that openpyxl preserves verbatim from the xlsx XML.

**Fix sketch:** in the Rust `styles.rs` number-format extraction, preserve
the raw string from `xl/styles.xml` (don't post-process). openpyxl returns
exactly what's in the file.

### `max_row`/`max_column` off-by-one on sheets with trailing blanks

WolfXL reports `max_row=200`, openpyxl reports `201` when the trailing rows
are styled-but-blank. Affected: sheets with template formatting in unused
rows.

**Fix sketch:** match openpyxl's heuristic — `max_row` is the highest row
index with ANY occupied cell (value OR explicit format), not just the
highest with a non-empty value.

## Out of scope (documented, not planned)

- Writing encrypted xlsx. Decision: T3 per plan — document in migration guide.
- Rich-text write. Decision: T3 — SynthGL has no current write use case.
- Pivot tables, charts, images, data validation — not in SynthGL's openpyxl surface.
