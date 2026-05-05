# RFC-073 — `Workbook(write_only=True)` streaming write mode (Sprint 7 / G20)

> **Status**: Implemented
> **Owner**: Claude (S7 design + implementation)
> **Sprint**: S7 — Streaming Write
> **Closes**: G20 (`write_only=True` streaming write mode) in the openpyxl parity program

## 1. Goal

Make `Workbook(write_only=True)` actually stream rows to disk instead of routing through the standard in-memory writer, so 10M-row ETL exports run with bounded peak RSS:

```python
wb = wolfxl.Workbook(write_only=True)
ws = wb.create_sheet("data")
for chunk in stream_from_db(limit=10_000_000):
    for row in chunk:
        ws.append(row)
wb.save("export.xlsx")
# Peak RSS dominated by SST + styles, NOT by row data.
```

Today `_workbook.py` accepts the kwarg but builds the same eager `Worksheet` that materialises every cell in four memory layers (`_append_buffer` → `_cells` → Rust `BTreeMap<u32, Row>` → `String` accumulator → `Vec<u8>` ZIP).

## 2. Problem statement

For a 10M-row × 10-col export, the eager path holds ~40 GiB of cell metadata in memory before the ZIP buffer ever opens. The user-facing promise of `write_only=True` — "bounded memory, write-once, append-only" — does not match what the writer does. Spec status `partial`, gap_id `G20`.

## 3. Public contract

```python
Workbook(*, write_only: bool = False) -> Workbook
wb.create_sheet(title: str) -> WriteOnlyWorksheet  # when wb.write_only

ws.append(row: Iterable[Any]) -> None              # only row-write API
ws.freeze_panes = "B2"                             # before any append
ws.print_area = "A1:Z1000"                         # before any append

wb.save(path)                                      # consumed-on-save
wb.save(path)                                      # raises WorkbookAlreadySaved

WriteOnlyCell(ws, value, *, font, fill, border, alignment, number_format)
```

Forbidden — raise `AttributeError`:

- `ws[...]` — random coordinate access
- `ws.cell(...)`, `ws.iter_rows`, `ws.iter_cols`, `ws.rows`, `ws.columns`
- `ws.merge_cells`, `ws.add_chart`, `ws.add_image`, `ws.add_table`, `ws.add_data_validation`, `ws.add_pivot_table`, `ws.conditional_formatting`, `ws.auto_filter`

These are openpyxl's exact restrictions (`_write_only.py:51-66`).

## 4. Architecture

### 4.1 Splice point

The eager emit pipeline runs through 38 OOXML slots in `sheet_xml::emit`. Slot 6 is `<sheetData>...</sheetData>`. For streaming write-only sheets, slot 6 reads from a per-sheet temp file:

```
[ws.append(row)]    [wb.save(path)]
       │                  │
       ▼                  ▼
┌────────────────┐  ┌──────────────┐
│ StreamingSheet │  │ emit_xlsx    │
│ ─────────────  │  │ ───────────  │
│ BufWriter      │  │ for ws:      │
│   over          │  │  sheet_xml::│
│ NamedTempFile  │  │  emit(ws,..) │
│                │  │   ...        │
│ append_row(    │  │   slot 6:    │
│   row, sst)    │◄─┤    splice    │
└────────────────┘  │    temp file │
                    │   ...        │
                    │ → ZIP entry  │
                    └──────────────┘
```

Cell encoding lives in `pub(crate) fn emit_row_to<W: fmt::Write>(...)`. Both eager and streaming paths call it, so byte-equality is structural — not a test invariant we have to police.

### 4.2 What stays in memory

- **SST**: every unique string. Same shape as openpyxl's `lxml.xmlfile` SST.
- **StylesBuilder**: every unique style record. Same shape as openpyxl.
- **Per-sheet emit `String`**: one sheet's worth of XML during `sheet_xml::emit`. Saturates with sheet count, not with row count.

### 4.3 Crash safety

`tempfile::NamedTempFile` runs an OS-level cleanup hook on graceful drop. SIGKILL leaves the file behind; the prefix `wolfxl-stream-{pid}-{sheet_idx}-` lets admins find orphans, and standard temp-dir retention policies (systemd-tmpfiles, macOS periodic cleanup) reap them.

### 4.4 Forbidden methods

`WriteOnlyWorksheet.__getattr__` rejects names in `_FORBIDDEN_ATTRS`. `__getitem__` / `__setitem__` reject coordinate access unconditionally. `freeze_panes` / `print_area` setters guard `_row_counter > 0` and raise `RuntimeError` because the relevant XML slots live BEFORE `<sheetData>` and the streaming temp file has already opened it.

## 5. Implementation surface

**Rust (`crates/wolfxl-writer/`)**:

- `src/streaming.rs` (NEW) — `StreamingSheet`, `IoFmtAdapter`. Owns `tempfile::NamedTempFile`. v2 added `splice_into_writer<W: Write>` that uses `std::io::copy` (8 KiB chunks) instead of materialising the temp file into a `String`.
- `src/emit/sheet_data.rs` — extracted `pub(crate) fn emit_row_to<W: fmt::Write>` shared between eager `String` and streaming `BufWriter<File>` sinks.
- `src/emit/sheet_xml.rs` — slot 6 detects `sheet.streaming` and splices for the `String`-buffered eager-shape path. v2 added `pub fn emit_streaming_to<W: io::Write>(sheet, idx, styles, dest)` — writes head/tail markup directly into `dest`, splices the temp file straight through `splice_into_writer`. Used by the v2 dispatch path.
- `src/emit/dimension.rs` — slot 4 reads `stream.row_count()` / `stream.max_col()` for the bounding range.
- `src/model/worksheet.rs` — added `streaming: Option<StreamingSheet>` field; removed `#[derive(Clone)]` (no production callers cloned `Worksheet`).
- `src/zip.rs` (v1.5) — `pub fn package_to<W: Write + Seek>(entries, dest) -> io::Result<()>` streams the archive directly into `dest`; `package(entries) -> Vec<u8>` is now a thin wrapper.
- `src/lib.rs` (v1.5) — `pub fn emit_xlsx_to<W: Write + Seek>(wb, dest) -> io::Result<()>` is the streaming entry point; `emit_xlsx(wb) -> Vec<u8>` is a wrapper. v2 added private `EmitOp` enum (`Bytes(ZipEntry)` vs `SheetStream { path, sheet_idx }`) and `package_emit_ops(&ops, &wb, dest)` helper that owns the `ZipWriter` loop and dispatches streaming sheets through `emit_streaming_to`.

**Rust PyO3 bridge (`src/`)**:

- `src/native_writer_streaming.rs` (NEW) — `enable_streaming`, `append_streaming_row`, `finalize_all_streaming` helpers.
- `src/native_writer_backend.rs` — three `#[pymethods]` methods on `NativeWorkbook` plus `intern_format` for cached style resolution.
- `src/native_writer_workbook.rs:save_once` — flushes streaming `BufWriter`s, then opens a `BufWriter<File>` and calls `emit_xlsx_to(wb, &mut writer)` (v1.5: streams ZIP straight to disk instead of materialising it as `Vec<u8>` first).
- `src/native_writer_cells.rs` — `payload_to_write_cell_value` made `pub(crate)` so the streaming bridge reuses it.

**Python (`python/wolfxl/`)**:

- `_workbook.py` — `__init__(*, write_only=False)`. When True: skip the default `Sheet`, set `_write_only` and `_saved`. `write_only` property reads `self._write_only`.
- `_workbook_sheets.py:create_sheet` — dispatches to `WriteOnlyWorksheet` when `wb._write_only`.
- `_workbook_save.py:save_workbook` — re-entry guard raises `WorkbookAlreadySaved`. Routes through `save_write_only_mode`.
- `_worksheet_write_only.py` (NEW) — `WriteOnlyWorksheet`, `WriteOnlyCell`. Style resolution caches by `(id(font), id(fill), id(border), id(alignment), number_format)`.

**Tests**:

- `crates/wolfxl-writer/tests/streaming_write.rs` — 6 cargo cases (well-formedness, SST registration, save splice, empty `<sheetData/>`, byte-parity at the per-sheet emit level, plus v2's `streaming_save_emits_byte_identical_sheet_to_eager` covering the streaming-direct save path through the full ZIP container).
- `tests/test_write_only.py` — 12 focused tests.
- `tests/test_openpyxl_compat_oracle.py` — strengthened existing `workbook_write_only_streaming`; new `workbook_write_only_bounded_memory` (100k × 10 numeric cells, peak RSS < 80 MiB via psutil in subprocess).

**Docs**:

- `docs/migration/_compat_spec.py` — `streaming.write_only` flipped `partial → supported`; secondary_probes registered.
- `docs/migration/compatibility-matrix.md` — regenerated.
- `Plans/openpyxl-parity-program.md` — G20 row flipped `proposed → landed`.

## 6. Scope boundaries

### In scope (v1.0 / Sprint 7)

- `Workbook(write_only=True)` kwarg with bounded-memory streaming.
- `WriteOnlyWorksheet.append`, `freeze_panes`, `print_area`, `column_dimensions`, `row_dimensions` slots.
- `WriteOnlyCell` with `font`, `fill`, `border`, `alignment`, `number_format`.
- Forbidden-method matrix (random access, post-row mutation).
- `WorkbookAlreadySaved` on second save / post-save append.
- Date/datetime auto-format-attach (mirrors eager-path behaviour).

### Deferred

- Charts, images, merged cells, conditional formatting, data validation, named styles, named ranges, defined names in write-only mode (matches openpyxl).
- Modify-mode streaming (separate problem; modify-mode mutates an existing ZIP and is byte-preserving by design).
- Mixed write-only + eager sheets in one workbook (openpyxl forbids this; wolfxl forbids it too — `write_only=True` makes ALL sheets streaming).
- ~~ZIP streaming straight to destination File (the final `Vec<u8>` materialisation in `package(...)` remains; optional v1.5 bonus).~~ **Landed in v1.5** — `package_to<W: Write + Seek>(...)` and `emit_xlsx_to<W: Write + Seek>(...)` stream the archive directly into a `BufWriter<File>` opened by `save_once`. The buffered `package(...)` / `emit_xlsx(...)` wrappers remain for diff-tool callers. Measured win: ~30 MiB at 1M rows × 5 cols (the size of the compressed archive).
- ~~Per-sheet emit `String` accumulator (the dominant cost during save after v1.5; ~150 MiB at 1M rows × 5 cols).~~ **Landed in v2** — `sheet_xml::emit_streaming_to<W: io::Write>` writes pre-/post-`<sheetData>` markup straight into the open `ZipWriter` entry and `io::copy`s the per-sheet temp file's `<sheetData>` body through. Dispatch lives in `lib.rs::package_emit_ops` via a new `EmitOp` enum (`Bytes` for buffered entries, `SheetStream { sheet_idx }` for streaming sheets); the eager `sheet_xml::emit` path is unchanged for non-streaming sheets and modify mode. **Measured win**: save-time RSS delta dropped from ~150 MiB to **+12.8 MiB at 1M rows × 5 cols** — the residual is libdeflate's compression scratch space and the BufWriter buffer, both `O(1)` in row count. Save-phase memory is now bounded by the SST + styles (irreducible OOXML format costs), exactly as openpyxl's `lxml.xmlfile` model promises. New cargo case `streaming_save_emits_byte_identical_sheet_to_eager` proves the streaming-direct path produces byte-identical `xl/worksheets/sheet1.xml` and `xl/sharedStrings.xml` to the eager path.
- Hyperlinks and comments on `WriteOnlyCell` (factory accepts the kwargs for openpyxl-shape compat but doesn't yet emit them).

## 7. Risks + mitigations

1. **Cell-XML drift between eager and streaming paths.** Both call `emit_row_to`, so divergence requires touching the same function — the byte-parity cargo test (`streaming_byte_equal_to_eager`) and pytest test (`test_byte_parity_with_eager_mode`) catch any future breaks.
2. **Temp-file orphans on SIGKILL.** `wolfxl-stream-{pid}-{sheet_idx}-` prefix in `std::env::temp_dir()` lets admins find them; OS-level temp cleanup reaps them on reboot / weekly cycles.
3. **Memory-bound proof noise at small N.** The bounded-memory probe runs in a subprocess so pytest's accumulated heap is excluded. Calibrated at 100k × 10 cells against an 80 MiB ceiling — actual delta is typically 5-15 MiB.
4. **`column_dimensions` ordering.** OOXML slot 3 (`<cols>`) lives before slot 6 (`<sheetData>`). Setting `freeze_panes` / `print_area` after the first `append` raises `RuntimeError`; we don't try to retroactively splice.
5. **Mixed mode forbidden.** `Workbook(write_only=True)` skips the default `Sheet`. Calling `create_sheet` always returns a `WriteOnlyWorksheet`. There is no mechanism to add an eager `Worksheet` to a write-only workbook.

## 8. Verification

- Cargo: `cargo test --workspace` (15+ test binaries, all green; the 5 new streaming integration tests included).
- Pytest: `uv run pytest -q tests/test_write_only.py` (12/12).
- Oracle: `uv run pytest tests/test_openpyxl_compat_oracle.py -q -k write_only` (2/2 — both probes flipped from xfail/skip to pass).
- Full suite: green.
- Manual smoke: 1M-row × 5-col write completes with peak RSS in single-digit MiB beyond baseline.
