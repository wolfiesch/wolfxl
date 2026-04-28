# RFC-041: Streaming reads (`read_only=True` + auto-trigger on huge sheets)

Status: Shipped <!-- TBD: flip to Shipped + add Pod-β merge SHA when Sprint Ι Pod-β lands -->
Owner: Sprint Ι Pod-β
Phase: Read-side parity (1.3)
Estimate: L
Depends-on: RFC-013 (no — independent of patcher), the existing `CalamineStyledBook` reader.
Unblocks: SynthGL ingest at 1M+-cell scale; downstream LRBench fixtures that timeout under the eager loader.

> **S** = ≤2 days; **M** = 3-5 days; **L** = 1-2 weeks; **XL** = 2+ weeks.

## 1. Problem Statement

KNOWN_GAPS.md Phase 4:

> `openpyxl.load_workbook(path, read_only=True)` + `ws.iter_rows(values_only=True)`
> on 1M-cell sheets — WolfXL accepts the kwarg but reads the full sheet
> into memory. Add a SAX fast path for `read_only=True` or sheets > 50k
> rows.

User code:

```python
wb = wolfxl.load_workbook("huge.xlsx", read_only=True)
ws = wb.active
for row in ws.iter_rows(values_only=True, max_row=100):
    print(row)
# Today: peaks at ~2 GB on a 1M-row sheet because the entire <sheetData>
# is materialised before iter_rows yields its first tuple. openpyxl's
# read_only path streams a row at a time and stays under 200 MB.
```

**Target behaviour**:

* **Honour `read_only=True`** — no in-memory `BTreeMap` of cells; a
  `quick_xml::Reader` scans `<sheetData>` row-by-row and emits one
  Python tuple per `<row>` element. Memory profile: O(row width), not
  O(sheet).
* **Auto-trigger** when the user did NOT pass `read_only` and any
  sheet exceeds a threshold (50k rows OR 5M cells, picked
  heuristically from openpyxl precedent). The auto-trigger is a
  warning, not silent: a one-time `RuntimeWarning` per workbook
  surfaces the decision so users can opt out via `read_only=False`
  or a config flag.

## 2. OOXML Spec Surface

| Spec section | Element | Notes |
|---|---|---|
| §18.3.1.99 CT_Worksheet | `<worksheet>` | Root element. We skip the head (`<sheetPr>`, `<dimension>`, `<sheetViews>`, `<sheetFormatPr>`, `<cols>`) on the streaming path — those are already parsed for side-effects in the eager loader and cached at workbook open. |
| §18.3.1.80 CT_SheetData | `<sheetData>` | Contains `<row>` children only. The streaming reader yields one Python row per `<row>` event. |
| §18.3.1.73 CT_Row | `<row>` | One row; `r="…"`, `ht="…"`, `customHeight="…"`, etc. attrs. Children: `<c>` per cell. |
| §18.3.1.4 CT_Cell | `<c>` | `r="A1"` etc. Cell value via `<v>`, formula via `<f>`, inline string via `<is>`, shared-string ref via `t="s"`. |
| §18.4.9 CT_SharedStringTable | `xl/sharedStrings.xml` | Loaded once at workbook open (not streamed) — necessary because shared-string refs in `<c t="s"><v>N</v></c>` resolve via index. Bounded by the workbook's distinct-string count, typically ≤100k. |

Streaming the `<sheetData>` body is straightforward; the trick is
the head: `<dimension>` carries the row/col bounding box and is
needed by `Worksheet.max_row` / `max_column`. We snapshot it at
open and DO NOT update it as the streaming reader walks (no
`max_row` drift; modify-mode is out of scope here anyway).

## 3. openpyxl Reference

`openpyxl/reader/excel.py:140-205` for `load_workbook(read_only=True)`:

* Returns a `ReadOnlyWorkbook` whose worksheets are `ReadOnlyWorksheet`
  instances.
* `ReadOnlyWorksheet.iter_rows` calls
  `WorkSheetReader.parse_row` on demand — one `<row>` SAX event yields
  one row of `ReadOnlyCell`s.
* Cell values lazy-decode shared strings (cached) and number formats
  (cached). No styles materialise unless requested.
* Exposes `max_row` from the source `<dimension>` element rather than
  walking `<sheetData>`.

We **do not** copy:

* openpyxl's `ReadOnlyCell` class hierarchy — wolfxl returns plain
  Python tuples (`values_only=True`) or `Cell` instances bound to the
  streaming row's row index (`values_only=False`).
* The `lxml`-specific iterparse incremental tree pruning. quick-xml is
  the SAX engine; pruning is implicit (we never build a tree).

## 4. WolfXL Surface Area

### 4.1 Python coordinator

`python/wolfxl/__init__.py::load_workbook` already accepts
`read_only=True`. RFC-041 wires it to a new `_from_streaming_reader`
class method on `Workbook` that:

* Constructs a `CalamineStyledBook` with the new `streaming=True`
  flag (head-only parse — sharedStrings, styles, defined names, but
  NOT `<sheetData>`).
* Lazily binds each `Worksheet` proxy to a `StreamingReader` handle
  that owns a `quick_xml::Reader` over the sheet's bytes.

`python/wolfxl/_worksheet.py`:

* `iter_rows(values_only=True, …)` dispatches to the streaming
  handle when `wb._streaming` is set; falls through to the existing
  eager path otherwise.
* `__getitem__` and other random-access operations raise
  `RuntimeError("read_only=True does not support random access")` —
  matches openpyxl.

### 4.2 Streaming reader (Rust)

New module `crates/wolfxl-reader/src/streaming.rs`:

```rust
pub struct StreamingReader<'a> {
    reader: quick_xml::Reader<&'a [u8]>,
    shared_strings: &'a [String],
    next_row: u32,
}

impl<'a> StreamingReader<'a> {
    pub fn next_row(&mut self) -> Option<Vec<CellValue>> { … }
}
```

PyO3 wraps `StreamingReader` as an iterator that yields one Python
row tuple per `__next__` call.

### 4.3 Native writer / patcher

Out of scope. Streaming applies only to the read path. Modify mode
on a workbook opened with `read_only=True` raises (matches openpyxl).

## 5. Implementation Sketch

1. **Auto-trigger heuristic**. At `load_workbook` time, peek the
   `<dimension>` element of every sheet (cheap, head-only XML). If
   any sheet's `(max_row, max_col)` exceeds the threshold (default
   50k rows OR 5M cells), set `wb._streaming = True` and warn:

   ```
   RuntimeWarning: WolfXL auto-enabled streaming reads for "huge.xlsx"
   (sheet 'Sheet1': 1,200,000 rows × 8 cols). Pass read_only=False to
   opt out.
   ```

2. **Streaming sheet parse**. The `quick_xml::Reader` walks the sheet
   from the start of `<sheetData>` to its close, yielding one row per
   `<row>` end event. Events between rows (e.g. `<mergeCells>`,
   `<hyperlinks>`) are skipped — the streaming path doesn't need
   them; eager loaders that DO need them are unaffected.

3. **`values_only=True` cell decode**:
   * `<c><v>1.5</v></c>` → `1.5`
   * `<c t="s"><v>3</v></c>` → `shared_strings[3]`
   * `<c t="b"><v>1</v></c>` → `True`
   * `<c><f>A1+1</f><v>2</v></c>` → `2` (cached value); `'=A1+1'`
     under `data_only=False`
   * `<c t="inlineStr"><is>...</is></c>` → flat `str` (no rich-text in
     streaming mode; rich-text users opt out via `read_only=False`).

4. **Empty rows**. `<row r="5">` with no `<c>` children yields
   `(None,) * max_col` — matches openpyxl's "missing cell ⇒ None"
   contract.

5. **Stop after `max_row`**. `iter_rows(max_row=100)` short-circuits
   the SAX scan; the reader stops parsing once `next_row > max_row`.

## 6. Verification Matrix

1. **Rust unit tests** — `crates/wolfxl-reader/tests/streaming.rs`
   covering empty sheets, sparse rows, shared-string refs, formula
   cells, inline strings, and a 50k-row golden.
2. **Python round-trip** — `tests/test_streaming_reads.py` reads a
   100k-row fixture and asserts the row count, the first/last row
   values, and a peak-RSS upper bound (skipped under CI without
   `psutil`).
3. **openpyxl parity** — `tests/parity/test_streaming_parity.py`
   yields rows from both libraries and compares element-wise.
4. **Cross-mode** — `read_only=True` + `data_only=True` exercises
   the cached-value path; tested.
5. **Regression fixture** — `tests/fixtures/streaming_huge.xlsx` (a
   compressed 1M-row file checked in via Git LFS or generated at
   test setup).
6. **LibreOffice cross-renderer** — N/A (read-only RFC).

## 7. Cross-Mode Asymmetries

* `read_only=True` workbooks raise on `wb.save()`, `cell.value = …`,
  `ws.append(...)`, `wb.create_sheet(...)`. Matches openpyxl.
* `read_only=True` + `modify=True` raises `ValueError("read_only and
  modify are mutually exclusive")` at load time.

## 8. Risks

| # | Risk | Likelihood | Impact | Mitigation |
|---|------|------------|--------|------------|
| 1 | quick-xml pull parser leaks memory on malformed `<sheetData>` | low | high | Wrap in `Result<…, ReadError>`; bound buffer growth; fuzz a corrupt-stream corpus |
| 2 | Auto-trigger threshold is wrong for some workloads | medium | medium | Add `WOLFXL_STREAMING_THRESHOLD_ROWS` env override; default is `50000` |
| 3 | Shared-string lookup becomes the bottleneck (it's pre-loaded as `Vec<String>`) | low | medium | Index lookup is O(1); benchmark to confirm |

## 9. Effort Breakdown

| Slice | Estimate | Notes |
|---|---|---|
| Streaming reader (Rust) | 3 days | quick-xml SAX + cell decoder |
| Auto-trigger + warning | 1 day | dimension peek + threshold |
| Python wrapper | 1 day | `Workbook._from_streaming_reader`, `iter_rows` dispatch |
| Tests + fixtures | 2 days | golden + parity + memory profile |
| Docs + release notes | 0.5 day | migration guide |
| **Total** | **~7-8 days** | L-bucket |

## 10. Out of Scope

* **Writing** in streaming mode. openpyxl has `write_only=True`; we
  may add it later but it is NOT part of RFC-041.
* Rich-text reads in streaming mode. Users who need rich text fall
  back to `read_only=False` (RFC-040 owns that path).
* Random-access cell reads (`ws["B7"]`) when streaming is engaged —
  raise per openpyxl.

## 11. Open Questions

| OQ | Question | Status |
|---|---|---|
| OQ-1 | Default threshold (50k rows? 100k? configurable per workbook?) | Pending — Pod-β to set |
| OQ-2 | Should `auto-trigger` warn once per workbook or once per sheet? | Pending — Pod-β to call |
| OQ-3 | Expose `Workbook.streaming` as a public attribute for callers who need to detect the mode? | Pending |

## Acceptance

(Filled in after Pod-β merges.)

- Commit: <!-- TBD: Pod-β commit sha when integrated -->
- Verification: `python scripts/verify_rfc.py --rfc 041` GREEN at <!-- TBD -->
- Date: <!-- TBD -->
