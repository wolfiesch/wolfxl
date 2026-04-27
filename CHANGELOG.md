# Changelog

## wolfxl 1.6.0 (TBD-DATE) — chart construction (8 types, full depth)

User-facing release notes: `docs/release-notes-1.6.md`.

Sprint Μ ("Mu") closes the chart-construction gap that the 1.0–1.5
arc deferred as out-of-scope. Eight 2D chart families ship at full
openpyxl 3.1.x per-type feature depth: `BarChart`, `LineChart`,
`PieChart`, `DoughnutChart`, `AreaChart`, `ScatterChart`,
`BubbleChart`, `RadarChart`, plus `Reference`, `Series`, and
`Worksheet.add_chart(chart, anchor)`. The 3D / Stock / Surface /
ProjectedPieChart variants ship as `_make_stub` classes raising
`NotImplementedError` with a v1.6.1-pointer message and land in
v1.6.1. Pivot-chart linkage depends on Sprint Ν / v2.0.0 pivot
tables.

### Added

- **RFC-046 — Chart construction** (Sprint Μ Pod-α/β,
  <!-- TBD: SHA -->). `wolfxl.chart.{Bar,Line,Pie,Doughnut,Area,Scatter,Bubble,Radar}Chart`
  are real classes replacing the `_make_stub` definitions at
  `python/wolfxl/chart/__init__.py`. `Reference(ws, min_col, min_row,
  max_col, max_row)` and `Series(values, categories=None,
  title=None, title_from_data=False)` ship together. Per-type unique
  features land at full openpyxl depth: bar `gap_width`/`overlap`/
  `grouping`/`bar_dir`; line `smooth`/`up_down_bars`/`drop_lines`/
  `hi_low_lines`/per-series `marker`; pie `vary_colors`/
  `first_slice_ang`; doughnut `hole_size`; area `grouping`/
  `drop_lines`; scatter `scatter_style`; bubble `bubble_3d`/
  `bubble_scale`/`show_neg_bubbles`/`size_represents`; radar
  `radar_style`. `Worksheet.add_chart(chart, anchor)` accepts a
  coordinate string (one-cell anchor at the top-left of that cell)
  or one of the RFC-045 anchor helper classes (`OneCellAnchor`,
  `TwoCellAnchor`, `AbsoluteAnchor`).
- **`copy_worksheet` chart deep-clone with cell-range re-pointing**
  (Sprint Μ Pod-γ, RFC-035 §10 lift, <!-- TBD: SHA -->). The
  deferred limit at `Plans/rfcs/035-copy-worksheet.md` lines 924-929
  is lifted. Charts in copied sheets now deep-clone with cell-range
  re-pointing on every series. Self-references
  (`SourceSheet!$B$2:$B$10`) get rewritten to the copy's sheet name;
  cross-sheet references (`Other!$A$1:$A$5`) are preserved verbatim.
  Cached values inside `<c:strCache>` / `<c:numCache>` are preserved
  as-is (Excel rebuilds them on next open if stale). The new
  behaviour is the default; there is no opt-in flag.
- **Modify-mode `add_chart`** (Sprint Μ Pod-γ, <!-- TBD: SHA -->).
  `wb = load_workbook(..., modify=True); ws.add_chart(chart, "G2");
  wb.save()` works. New patcher Phase 2.5l drains the queued chart
  adds per sheet, allocates fresh `chartN.xml` and `drawingN.xml`
  numbers via `PartIdAllocator`, emits chart bytes through
  `file_adds`, and splices `<xdr:graphicFrame>` blocks into the
  sheet's existing drawing part (or creates one if the sheet had no
  drawing yet). Composes cleanly with RFC-045 `add_image` and
  RFC-035 `copy_worksheet`.

### Internal / infra

- **`RT_CHART` const** added to `crates/wolfxl-rels/src/lib.rs`:
  `pub const RT_CHART: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";`.
- **`crates/wolfxl-writer/src/model/chart.rs`** — typed model.
  `ChartKind` enum (8 variants), `Chart`, `Series`, `Reference`,
  axis types, marker, layout, legend, title.
- **`crates/wolfxl-writer/src/emit/charts.rs`** — per-type emit
  (~1500 LOC). One `emit_chart_xml(&Chart, &mut Writer)` entry
  point dispatched on `chart.kind` to per-type emit fns
  (`emit_bar_chart`, `emit_line_chart`, …). Shared helpers for axes
  and legend live alongside.
- **`crates/wolfxl-writer/src/emit/drawings.rs`** extended (NOT
  rewritten) for `<xdr:graphicFrame>` (charts) alongside the
  existing `<xdr:pic>` (RFC-045 images). A worksheet with both
  shares a single drawing part with one `<xdr:pic>` block and one
  `<xdr:graphicFrame>` block.
- **`PartIdAllocator` extended for charts** (Pod-γ). The centralized
  allocator at `crates/wolfxl-rels/src/part_id_allocator.rs` (RFC-035
  §5.2) gains `alloc_chart()` and reuses `alloc_drawing()`, so
  multiple `add_chart` calls in the same save plus concurrent
  RFC-045 `add_image` / RFC-035 sheet copies all get collision-free
  numbers.
- **`XlsxPatcher` Phase 2.5l** — `chart_adds` drain in modify mode.
  Sequenced after Phase 2.5j images / 2.5g comments / 2.5f tables
  and before 2.5c content-types aggregation.
- **PyO3 binding `Workbook.add_chart_native(sheet_idx, chart_payload,
  anchor_dict)`** in `src/lib.rs`.
- **10 new entries in `tests/parity/openpyxl_surface.py`
  `_GAP_ENTRIES`** (Pod-δ): one per chart class + `Reference` +
  `Series` + `Worksheet.add_chart`. Integrator flips them to
  `wolfxl_supported=True` post-merge with the `shipped-1.6` tag.

### Documentation

- `Plans/rfcs/046-chart-construction.md` — Sprint Μ Pod-ε RFC.
- `Plans/rfcs/INDEX.md` bumped 22 → 23 RFCs (RFC-046 row added; DAG
  extended with Phase 5 1.6 deliverable).
- `tests/parity/KNOWN_GAPS.md` — 1.6 roadmap entry added; "Chart
  construction" lifted from "Out of scope" to "Closed in 1.6 (Sprint
  Μ)"; 3D / Stock / Surface / ProjectedPie deferral to 1.6.1
  documented; pivot-chart linkage explicitly deferred to Sprint Ν /
  v2.0.0.
- `docs/release-notes-1.6.md` — user-facing 1.6 release notes
  (Sprint Μ Pod-ε).
- `Plans/rfcs/035-copy-worksheet.md` — chart-aliasing limit at lines
  924-929 marked `~~strikethrough~~` with a "✅ Lifted in 1.6 (Sprint
  Μ Pod-γ)" pointer to RFC-046 §7.

### Test totals (post-1.6)

- `cargo test --workspace --exclude wolfxl`: ~660 + N green
  (Pod-α adds per-type chart-emit tests, axis-id-allocation tests).
- `pytest tests/`: **~1235+ → ~1300+ passed** (Pod-β adds ~40
  write-mode tests, Pod-γ adds ~25 modify-mode + chart-deep-clone
  tests; final count filled in on integrator merge).
- `pytest tests/parity`: **~165+ → ~190+ passed** (Pod-δ adds ~25
  parity tests).

## wolfxl 1.5.0 (2026-04-26) — encrypted writes + image construction + streaming-datetime fix

User-facing release notes: `docs/release-notes-1.5.md`.

Sprint Λ ("Lambda") closes the last two "construction" gaps that the
1.0–1.4 arc deferred as out-of-scope: write-side OOXML encryption
(Pod-α, RFC-044) and image construction (Pod-β, RFC-045). Pod-γ
closes the streaming-reads datetime divergence Pod-β surfaced in 1.3.
After 1.5 the openpyxl-parity surface is exhausted at the
construction level too — only chart construction (v1.6.0 / Sprint Μ)
and pivot table construction (v2.0.0 / Sprint Ν) remain.

### Added

- **RFC-044 — Write-side OOXML encryption** (Sprint Λ Pod-α,
  feat `4bc806c`). `Workbook.save(path, password="...")` now
  emits an Agile (AES-256 / SHA-512) encrypted file via
  `msoffcrypto-tool`'s high-level `OOXMLFile.encrypt()`. Lifts the
  `NotImplementedError` at `python/wolfxl/_workbook.py:1032`. Empty
  string is a literal empty-key password (NOT equivalent to
  `password=None`). Encryption pass is mode-agnostic — works for
  both write-mode and modify-mode workbooks. The
  `wolfxl[encrypted]` extra now covers writes too (was already on
  reads from Sprint Ι Pod-γ).
- **RFC-045 — Image construction** (Sprint Λ Pod-β,
  writer `0ace8c5` + Image+add_image `d9cb569` + tests `a73737e`). `wolfxl.drawing.image.Image(...)` and
  `Worksheet.add_image(...)` are real, replacing the `_make_stub` at
  `python/wolfxl/drawing/image.py`. Supports PNG / JPEG / GIF / BMP;
  one-cell, two-cell, and absolute anchors. Magic-byte format
  sniffing + auto-detected width/height from format-specific
  headers (PNG IHDR, JPEG SOF, GIF LSD, BMP DIB). Both write mode
  (native writer emits drawingN.xml + media + rels through a new
  images-emit pass) and modify mode (patcher's new Phase 2.5j
  drains queued images and routes through `file_adds`). New anchor
  helper classes under `wolfxl.drawing.spreadsheet_drawing` and
  `wolfxl.drawing.xdr` to mirror openpyxl's module layout. Composes
  cleanly with RFC-035 `copy_worksheet`.
- **Streaming-datetime fix** (Sprint Λ Pod-γ, fix `98cd147`).
  `iter_rows(values_only=True)` and `StreamingCell.value` now return
  `datetime` objects for date-formatted cells under
  `read_only=True`. Closes the documented Phase 4 divergence in
  `tests/parity/KNOWN_GAPS.md` lines 116-122. The streaming reader
  consults the styles table for the cell's number format and
  converts Excel serial floats inline.

### Internal / infra

- **`PartIdAllocator` extended for images** (Pod-β). The centralized
  allocator at `crates/wolfxl-rels/src/part_id_allocator.rs`
  (introduced in RFC-035 §5.2) gains `alloc_image(extension: &str)`
  with per-extension counters, so multiple `add_image` calls in the
  same save plus concurrent RFC-035 sheet copies all get
  collision-free numbers.
- **Phase 2.5j patcher hook** (Pod-β). New per-sheet phase in
  `XlsxPatcher::do_save`, sequenced after Phase 2.5g comments /
  2.5f tables and before 2.5c content-types aggregation. Drains
  `queued_image_adds` per sheet, builds or extends drawing parts,
  and emits new image bytes through `file_adds`.
- **`_save_plaintext` bytes-output seam** (Pod-α). Symmetric with
  Sprint Κ Pod-β's bytes-input plumbing — the existing writer /
  patcher pipeline now accepts `Path | str | BinaryIO` for the
  output target, enabling the encryption pass to buffer plaintext
  bytes through `BytesIO` before writing the encrypted envelope.
- **uv.lock updated** for `msoffcrypto-tool >= 5.4` write-side
  surface (no new transitive deps; the read-side already pulled
  the package in 1.3 Sprint Ι).

### Documentation

- `Plans/rfcs/044-encryption-writes.md` — Pod-α RFC.
- `Plans/rfcs/045-image-construction.md` — Pod-β RFC.
- `Plans/rfcs/INDEX.md` bumped 20 → 22 RFCs.
- `tests/parity/KNOWN_GAPS.md` — "Writing encrypted xlsx" and
  "Image construction" lifted from "Out of scope" to "Closed in
  1.5"; chart construction explicitly scheduled for v1.6.0,
  pivot table construction for v2.0.0; Pod-γ closes the streaming-
  datetime divergence note.
- `docs/release-notes-1.5.md` — user-facing 1.5 release notes
  (Sprint Λ Pod-δ).

### Test totals (post-1.5)

- `cargo test --workspace --exclude wolfxl`: ~660 + N green
  (Pod-β adds image-meta sniffer + dim-extraction tests; Pod-γ adds
  styles-table-aware streaming-datetime tests).
- `pytest tests/`: **1175+ → ~1235+ passed** (Pod-α/β/γ each add
  cases; final count filled in on integrator merge).
- `pytest tests/parity`: **140+ → ~165+ passed**.

## wolfxl 1.4.0 (2026-04-26) — `.xlsb` / `.xls` reads + bytes / BytesIO / file-like input

User-facing release notes: `docs/release-notes/1.4.md`.

Sprint Κ ("Kappa") closes Phase 5 — the last open row in
`tests/parity/KNOWN_GAPS.md`. Four parallel pods landed runtime-dispatched
calamine backends for `.xlsb` and `.xls`, a unified bytes / `io.BytesIO`
/ file-like input path on `load_workbook` across all formats, and the
matching parity fixtures vs `pandas.read_excel(engine="calamine")`.

### Added

- **RFC-043 — `.xlsb` / `.xls` reads via runtime-dispatched calamine
  backends** (Sprint Κ Pod-α, `b805aac`). New `CalamineXlsbBook` and
  `CalamineXlsBook` Rust pyclasses dispatched at `load_workbook` time
  by the new `_rust.classify_file_format(path_or_bytes)` magic-byte sniffer.
  Reads return values + cached formula results; style accessors
  (`cell.font` / `.fill` / `.border` / `.alignment` / `.number_format`)
  raise `NotImplementedError` on non-xlsx workbooks. `Workbook._format`
  attribute (`'xlsx' | 'xlsb' | 'xls'`) for caller-side branching.
- **Bytes / `BytesIO` / file-like input on `load_workbook`** (Sprint Κ
  Pod-β, `ddf0dc5`). Each backend exposes a `Source` enum with
  `File(BufReader<File>)` and `Bytes(Cursor<Vec<u8>>)` arms that
  uniformly implement `Read + Seek`. Replaces Sprint Ι Pod-γ's
  tempfile workaround for password reads — decrypted bytes now route
  through `open_from_bytes` end-to-end.
- **`.xlsb` / `.xls` parity fixtures + assertions vs
  `pandas.read_excel(engine="calamine")`** (Sprint Κ Pod-γ,
  `97585a5` (fixtures) + `49e95d5` (parity tests)). New `tests/parity/test_xlsb_reads.py` +
  `tests/parity/test_xls_reads.py` element-wise pin shape + values +
  cached formula results.

### Internal / infra

- **Parity ratchet flipped on Phase 5 entries**. The two open
  `_GAP_ENTRIES` rows in `tests/parity/openpyxl_surface.py`
  (`openpyxl.load_workbook('foo.xlsb')` /
  `openpyxl.load_workbook('foo.xls')`) flip to
  `wolfxl_supported=True` and are tagged `shipped-1.4` post-merge by
  the integrator. The xfail strict pins in
  `tests/parity/test_surface_smoke.py` flag the flip is required.
- **KNOWN_GAPS Phase 5 section removed**. The openpyxl-parity roadmap
  is exhausted; only out-of-scope items (write-side encryption,
  OpenDocument, charts/pivots/images) remain.

### Documentation

- `Plans/rfcs/043-xlsb-xls-reads.md` — RFC for the .xlsb / .xls reads
  slice (Sprint Κ Pod-δ).
- `Plans/rfcs/INDEX.md` bumped 19 → 20 RFCs.
- `tests/parity/KNOWN_GAPS.md` — Phase 5 row replaced with a
  "✅ SHIPPED in 1.4" note; new "Roadmap status" overview at the top.
- `docs/release-notes/1.4.md` — user-facing 1.4 release notes
  (Sprint Κ Pod-δ).

### Test totals (post-1.4)

- `cargo test --workspace --exclude wolfxl`: ~660 + N green (Pod-α
  adds magic-byte sniffer + bytes-input round-trip tests).
- `pytest tests/`: **1106 → ~1175+ passed** (Pod-α/β/γ each add
  cases; final count filled in on integrator merge).
- `pytest tests/parity`: **102 → ~140+ passed**.

## wolfxl 1.3.0 (2026-04-26) — Read-side parity (rich text + streaming + password)

User-facing release notes: `docs/release-notes-1.3.md`.

Sprint Ι ("Iota") closes the three highest-impact read-side gaps from
KNOWN_GAPS Phase 2/3/4 and lifts the implicit T3 rich-text-write
deferral. Four parallel pods landed; the parity ratchet is re-engaged
with the closed rows flipped to `wolfxl_supported=True` and tagged
`shipped-1.3`.

### Added

- **RFC-040 — Rich-text reads + writes (round-trip)** (Sprint Ι Pod-α,
  `381813a`). New `python/wolfxl/cell/rich_text.py` ships
  `CellRichText`, `TextBlock`, `InlineFont` shims that match openpyxl's
  iteration / equality / constructor contract without pulling openpyxl
  as a runtime dep. `Cell.rich_text` always exposes structured runs
  (or `None` for plain cells). `Cell.value` flips to `CellRichText`
  only under `load_workbook(rich_text=True)`, matching openpyxl's
  flag-gated default. Inline-string emit on write
  (`<c t="inlineStr"><is>...</is></c>`) sidesteps SST mutation
  entirely; `crates/wolfxl-writer/src/rich_text.rs` is the single
  source of truth for run grammar (parse + emit, 13 cargo unit tests).
  Round-trip verified wolfxl→openpyxl, openpyxl→wolfxl, wolfxl→wolfxl
  via 24 new pytest cases.
- **RFC-041 — SAX streaming reads (values + styles)** (Sprint Ι Pod-β,
  `75de628`). New `src/streaming.rs` Rust module + `python/wolfxl/_streaming.py`
  expose a true SAX path activated via `load_workbook(read_only=True)`
  or auto-triggered for sheets > 50000 rows. `iter_rows()` yields
  read-only `StreamingCell` proxies with full
  `.font/.fill/.border/.alignment/.number_format` access (lazy O(1)
  styles-table lookup). `iter_rows(values_only=True)` yields plain
  tuples. Mutation centralized in `__setattr__` raises
  `RuntimeError` immediately. Benchmark on 100k-row × 10-col
  fixture: wolfxl `read_only=True` is **~5.7× faster** than openpyxl
  `read_only=True` on wall time (0.700 s vs 4.017 s), and ~2× faster
  than wolfxl's bulk-FFI eager path. 22 new pytest cases (16
  streaming + 5 parity + 1 modified workbook-compat assertion).
- **RFC-042 — Password-protected reads via msoffcrypto-tool** (Sprint
  Ι Pod-γ, `f0ea2d1`). New `password=` kwarg on
  `wolfxl.load_workbook(...)` lazy-imports
  `msoffcrypto-tool` (optional dep, install via
  `pip install wolfxl[encrypted]`) and dispatches the decrypted bytes
  through a tracked tempfile to the existing path-based readers.
  Tempfile cleaned up by `Workbook.close()`. Modify mode + password
  works: `load_workbook(path, password=..., modify=True)` →
  mutate → `wb.save(out)` emits plaintext. Write-side encryption
  out of scope (raises `NotImplementedError` if `password=` passed
  to `save()`). 9 new pytest cases.
- **Workbook.defined_names `__setitem__`** (Sprint Ι Pod-δ D4,
  `b64c364`). Closes KNOWN_GAPS Phase 1 row.
  `wb.defined_names["MyName"] = DefinedName(name="MyName",
  attr_text="Sheet1!$A$1")` now routes through the existing Rust
  `add_named_range` path with Excel-compliant name validation
  (no whitespace, no leading digit, not an A1 ref, not R/C R1C1
  reserved tokens). Sheet-scope names (`localSheetId` set) route via
  `scope=sheet` plus the resolved sheet name. 11 new pytest cases.

### Fixed

- **Native-writer VML margin honors per-column widths** (Sprint Ι
  Pod-δ D3, `92c901d`).
  `crates/wolfxl-writer/src/emit/drawings_vml.rs::compute_margin` was
  hard-coding `COL_WIDTH_PT = 48.0`; sheets with custom column widths
  rendered comment popups over the wrong cell area. New
  `compute_margin_with_widths` walks `worksheet.columns` and sums
  per-column widths in points, mirroring the modify-mode patcher's
  helper. Empty `<cols>` falls back to the legacy math so existing
  fixtures stay byte-stable. 3 new Rust unit tests + 2 Python
  round-trip tests. Closes
  `Plans/followups/native-writer-vml-margin-fix.md`.

### Internal / infra

- **Parity ratchet re-engaged** (Sprint Ι Pod-δ D1 `751760f` +
  integrator flip-up `71d1d4f`). `tests/parity/openpyxl_surface.py`
  now carries five fine-grained `_GAP_ENTRIES` rows. The three
  Sprint-Ι-closed rows (rich text, streaming, password) are flipped
  to `wolfxl_supported=True` and tagged `shipped-1.3`. The two
  remaining Phase-5 rows (`.xls` / `.xlsb`) keep
  `wolfxl_supported=False` and continue to xfail strictly via
  `test_known_gap_still_gaps` so a future closer flips the test
  green.
- **Custom pytest marks registered** (Sprint Ι Pod-δ D2, `ce9dda3`).
  `rfc035`, `rfc031`, `rfc036`, and `manual` are added to
  `pyproject.toml`'s `[tool.pytest.ini_options].markers`, silencing
  the recurring `PytestUnknownMarkWarning` noise.
- **uv.lock updated** for msoffcrypto-tool 6.0.0 + olefile + pycparser
  transitive deps (introduced as an optional dep, not a runtime dep).

### Documentation

- `Plans/rfcs/040-rich-text.md` — Pod-α RFC (~210 lines).
- `Plans/rfcs/041-streaming-reads.md` — Pod-β RFC (~225 lines).
- `Plans/rfcs/042-password-reads.md` — Pod-γ RFC (~230 lines).
- `Plans/rfcs/INDEX.md` bumped 16 → 19 RFCs.
- `Plans/followups/native-writer-vml-margin-fix.md` marked Closed.
- `tests/parity/KNOWN_GAPS.md` — Phase 2/3/4 rows moved to "Shipped"
  status; T3 rich-text-write entry retired.
- `docs/release-notes-1.3.md` — user-facing 1.3 release notes
  (~310 lines).

### Test totals (post-1.3)

- `cargo test --workspace --exclude wolfxl`: ~660 green.
- `pytest tests/`: **1106 passed, 15 skipped, 2 xfailed** (1.2 had
  0 xfails; 1.3 adds 2 from the parity ratchet pinning .xls/.xlsb
  Phase-5 deferred items).
- 100k-row × 10-col streaming benchmark: ~5.7× faster than openpyxl
  `read_only=True`.

## wolfxl 1.2.0 (2026-04-26) — RFC-035 follow-ups + composition hardening

User-facing release notes: `docs/release-notes-1.2.md`.

Sprint Θ ("Theta") closes every RFC-035 follow-up that 1.1 deferred,
landing in four parallel pods (A, B, C, D). The `copy_worksheet`
surface now ships zero `xfail(strict=True)` markers in
`tests/test_copy_worksheet_modify.py`.

### Added

- **RFC-035 §3 OQ-a — Write-mode `copy_worksheet`** (Sprint Θ Pod-C1, `46862b9`).
  `Workbook.copy_worksheet(source, name=None)` now works in pure
  write mode (Workbook() + create_sheet + copy_worksheet, never
  loaded from disk). Walks the in-memory `NativeWorkbook` model and
  clones every sub-record (cells, styles, tables, DV, CF,
  hyperlinks, comments, defined names) into a fresh sheet appended
  at end. New tests under `tests/test_copy_worksheet_write_mode.py`
  (7 cases). The §3 OQ-a `NotImplementedError` is gone.
- **RFC-035 §5.3 — Image deep-clone via `wb.copy_options`**
  (Sprint Θ Pod-C2, `89fb68f`). New `CopyOptions` dataclass exposed
  as `wb.copy_options` with the `deep_copy_images: bool` field
  (default `False`, preserving 1.1's alias behaviour). When set
  before `copy_worksheet()`, the planner duplicates `xl/media/
  imageN.png` parts and re-points the cloned drawing rels at the
  fresh `xl/media/imageM.png`. The flag is snapshot at queue time
  so toggling between calls produces clones with different image
  strategies in the same save. New tests under
  `tests/test_copy_worksheet_deep_clone_images.py` (4 cases).
- **RFC-035 §10 — `xl/calcChain.xml` rebuild** (Sprint Θ Pod-C3,
  `d6524c2`). The patcher gains a Phase 2.8 walk that scans every
  sheet's post-mutation XML for `<f>` cells and emits the matching
  `xl/calcChain.xml`; the native writer mirrors the behaviour at
  write time. Workbooks with zero formulas omit the part entirely.
  External readers that consume `calcChain.xml` directly no longer
  see a stale chain after `copy_worksheet`. New tests under
  `tests/test_calcchain_rebuild.py` (7 cases). New module
  `src/wolfxl/calcchain.rs` and crate-level emitter
  `crates/wolfxl-writer/src/emit/calc_chain_xml.rs`.
- **`load_workbook(..., permissive=True)` mode** (Sprint Θ Pod-A,
  `c6f94fc`). New opt-in flag for slightly-malformed workbook.xml
  inputs. When the `<sheets>` block is empty / self-closing, the
  loader walks `xl/_rels/workbook.xml.rels` to discover worksheet
  targets, synthesises titles (`Sheet1`, `Sheet2`, …), and rewrites
  the in-memory workbook.xml so downstream phases see a well-formed
  document. The flag defaults to `False`; well-formed inputs are
  unaffected. Closes RFC-035 KNOWN_GAPS bug #4
  (`test_p_self_closing_sheets_block`).

### Fixed

- **RFC-035 KNOWN_GAPS bug #6** (Sprint Θ Pod-B, `b27d177`).
  Replaces the naive byte-substring `</sheets>` locator in
  `splice_into_sheets_block` with a `quick-xml` SAX scan that
  respects element nesting. Comments, CDATA sections, and PIs
  containing the literal `</sheets>` token can no longer fool the
  splice. Five new Rust unit tests in `src/wolfxl/mod.rs::rfc013_tests`
  pin the invariant (normal, self-closing, comment fakeout, CDATA
  fakeout, malformed). As a side-fix, the test fixture helper
  `_inject_comment_with_sheets_token` was anchoring on `?>` (XML
  declaration close) — but openpyxl-saved workbooks omit the XML
  decl, so the injection was a silent no-op that masked the bug.
  Helper now anchors on `<workbook ...>` via regex.

### Documentation

- `tests/parity/KNOWN_GAPS.md` — Sprint Θ Pod-D1 reconciled the
  stale "Modify mode — T1.5-deferred features" table. Phase 3
  shipped every entry (RFC-020/021/022/023/024/025/026); the table
  is now an audit log of ✅ Shipped rows pointing at the per-RFC
  modify-mode test files. Bug #4 and #6 moved from "Deferred to
  1.2" to "Fixed in 1.2 (Sprint Θ Pods A + B)". The "Deferred queue
  (post-1.2)" section is now empty.
- `Plans/rfcs/035-copy-worksheet.md` — new §8.5 "Sprint Θ
  deliverables (1.2)" subsection populated with per-pod commit
  references.
- `docs/release-notes-1.2.md` — user-facing 1.2 release notes
  (~210 lines) covering every Sprint Θ deliverable + carry-forward
  limitations.

### Test totals (post-1.2)

- `cargo test --workspace --exclude wolfxl`: ~648 green.
- `pytest tests/`: **1038 passed, 16 skipped, 0 xfailed** (1.1 had
  2 xfails; both flipped to PASS in 1.2).
- `pytest tests/parity`: **102 passed, 2 skipped, 0 failures**.

## wolfxl 1.1.0 (2026-04-26) — Full structural parity

User-facing release notes: `docs/release-notes-1.1.md`.

### Added

- **RFC-035** — `Workbook.copy_worksheet(source, name=None)` (modify
  mode). Clones an entire sheet subgraph: sheet XML + ancillary parts
  (tables, comments + VML, drawings, hyperlinks, DV, CF) + rels +
  content-types + workbook entry + sheet-scoped defined names. Image
  media is aliased rather than deep-copied per RFC-035 §5.3 (avoids
  50× bloat on workbooks with logos). Tables are auto-renamed
  (`{base}_{N}`, N starts at 2 per RFC-035 §3 OQ-b) so the workbook-
  wide `displayName` uniqueness constraint holds. Sheet-scoped
  defined names (e.g. `_xlnm.Print_Area`) get fresh entries with
  `localSheetId == new_idx`, routed through RFC-021's merger queue.
  See `tests/parity/KNOWN_GAPS.md` "RFC-035 — copy_worksheet
  divergences from openpyxl (SHIPPED 1.1)" for the five places where
  wolfxl deliberately preserves what openpyxl drops. Sprint Ζ
  collaboration: Pod-α (planner crate `wolfxl-structural::sheet_copy`),
  Pod-β (patcher Phase 2.7 + Python coordinator), Pod-γ (19-case
  pytest harness + openpyxl parity + byte-stability + LibreOffice
  cross-renderer), Pod-δ (4 cross-RFC composition bug-fixes +
  `KNOWN_GAPS` cleanup + status flip).
- **RFC-034** — `Worksheet.move_range(cell_range, rows=0, cols=0,
  translate=False)` (modify mode). Paste-style relocation of a
  rectangular block of cells. Formulas inside the moved block are
  paste-translated (`respect_dollar=true` per RFC-012 §5.5):
  relative refs shift by `(rows, cols)`; `$`-marked refs do NOT
  shift. With `translate=True`, formulas in cells outside the
  moved block that reference cells inside the source rectangle
  are also re-anchored. MergeCells / hyperlinks / DV-CF sqref
  pieces fully inside the source rectangle shift with the block;
  pieces straddling the boundary or outside are left in place.
  New crate module: `crates/wolfxl-structural/src/range_move.rs`.
  Patcher Phase 2.5j drain in `src/wolfxl/mod.rs`. Tests:
  `tests/test_move_range_modify.py` (15 cases).
- **RFC-001** — Removed leftover `rust_xlsxwriter` workspace dependency. Native writer (RFC W5 replacement) has been the sole xlsx-write path since Phase 2; the dep was unused.
- **RFC-036** — `Workbook.move_sheet(sheet, offset)` (modify mode).
  Reorders sheets in-place; updates `<sheets>` order in workbook.xml
  and any internal references that depend on tab index.
- **RFC-030** — `Worksheet.insert_rows(idx, amount=1)` /
  `delete_rows(idx, amount=1)` (modify mode). Pure-Rust XML rewrite
  via the new `crates/wolfxl-structural` workspace crate. Shifts
  every cell-coord, dimension, merge, hyperlink, table, DV, CF
  anchor, defined-name, and formula reference touched by the row
  band. Delete tombstones (`#REF!`) are emitted per OOXML semantics.
- **RFC-031** — `Worksheet.insert_cols(idx, amount=1)` /
  `delete_cols(idx, amount=1)` (modify mode). Symmetric to RFC-030
  on the column axis; shares the `wolfxl-structural` crate. Adds the
  col-only `<col>` span splitter (`crates/wolfxl-structural/src/cols.rs`)
  so per-column width / style metadata is preserved across inserts
  and deletes. `idx` accepts either a 1-based int or an Excel
  column letter.

### Fixed

- **RFC-031 round 2** — `<tableColumns>` on `tableN.xml` now
  correctly grows / shrinks when `insert_cols` / `delete_cols`
  overlap the table's column band. Previously the `count="N"`
  attribute and `<tableColumn>` element list stayed at the
  pre-shift size, producing an xlsx that Excel and openpyxl
  refused to load. Fixed by `crates/wolfxl-structural/src/shift_workbook.rs`
  (`extract_table_col_band` + `rewrite_table_columns_block`).
  Regression: `tests/test_col_shift_modify.py::test_rfc031_round2_*`
  (4 cases). Also closes the corresponding action items in
  `Plans/followups/rfc-030-031-api-coordination.md`.
- **RFC-035 cross-RFC composition (Sprint Ζ Pod-δ)** — closed four
  composition defects surfaced by Pod-γ's full harness and tracked in
  `tests/parity/KNOWN_GAPS.md`:
  - `copy_worksheet` + cell edit on the clone in the same `save()`
    no longer raises `OSError: Missing zip entry
    xl/worksheets/sheetN.xml`. Phase 3 reads cloned-sheet bytes from
    `file_adds` / `file_patches` first and routes the rewrite back to
    `file_adds` for cloned paths.
    (`tests/test_copy_worksheet_modify.py::test_i_…`)
  - `copy_worksheet` + `move_sheet(clone_title, …)` in the same
    `save()` no longer drops the cloned `<sheet>` from
    `xl/workbook.xml`. Phase 2.5h's reorder pass now reads from
    `file_patches["xl/workbook.xml"]` so the Phase 2.7 → Phase 2.5h
    handoff happens through the shared `file_patches` map per
    RFC-035 §5.4. (`tests/test_copy_worksheet_modify.py::test_j_…`)
  - `copy_worksheet` + `add_table` on the clone no longer raises
    `OSError`. Same root-cause fix as the cell-edit case, plus a
    `file_adds` / `file_patches` probe in the Phase 2.5f rels-graph
    load. (`tests/test_copy_worksheet_modify.py::test_k_…`)
  - User-queued sheet-scoped defined name colliding on
    `(name, localSheetId)` with the `copy_worksheet` planner's emit
    no longer produces two `<definedName>` entries with the
    planner's value silently winning. Phase 2.7 now skips its push
    when the user has already queued a matching key — user wins
    (per RFC-035 §5.4 and Pod-β's last-write-wins-on-the-USER
    invariant). (`tests/test_copy_worksheet_modify.py::test_q_…`)

### Known issues (RFC-035 — deferred to 1.2)

Two corner cases from Pod-γ's harness remain pinned `xfail` and are
tracked in `tests/parity/KNOWN_GAPS.md`:

- **Self-closing `<sheets/>` workbook.xml fixtures** — wolfxl's
  loader rejects them, so Phase 2.7's self-closing splice branch is
  unreachable through the public API. Real Excel never emits
  `<sheets/>` for a non-empty workbook; reachable only via direct
  ZIP edit. (`test_p_self_closing_sheets_block`)
- **CDATA / processing-instruction fakeout** — a workbook.xml
  comment containing the literal `</sheets>` token can fool the
  byte-level locator for the `<sheets>` block. Acknowledged in
  Pod-β's handoff note as acceptable for 1.1 since no real Excel-
  emitted workbook contains it. (`test_r_cdata_pi_fuzz_fakeout`)

### Notes

- Phase-4a (RFC-030, RFC-031, RFC-036) was dispatched as three
  parallel pods. RFC-031 reconciliation required hand-porting the
  `<col>` splitter onto RFC-030's crate layout — sprint retro now
  documents the rule "if two pods touch a shared crate, sequence
  them, don't parallelize."
- Fuzz / property test for `apply_workbook_shift` added under
  `crates/wolfxl-structural/tests/prop_apply_workbook_shift.rs`
  (5000 deterministic iterations on a small fixture; asserts
  no panic + well-formed XML output).

## wolfxl 0.5.0 (2026-04-20) - PyPI cdylib parity release

### Added

- **`Worksheet.schema()`** and **`Worksheet.classify_format(fmt)`**
  Python methods, plus a module-level `wolfxl.classify_format(fmt)`.
  Both delegate to the bridge added in the previous entry so Python
  callers get byte-compatible answers with `wolfxl schema --format
  json` for the structural fields (name, row count, column names,
  null counts, unique counts, cardinality, samples). See
  "Known divergences" in `tests/test_classifier_parity.py` for the
  two fields that don't yet match (numeric `int` vs `float` and the
  openpyxl-styled format-category gap — both close when the sprint-3
  "Option A" engine-collapse work lands).
- **`infer_sheet_schema(rows, name, number_formats=None)`** — the
  bridge now accepts an optional parallel `List[List[Optional[str]]]`
  of per-cell `number_format` strings. Without it, Python's inferred
  `format_category` would silently drift to `"general"` for every
  column because `py_to_cell` couldn't see format metadata. The
  Python `Worksheet.schema()` passes formats from
  `iter_cell_records(include_format=True)` so both surfaces see the
  same format context going into `wolfxl_core::infer_sheet_schema`.
- **`tests/test_classifier_parity.py`** (~240 LOC) — cross-surface
  drift test. Runs `cargo run --quiet --release -p wolfxl-cli --
  schema <fixture> --format json` as a subprocess and compares the
  result to `Worksheet.schema()` on the same workbook. Four cases:
  structural parity (row/column counts + names + null/unique/cardinality),
  sample-list parity (as multisets), direct `classify_format`
  round-trip over every `FormatCategory` variant, and
  `Worksheet.classify_format` → module-level identity.

### Notes

- **Task #22b ships net-new Python surface, not replacements.** The
  sprint-2 plan described #22b as "replace duplicate classifiers in
  `calamine_styled_backend.rs`" — but inspection showed the cdylib
  doesn't actually duplicate any classifier logic; it just returns
  raw `number_format` strings. The "authoritative classifier" work
  was already fully in `wolfxl-core`. So #22b collapsed to exposing
  the bridge methods on the Python surface and adding the parity
  test, both of which were also in the sprint plan's scope.
- **Parity test is the drift detector going forward.** Any future
  change to either the CLI's schema output or Python's
  `worksheet.schema()` that breaks structural parity (null counts,
  unique counts, column names, etc.) fails CI immediately. The
  narrower format-category and int/float parity will tighten when
  Option-A collapses the reader paths.

### Core Bridge Groundwork

#### Added

- **`wolfxl_core_bridge` PyO3 module** (new `src/wolfxl_core_bridge.rs`,
  ~260 LOC). Exposes three `wolfxl-core` classifiers on the `_rust`
  extension module:
  - `classify_format(fmt: str) -> str` — thin wrapper on
    `wolfxl_core::classify_format`, returns the category string
    (`"general"`, `"currency"`, `"date"`, ...) that `wolfxl schema
    --format json` emits in the `format` field.
  - `classify_sheet(rows: List[List[Any]], name: str = "Sheet1") -> str`
    — returns the sheet-class string (`"empty"`, `"readme"`,
    `"summary"`, `"data"`) that `wolfxl map --format json` emits in
    the `class` field.
  - `infer_sheet_schema(rows, name = "Sheet1") -> dict` — returns the
    per-column schema dict in the same shape as `wolfxl schema
    --format json`, minus the outer `"sheets"` wrapper.
- **Native Python input coercion** in the bridge: `None` / `bool` /
  `int` / `float` / `datetime.datetime` / `datetime.date` /
  `datetime.time` / `str` all map to their `CellValue` counterparts.
  Unknown types fall back to `str()` so the bridge never raises on a
  novel type.
- **`wolfxl-core` dep added to the cdylib's `Cargo.toml`** (version
  `0.8`, path `crates/wolfxl-core`). First time the PyO3 surface has
  taken a direct dep on the core crate — prerequisite for the
  classifier-collapse work in the follow-up PR.
- **`Sheet::from_rows` promoted to `pub`** in `wolfxl-core`. The CSV
  backend already used it crate-internally; making it public lets the
  bridge feed externally-sourced Python lists through
  `infer_sheet_schema` / `classify_sheet` without round-tripping
  through a file.

#### Notes

- **Purely additive surface.** This PR does not replace the duplicate
  per-cell classification calls that already live inside
  `calamine_styled_backend.rs` — that wiring is the follow-up PR
  (sprint-2 task #22b). All 617 existing pytest cases still pass
  unchanged; the bridge is extra surface, not a rewrite.
- **Single source of truth for future consumers.** Python callers that
  want a classifier answer can now go through the bridge and get
  byte-identical results to `wolfxl <subcommand> --format json`. The
  cross-surface parity test lands with task #22b.

## wolfxl-core 0.8.0 / wolfxl-cli 0.8.0 (2026-04-20)

### Added

- **Multi-format `Workbook::open`**: `.xls`, `.xlsb`, `.ods`, and `.csv`
  paths now open through the same API as `.xlsx`. Dispatch lives in
  `Workbook::open`, so `wolfxl peek`, `wolfxl map`, `wolfxl agent`, and
  `wolfxl schema` all gain the new format coverage for free. This
  closes the breadth regression relative to `xleak` (the pre-2.0
  predecessor), which handled the same four formats.
- **CSV backend** (`wolfxl_core::csv_reader`, crate-private): reads a
  CSV into a single synthetic `Sheet` named after the filename stem.
  RFC-4180-ish parser handles quoted fields with embedded commas,
  doubled quotes (`""` → `"`), and `\r\n` / `\n` line endings; ragged
  rows are padded to the max column width so downstream
  `dimensions()` / `headers()` consumers see a rectangular shape.
  Cells land as `CellValue::String` — per invariant B4, schema
  inference is the single source of truth for per-column types.
- **Schema inference parses numeric-looking strings**: a CSV column of
  `"100","200",...` now classifies as `Int` instead of `String`.
  `CellValue::String` cells that parse cleanly as `i64` / `f64` are
  counted as the parsed type in `TypeCounts::observe`; strings with
  currency / thousand-separator / percent markers stay as `Other` and
  classify as `String` (the number-format string, when present, still
  drives the separate `format_category`).
- **`SourceFormat` enum** and `Workbook::format()` accessor expose
  which backend the dispatch routed to — `Xlsx`, `Xls`, `Xlsb`, `Ods`,
  or `Csv` — for callers that need to condition on it.
- **CLI multi-format smoke tests**: `tests/cli.rs` now drives `peek`
  against `.csv`, `.xls`, and `.ods` fixtures and drives `schema`
  against `.csv`, asserting the CSV's numeric columns classify as
  `int`. No goldens locked for non-xlsx renders since calamine's
  xls/xlsb/ods readers return empty styles (R1 risk from the sprint plan)
  and the boxed renderer's column widths can drift without the
  styled fast path.
- **Expanded CLI confidence matrix**: committed a tiny `.xlsb` fixture
  sourced from calamine's MIT-licensed test corpus and now smoke-tests
  `peek`, `map`, `schema`, and `agent` across `.csv`, `.xls`, `.xlsb`,
  and `.ods`. CLI help and README text now describe the broad input
  surface instead of implying `.xlsx` only.
- **Number-format-aware CLI previews**: added a formatted workbook
  fixture and assertions that human-facing `peek` box/text/CSV renders
  currency symbols and percentage formats while JSON preserves raw
  machine values. `agent` keeps compact raw numerics intentionally to
  protect token budgets.

### Changed

- **`Workbook::styles()` errors for non-xlsx formats** with a clear
  "`WorkbookStyles` only supports xlsx" message. xls/xlsb/ods carry no
  style information in calamine's public API, and CSV has no
  concept of styles. Callers that want styled rendering should
  branch on `Workbook::format()` before reaching for `styles()`.
- **`WorkbookMap` on CSV** reports a single sheet entry classified
  via the same heuristics as any other sheet; `named_ranges()`
  returns an empty slice on CSV (no workbook-level metadata
  exists).
- **`wolfxl-cli` depends on `wolfxl-core 0.8`** (was 0.7). CLI
  version bumps to 0.8.0 alongside core — shipping the two in
  lockstep keeps the version math honest for users installing via
  `cargo install wolfxl-cli`.

### Notes

- **xls / xlsb / ods are value-only today.** calamine-styles leaves
  `worksheet_style()` empty for these formats, so
  `Cell::number_format` is always `None`; schema inference still
  classifies numeric columns correctly because it reads values, not
  styles. This is the documented R1 mitigation from the sprint-2 plan.
- **CSV parsing intentionally minimal.** UTF-8 only, no custom
  delimiters, no BOM detection. If users hit workbooks that need more,
  the backend can swap to the `csv` crate later — the public API
  (`Workbook::open`, single synthetic sheet, string-valued cells)
  stays the same.

## wolfxl-core 0.7.0 (superseded by 0.8.0 above)

### Added

- **`xl/styles.xml` cellXfs walker** in `wolfxl-core`: new `ooxml`,
  `styles`, and `worksheet_xml` modules plus a `WorkbookStyles` bundle
  that parses cellXfs + numFmts and per-sheet `(row, col) → styleId`
  maps on demand. `Sheet::load` now resolves `number_format` via a
  two-step chain — calamine-styles' fast path first, then the walker
  fallback — so workbook shapes that leave `Style::get_number_format()`
  returning `None` (openpyxl-emitted styles with unpaired cellStyleXfs,
  and similar edge cases) still surface the author-intended currency /
  percentage / date codes. Public re-exports: `WorkbookStyles`,
  `XfEntry`, `BUILTIN_NUM_FMTS`, `builtin_num_fmt`, `resolve_num_fmt`.
- **Integration test**: `tests/styles_walker.rs` covers the combined
  fast-path + fallback end-to-end on a styled fixture, plus a direct
  `parse_cellxfs` + `parse_num_fmts` + `resolve_num_fmt` drive-through
  on synthetic OOXML.

### Notes

- The scope-docs "Not yet" bullet on the styles walker is now
  resolved; the `schema` format-detection note about openpyxl
  workbooks falling back to `general` no longer applies when the
  workbook actually carries `cellXfs` + `numFmts` (even if calamine
  can't see them). Workbooks that emit no styled cells at all still
  fall back to general because there is nothing to resolve.

## wolfxl-cli 0.7.0 / wolfxl-core 0.6.0 (2026-04-19)

### Added

- **`wolfxl schema <file>` subcommand**: per-column type, cardinality,
  null count, format category, and up to three sample values. Defaults
  to JSON for agent consumption; `--format text` produces a tabular
  terminal view. Pass `--sheet NAME` to scope to one sheet, omit to
  schema every sheet in the workbook.
- **`wolfxl-core::schema` module**: `InferredType`, `Cardinality`,
  `ColumnSchema`, `SheetSchema`, and the `infer_sheet_schema` entry
  point — callable from third-party Rust consumers, identical answers
  to the CLI.

### Notes

- **Cardinality buckets** are: `unique` (every non-null cell distinct),
  `categorical` (≤20 distinct AND distinct × 2 ≤ non-null — the
  "lookup-friendly dimension" bucket an agent needs to plan a `WHERE`
  clause), `high-cardinality` (everything else above the cap or with
  many distincts), and `empty`.
- **Type inference** collapses `Int + Float` in the same column to
  `Float` (numeric supertype). Any other multi-type column resolves to
  `Mixed` so an agent doesn't pick a dominant type from a noisy mix.
- **Unique-count tracking is capped at 10 000** distinct rendered
  values per column; columns past the cap report
  `unique_capped: true` and class as `high-cardinality` (the safer
  bucket — caller won't wrongly treat an unverified column as a
  categorical lookup). Picked so a million-row sheet doesn't blow
  memory on the per-column HashSet.
- **Format category is locked from the first non-empty cell** of each
  column. Mixed-format columns are rare in practice; if a user wanted
  per-cell formatting they would be looking at a CSV. Note: openpyxl-
  generated fixtures often emit no `cellXfs` styles, so format
  detection on those workbooks falls back to `general`. Real Excel-
  authored workbooks carry the styles correctly. The full styles.xml
  walker that lifts this limitation is tracked separately.

## wolfxl-cli 0.6.0 (2026-04-19)

### Added

- **`wolfxl agent <file> --max-tokens N` subcommand**: composes a
  token-budgeted workbook briefing for an LLM context window. Emits a
  workbook overview (every sheet with dims/class/first-column header),
  picks the largest `data`-class sheet (or `--sheet` override), then
  greedily fills the remaining budget with header row, head 3 rows, tail
  2 rows, and up to 8 stratified middle samples. Token counts use
  `tiktoken-rs::cl100k_base` to match the GPT-4 family tokenizer (and
  `spreadsheet-peek/benchmarks/measure_tokens.py`); verified at 0-token
  drift against Python `tiktoken`. Falls back to orientation-only output
  if the budget is too tight (and reports the overage in the footer
  rather than silently truncating).
- **Stratified row sampling**: head + tail + uniform-stride middle
  samples instead of head-only. An LLM seeing rows 1, 2, 3 of a 50-row
  P&L can't tell totals from line items; rows 1-3, 25-26, 49-50 plus
  middle samples surface the shape of the data.
- **Token budget tracker**: `Budget::used_with(buf, section)` re-encodes
  the full concatenation rather than summing per-section counts, because
  cl100k_base BPE merges across boundaries (additive checks would
  over-count and reject sections that actually fit). After PR review,
  the budget reserves a worst-case footer cost up-front so the printed
  `--max-tokens N` is honored end-to-end (body + footer).
- Best-effort `NAMED_RANGES` block (capped at 8 entries with overflow
  marker) gated through `try_append`, so a workbook with hundreds of
  named ranges cannot single-handedly drain the agent's budget.

### Notes

- `--agent` deliberately does NOT thousand-group integers (`1234567`,
  not `1,234,567`). Every comma is a token boundary in cl100k_base, so
  ungrouped costs ~2 tokens vs grouped ~5 for the same number. Pretty
  output costs the agent context.
- The orientation core (workbook overview + sheet header + columns) is
  emitted even when it overflows the budget. We'd rather report the
  overage in the footer than hide workbook structure from the agent.

## wolfxl-core 0.5.0 / wolfxl-cli 0.5.0 (2026-04-19)

### Added

- **`wolfxl map <file>` subcommand**: one-page workbook overview for agents
  that need to orient before fetching cell ranges. Emits per-sheet name,
  dimensions, headers, anchored tables, and a coarse classification
  (`empty` / `readme` / `summary` / `data`) plus workbook-level defined
  names. Two output formats: `--format json` (default, machine-parseable)
  and `--format text` (terminal-friendly, sectioned per sheet, header
  preview capped at 8 columns with overflow count).
- **`wolfxl-core::map` module**: `WorkbookMap`, `SheetMap`, `SheetClass`,
  and the `classify_sheet` heuristic, callable from third-party Rust
  consumers via the new `Workbook::map()` method.
- `Workbook::named_ranges()` and `Workbook::table_names_in_sheet(name)`
  pass-throughs to calamine's metadata accessors. Tables are now eagerly
  loaded at `Workbook::open` so these accessors stay infallible (calamine
  panics on `table_names*` without a prior `load_tables`).
- Test-only `Sheet::from_rows_for_test` constructor (gated behind
  `#[cfg(test)]`) so the classifier can exercise `Empty` / `Readme` /
  sparse `Summary` branches that the committed xlsx fixtures don't hit.

### Notes

- The classifier intentionally does not look at merged cells (the upstream
  PyO3 layer does, but `wolfxl-core` doesn't expose merge metadata yet).
  A merged-title-row sheet today classifies as `summary` via the size +
  density rule, which is the right answer for a typical dashboard.
- Pivot detection is still out of scope — calamine doesn't surface pivot
  parts directly, and the agent value of "this sheet is a pivot" is
  marginal next to dimensions + headers.

## 0.4.0 (2026-04-19)

### Added

- **`wolfxl-core` crate** (crates.io): pure-Rust xlsx reader with Excel
  number-format-aware cell rendering. Exposes `Workbook`, `Sheet`, `Cell`,
  `CellValue`, `FormatCategory`, and `format_cell` for third-party Rust
  consumers. No PyO3 coupling.
- **`wolfxl-cli` crate** (crates.io): installs the `wolfxl` binary with a
  `peek` subcommand. `wolfxl peek <file> [-n N] [-s SHEET] [-w WIDTH]
  [-e {box,text,csv,json}]` produces a styled box preview by default and
  text/csv/json exports tuned for piping into agent or shell pipelines.
  Install via `cargo install wolfxl-cli`.

### Changed

- **PyO3 0.24 → 0.28**: required for Python 3.14 support. No public Python
  API changes; all 611 pytest tests pass on 3.12 and 3.14.
- Repository converted to a Cargo workspace with the existing PyO3 cdylib
  at the root and the new `crates/wolfxl-core` + `crates/wolfxl-cli`
  members.

### Fixed

- `wolfxl-core` currency rendering: `format_currency(1.995, 2)` now returns
  `"$2.00"` (was `"$1.100"` due to splitting `trunc()`/`fract()` separately
  before rounding).

## 0.3.2 (2026-04-16)

### Added

- **Bulk styled cell records**: `Worksheet.iter_cell_records()` and `Worksheet.cell_records()` return populated cells as dictionaries with values, formulas, coordinates, and compact formatting metadata.
- **Record-shape controls**: `include_empty`, `include_format`, `include_formula_blanks`, `include_coordinate`, and per-call `data_only` options support ingestion, dataframe, and sparse-workbook workloads.
- **Robust dimensions**: `Worksheet.calculate_dimension()` now merges stale worksheet dimension tags with parsed value/formula storage and preserves offset used ranges such as `C4:C4`.

### Changed

- `max_row` / `max_column` now benefit from the same stale-dimension hardening while preserving their openpyxl-style bottom/right edge semantics.
- `calculate_dimension()` includes buffered `append()` / `write_rows()` data before save, making write-mode dimension reporting more useful for standalone callers.

## 0.3.1 (2026-02-20)

### Added

- **TIME functions**: `NOW()`, `HOUR()`, `MINUTE()`, `SECOND()` with `_serial_to_time` helper for fractional day extraction
- **OFFSET promoted to builtins**: OFFSET now registered in `_BUILTINS` via `_raw_args` protocol, making it visible in `supported_functions` (was previously a hidden evaluator special case)
- **Print area roundtrip**: `ws.print_area = "A1:D10"` now writes through to the xlsx file via the Rust backend (previously stored in Python but never flushed to the writer)

### Changed

- Builtins: 62 -> 67 (OFFSET + NOW + HOUR + MINUTE + SECOND)
- Whitelist: 63 -> 67 (now fully synced with builtins)
- Evaluator function dispatch refactored to use `_raw_args` attribute protocol instead of string-equality special case

## 0.3.0 (2026-02-19)

### Added

- **Formula engine self-sufficiency**: 62 builtin functions covering math, logic, text, lookup, date, financial, and conditional aggregation
- **openpyxl compat expansion**: freeze/split panes, unmerge_cells, print_area property, conditional formatting, data validation, named ranges, tables
- **VLOOKUP/HLOOKUP builtins**: native lookup functions without `formulas` library dependency
- **Conditional aggregation**: AVERAGEIF, AVERAGEIFS, MINIFS, MAXIFS
- **Text functions**: UPPER, LOWER, TRIM, SUBSTITUTE, TEXT, REPT, EXACT, FIND

## 0.1.1 (2026-02-16)

### Fixed

- Build full wheel matrix for macOS and Windows (Python 3.9-3.13)
- Use macos-14 (Apple Silicon) with cross-compilation for x86_64 macOS wheels (macos-13 Intel runners unavailable)
- Fix Windows build failure caused by PyO3 discovering Python 3.14 pre-release

## 0.1.0 (2026-02-15)

Initial release. Extracted from [ExcelBench](https://github.com/SynthGL/ExcelBench).

### Features

- **Read mode**: Full-fidelity xlsx reading via calamine-styles (Font, Fill, Border, Alignment, NumberFormat)
- **Write mode**: Full-fidelity xlsx writing via rust_xlsxwriter
- **Modify mode**: Surgical ZIP patching for fast read-modify-write workflows (10-14x vs openpyxl)
- **openpyxl-compatible API**: `load_workbook()`, `Workbook()`, Cell/Worksheet/Font/PatternFill/Border
- **Bulk operations**: `read_sheet_values()` / `write_sheet_values()` for batch cell I/O
- **Performance**: 3-5x faster than openpyxl for per-cell operations, up to 5x for bulk writes
