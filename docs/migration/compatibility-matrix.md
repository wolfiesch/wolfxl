# Compatibility Matrix

> **Reference**: WolfXL **v1.7.0** (Sprint Ξ).
> **Status as of**: 2026-04-27.

Status legend:

- **Supported** — implemented and covered by tests/fixtures.
- **Partial** — implemented with documented caveats.
- **Not Yet** — not implemented; tracked via RFC.
- **Out of scope** — explicitly out of roadmap.

## Construction-side parity (the v1.7 headline)

The following constructors all work at the same call site you'd use
in openpyxl 3.1.x:

### Workbook + Worksheet

| openpyxl path | WolfXL status | Notes |
|---|---|---|
| `Workbook()` | Supported | Write mode |
| `Workbook(write_only=True)` | Supported | Same as `Workbook()` — wolfxl always writes deterministically |
| `load_workbook(path)` | Supported | Read mode |
| `load_workbook(path, data_only=True)` | Supported | Cached values returned |
| `load_workbook(path, modify=True)` | Supported | Modify mode (surgical patcher; faster than DOM rewrite) |
| `load_workbook(path, read_only=True)` | Supported | Streaming reads (auto-engages > 50k rows) |
| `load_workbook(path, password=...)` | Supported | OOXML decryption (needs `wolfxl[encrypted]`) |
| `load_workbook(path, rich_text=True)` | Supported | `Cell.value` returns `CellRichText` |
| `wb.save(path)` | Supported | |
| `wb.save(path, password=...)` | Supported | Agile (AES-256) encryption |
| `wb["Sheet"]`, `wb.active`, `wb.sheetnames` | Supported | |
| `wb.create_sheet(title)` | Supported | |
| `wb.copy_worksheet(ws)` | Supported | RFC-035; diverges from openpyxl in 5 documented ways (always more preservation) |
| `wb.move_sheet(name, offset)` | Supported | RFC-036 |
| `wb.remove(ws)` | Supported | |

### Cell + style API

| openpyxl path | WolfXL status | Notes |
|---|---|---|
| `ws["A1"].value`, `ws.cell(row, col)` | Supported | |
| `cell.coordinate`, `cell.row`, `cell.column` | Supported | |
| `cell.font / fill / border / alignment / number_format` | Supported | |
| `cell.protection` | Supported | |
| `cell.comment = Comment(...)` | Supported | RFC-023 |
| `cell.hyperlink = Hyperlink(...)` | Supported | RFC-022 |
| `cell.rich_text` (read) | Supported | RFC-040 |
| `cell.value = CellRichText(...)` (write) | Supported | RFC-040 |
| `cell.data_type` | Supported | |
| `cell.is_date` | Supported | |

### Charts (v1.6 + v1.6.1 — RFC-046)

| openpyxl class | WolfXL status |
|---|---|
| `BarChart`, `LineChart`, `PieChart`, `DoughnutChart` | Supported |
| `AreaChart`, `ScatterChart`, `BubbleChart`, `RadarChart` | Supported |
| `BarChart3D`, `LineChart3D`, `PieChart3D` (alias `Pie3D`), `AreaChart3D` | Supported |
| `SurfaceChart`, `SurfaceChart3D` | Supported |
| `StockChart` (Open-High-Low-Close) | Supported |
| `ProjectedPieChart` (pie-of-pie / bar-of-pie) | Supported |
| `Reference`, `Series` | Supported |
| Per-series `marker`, `smooth`, `data_labels`, `trendline`, `error_bars` | Supported |
| `vary_colors`, `gap_width`, `overlap`, `hole_size`, `bubble_scale`, `radar_style`, `scatter_style` | Supported |
| `chart.title` accepts `str` / `Title` / `RichText` | Supported (v1.7) |
| `Worksheet.add_chart(chart, anchor)` (write + modify mode) | Supported |
| `Worksheet.remove_chart(chart)` (NEW v1.7) | Supported |
| `Worksheet.replace_chart(old, new)` (NEW v1.7) | Supported |
| Combination charts (e.g. bar + line on shared axes) | Not Yet — post-v1.7 |
| Pivot-chart linkage (`<c:pivotSource>`) | Not Yet — Sprint Ν / v2.0 |

### Images (v1.5 — RFC-045)

| openpyxl path | WolfXL status |
|---|---|
| `from openpyxl.drawing.image import Image` | Supported (`from wolfxl.drawing.image import Image`) |
| `Image("logo.png")` (PNG / JPEG / GIF / BMP) | Supported |
| One-cell anchor (`ws.add_image(img, "B5")`) | Supported |
| Two-cell anchor (`TwoCellAnchor`) | Supported |
| Absolute anchor (`AbsoluteAnchor`) | Supported |
| Modify-mode `add_image` | Supported |

### Worksheet structural ops (v1.1 — RFC-030/031/034/035/036)

| openpyxl path | WolfXL status |
|---|---|
| `ws.insert_rows(idx, amount)` | Supported |
| `ws.delete_rows(idx, amount)` | Supported |
| `ws.insert_cols(idx, amount)` | Supported |
| `ws.delete_cols(idx, amount)` | Supported |
| `ws.move_range(range, rows, cols)` | Supported |
| `wb.copy_worksheet(ws)` | Supported (with documented divergences) |
| `wb.move_sheet(name, offset)` | Supported |

### Modify-mode mutations (v1.0 / v1.1 — T1.5)

| openpyxl idiom | WolfXL status |
|---|---|
| `wb.properties.title = ...` (and other DocumentProperties) | Supported (RFC-020) |
| `wb.defined_names[name] = DefinedName(...)` | Supported (RFC-021) |
| `cell.comment = Comment(...)` | Supported (RFC-023) |
| `cell.hyperlink = Hyperlink(...)` | Supported (RFC-022) |
| `ws.add_table(Table(...))` | Supported (RFC-024) |
| `ws.data_validations.append(...)` | Supported (RFC-025) |
| `ws.conditional_formatting.add(...)` | Supported (RFC-026) |
| Cell value / font / fill / border / alignment / number-format mutation | Supported |

### Read-side parity (v1.3 / v1.4 — RFC-040 .. 043)

| openpyxl behaviour | WolfXL status |
|---|---|
| `.xlsx` reads | Supported |
| `.xlsb` reads | Supported (calamine; styles are `NotImplementedError` — read-only of values) |
| `.xls` reads (BIFF8) | Supported (calamine; same caveats as `.xlsb`) |
| `.ods` reads | Out of scope |
| `read_only=True` streaming | Supported (`tests/parity/test_streaming_parity.py`) |
| `password=` decryption | Supported (`pip install wolfxl[encrypted]`) |
| Rich-text reads | Supported (`Cell.rich_text` always; `Cell.value` opt-in via `rich_text=True`) |
| Cached formula results (`data_only=True`) | Supported |

### Utils

All seven openpyxl utility symbols ship under `wolfxl.utils.`:

| openpyxl path | WolfXL path |
|---|---|
| `openpyxl.utils.get_column_letter` | `wolfxl.utils.cell.get_column_letter` |
| `openpyxl.utils.column_index_from_string` | `wolfxl.utils.cell.column_index_from_string` |
| `openpyxl.utils.range_boundaries` | `wolfxl.utils.cell.range_boundaries` |
| `openpyxl.utils.coordinate_to_tuple` | `wolfxl.utils.cell.coordinate_to_tuple` |
| `openpyxl.utils.is_date_format` | `wolfxl.utils.numbers.is_date_format` |
| `openpyxl.utils.datetime.from_excel` | `wolfxl.utils.datetime.from_excel` |
| `openpyxl.utils.datetime.CALENDAR_WINDOWS_1900` | `wolfxl.utils.datetime.CALENDAR_WINDOWS_1900` |

Bound checks (`get_column_letter` capped at 18278 = ZZZ) and the
1900 leap-year correction (`from_excel`) match openpyxl
verbatim.

## What's *not* in v1.7 (deferred or out of scope)

| Capability | Status | Tracked at |
|---|---|---|
| Pivot table construction | **Not Yet — Sprint Ν / v2.0** | `Plans/sprint-xi.md` |
| Pivot cache definitions | **Not Yet — Sprint Ν / v2.0** | `Plans/sprint-xi.md` |
| Pivot-chart linkage | **Not Yet — Sprint Ν / v2.0** | RFC-046 §9 |
| Combination charts (multi-plot) | Not Yet | RFC-046 v1.6.1 release notes "Out of scope" |
| `<c:displayUnits>` on value axes | Not Yet | RFC-046 §9 |
| Per-data-point overrides (`<c:dPt>`) | Not Yet | RFC-046 §9 |
| Removal of charts that survive from source workbook | Not Yet — v1.8 | RFC-050 §6 / `Worksheet.remove_chart` docstring |
| OpenDocument (`.ods`) | Out of scope | Not on roadmap |
| `.xlsb` / `.xls` writes | Out of scope | xlsx-only; transcribe via fresh `Workbook()` |
| Style accessors on `.xlsb` / `.xls` reads | Not Yet | calamine doesn't surface non-xlsx styles |

## Ecosystem comparison

### Pure Python libraries

| Capability | openpyxl | XlsxWriter | pandas (`engine="openpyxl"`) | **WolfXL 1.7** |
|---|---|---|---|---|
| Read `.xlsx` | Yes | No | Yes | **Yes** |
| Read `.xlsb` / `.xls` | No | No | Yes (`engine="calamine"`) | **Yes** |
| Write `.xlsx` | Yes | Yes | Yes (via openpyxl) | **Yes** |
| Modify existing workbook | Yes (full DOM) | No | No | **Yes (surgical)** |
| Style read / write | Yes | Yes (write only) | No (DataFrame coercion) | **Yes** |
| Streaming reads | Yes (`read_only=True`) | N/A | No | **Yes** |
| Encryption (read) | No | No | No | **Yes** (`wolfxl[encrypted]`) |
| Encryption (write) | No | No | No | **Yes** |
| Chart construction | Yes (16 types) | Yes (12 types) | No | **Yes (16 types)** |
| Image construction | Yes | Yes | No | **Yes** |
| `copy_worksheet` | Yes (drops tables / DV / CF / sheet-scoped names) | N/A | N/A | **Yes (preserves everything)** |
| Pivot table construction | Yes | No | No | **Not yet (v2.0)** |
| Same API as openpyxl | Native | Different API | Different | **Native** |
| Performance (relative) | 1× | ~1.5× write | ~1× (wraps openpyxl) | **5–20×** read+write+modify |

### Rust-backed libraries

| Capability | fastexcel | python-calamine | FastXLSX | rustpy-xlsxwriter | **WolfXL 1.7** |
|---|---|---|---|---|---|
| Read `.xlsx` | Yes | Yes | Yes | No | **Yes** |
| Write `.xlsx` | No | No | Yes | Yes | **Yes** |
| Modify existing file | No | No | No | No | **Yes** |
| Style read | No | No | No | N/A | **Yes** |
| Style write | N/A | N/A | No | Partial | **Yes** |
| openpyxl-compatible API | No | No | No | No | **Yes** |
| Charts / images / pivots / structural ops | No | No | No | Partial (charts only) | **Yes** (charts, images, structural; pivots v2.0) |

WolfXL is the only Rust-backed Python library that targets the **full
openpyxl construction surface**, not just the read or write path.

## Footnotes on the underlying readers

- WolfXL's `.xlsx` reader uses [calamine-styles](https://crates.io/crates/calamine-styles),
  a fork of [calamine](https://github.com/tafia/calamine) that adds
  Font / Fill / Border / Alignment / NumberFormat extraction.
  Upstream calamine doesn't surface styles.
- `.xlsb` and `.xls` reads route through upstream calamine directly
  (no styles).

## Tracking

The exhaustive list of every openpyxl symbol WolfXL handles, with
`wolfxl_supported=True/False` per call site, lives in
`tests/parity/openpyxl_surface.py`. The ratchet test
`test_known_gap_still_gaps` fails red the moment a previously-deferred
gap closes — that's the signal to flip the entry to `True` and remove
the corresponding row from `tests/parity/KNOWN_GAPS.md`.
