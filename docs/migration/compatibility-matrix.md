# Compatibility Matrix

Status legend:

- `Supported` - implemented and covered by tests/fixtures
- `Partial` - implemented with caveats
- `Not Yet` - not implemented in WolfXL API surface

## WolfXL API Surface (current)

| Area | Status | Notes |
|---|---|---|
| `load_workbook(path)` | Supported | Read mode |
| `load_workbook(path, modify=True)` | Supported | Modify mode via patcher |
| `Workbook()` | Supported | Write mode |
| `wb.sheetnames` | Supported | List of sheet names |
| `wb.active` | Supported | Returns first sheet |
| `wb["Sheet"]` | Supported | Sheet by name |
| `wb.create_sheet(title)` | Supported | Write mode only |
| `wb.save(path)` | Supported | Write/modify mode |
| `ws["A1"]`, assignment | Supported | Cell access/updates |
| `ws.cell(row, column, value)` | Supported | openpyxl-like API |
| `ws.iter_rows(...)` | Supported | values and cell objects |
| `ws.merge_cells(range)` | Supported | Write mode |
| Font/fill/border/alignment styles | Supported | Via style dataclasses |
| Number format | Supported | `cell.number_format` |
| Full openpyxl API parity | Partial | Focused subset |

## Ecosystem Comparison

### Pure Python Libraries

| Capability | openpyxl | XlsxWriter | pandas (openpyxl engine) |
|---|---|---|---|
| Read `.xlsx` | Yes | No (write-only) | Yes |
| Write `.xlsx` | Yes | Yes | Yes |
| Modify existing workbook | Yes | No | No |
| Style read/write | Yes | Yes (write) | No (DataFrame coercion) |
| openpyxl-style API | Native | Different API | Different API |
| Speed (relative) | 1x baseline | ~1.5x write | ~1x (wraps openpyxl) |

### Rust-Backed Libraries

| Capability | fastexcel | python-calamine | FastXLSX | rustpy-xlsxwriter | WolfXL |
|---|---|---|---|---|---|
| Read `.xlsx` | Yes | Yes | Yes | No | **Yes** |
| Write `.xlsx` | No | No | Yes | Yes | **Yes** |
| Modify existing file | No | No | No | No | **Yes** |
| Style extraction (read) | No | No | No | N/A | **Yes** |
| Style writing | N/A | N/A | No | Partial | **Yes** |
| openpyxl-compatible API | No | No | No | No | **Yes** |

Key takeaways:

- **fastexcel** and **python-calamine** are excellent data readers — fast for ingestion into Arrow/DataFrames, but no formatting or write support.
- **FastXLSX** combines calamine + rust_xlsxwriter but skips formatting, merged cells, and modify mode.
- **rustpy-xlsxwriter** wraps rust_xlsxwriter for writes with some formatting, but no read or openpyxl API.
- **WolfXL** targets the full openpyxl workflow: read with styles, write with styles, and modify existing files.

Upstream [calamine](https://github.com/tafia/calamine) does not parse cell styles. WolfXL's read engine uses [calamine-styles](https://crates.io/crates/calamine-styles), a fork that adds Font/Fill/Border/Alignment/NumberFormat extraction from OOXML.
