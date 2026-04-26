<p align="center">
  <h1 align="center">WolfXL</h1>
  <p align="center">
    <strong>The fastest openpyxl-compatible Excel library for Python.</strong><br>
    Drop-in replacement backed by Rust — up to 5x faster with zero code changes.
  </p>
</p>

<p align="center">
  <a href="https://pypi.org/project/wolfxl/"><img src="https://img.shields.io/pypi/v/wolfxl?color=blue&label=PyPI" alt="PyPI"></a>
  <a href="https://pypi.org/project/wolfxl/"><img src="https://img.shields.io/pypi/pyversions/wolfxl?color=blue" alt="Python"></a>
  <a href="https://github.com/SynthGL/wolfxl/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-MIT-green" alt="License"></a>
  <a href="https://excelbench.vercel.app"><img src="https://img.shields.io/badge/benchmarks-ExcelBench-orange" alt="ExcelBench"></a>
</p>

---

## Replaces openpyxl. One import change.

```diff
- from openpyxl import load_workbook, Workbook
- from openpyxl.styles import Font, PatternFill, Alignment, Border
+ from wolfxl import load_workbook, Workbook, Font, PatternFill, Alignment, Border
```

Your existing code works as-is. Same `ws["A1"].value`, same `Font(bold=True)`, same `wb.save()`.

---

<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="assets/benchmark-dark.svg">
    <source media="(prefers-color-scheme: light)" srcset="assets/benchmark-light.svg">
    <img alt="WolfXL vs openpyxl benchmark chart" src="assets/benchmark-dark.svg" width="700">
  </picture>
</p>

<p align="center">
  <sub>Measured with <a href="https://excelbench.vercel.app">ExcelBench</a> on Apple M4 Pro, Python 3.12, median of 3 runs.</sub>
</p>

## Install

```bash
pip install wolfxl
```

## Quick Start

```python
from wolfxl import load_workbook, Workbook, Font, PatternFill

# Write a styled spreadsheet
wb = Workbook()
ws = wb.active
ws["A1"].value = "Product"
ws["A1"].font = Font(bold=True, color="FFFFFF")
ws["A1"].fill = PatternFill(fill_type="solid", fgColor="336699")
ws["A2"].value = "Widget"
ws["B2"].value = 9.99
wb.save("report.xlsx")

# Read it back — styles included
wb = load_workbook("report.xlsx")
ws = wb[wb.sheetnames[0]]
for row in ws.iter_rows(values_only=False):
    for cell in row:
        print(cell.coordinate, cell.value, cell.font.bold)
wb.close()
```

## Three Modes

<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="assets/architecture-dark.svg">
    <source media="(prefers-color-scheme: light)" srcset="assets/architecture-light.svg">
    <img alt="WolfXL architecture" src="assets/architecture-dark.svg" width="680">
  </picture>
</p>

| Mode | Usage | Engine | What it does |
|------|-------|--------|--------------|
| **Read** | `load_workbook(path)` | [calamine-styles](https://crates.io/crates/calamine-styles) | Parse XLSX with full style extraction |
| **Write** | `Workbook()` | [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) | Create new XLSX files from scratch |
| **Modify** | `load_workbook(path, modify=True)` | XlsxPatcher | Surgical ZIP patch — only changed cells are rewritten |

Modify mode preserves everything it doesn't touch: charts, macros, images, pivot tables, VBA.

## Supported Features

Features marked **Preserved** are kept verbatim on modify-mode round-trip (open, edit other cells, save). Construction from Python code for those features is tracked below as roadmap items.

| Category | Features |
|----------|----------|
| **Data** | Cell values (string, number, date, bool), formulas, comments, hyperlinks |
| **Styling** | Font (bold, italic, underline, color, size), fills, borders, number formats, alignment; `Color(theme=...)` and `Color(indexed=...)` accepted |
| **Structure** | Multiple sheets, merged cells, defined names (read + write), freeze panes, row heights, column widths, document properties |
| **Tables / Validation / CF** | `ws.tables`, `ws.add_table`, `ws.data_validations`, `ws.conditional_formatting` (read + write in `Workbook()` mode) |
| **Iteration** | `iter_rows`, `iter_cols`, `rows`, `columns`, `values`, range slicing (`ws["A1:B2"]`, `ws["A:B"]`, `ws[1:3]`) |
| **Utils** | `get_column_letter`, `column_index_from_string`, `coordinate_to_tuple`, `range_boundaries`, `absolute_coordinate`, `quote_sheetname`, `range_to_tuple`, `rows_from_range`, `cols_from_range`, `get_column_interval`, `dataframe_to_rows`, `is_date_format` |
| **Preserved (read-only)** | Charts, images, pivot tables, macros (VBA) — round-trip cleanly through modify mode |

### openpyxl compatibility status

Modules that import from openpyxl generally work against wolfxl. Unsupported classes raise `NotImplementedError` with a clear hint at the construction site - no silent no-ops.

| Class / API | Status |
|-------------|--------|
| `Font`, `PatternFill`, `Border`, `Side`, `Alignment`, `Color` | Full support |
| `Comment`, `Hyperlink` | Read + write (write mode); modify-mode setters T1.5 |
| `DataValidation`, `Table`, `TableStyleInfo`, `TableColumn` | Read + write (write mode); modify-mode setters T1.5 |
| `CellIsRule`, `FormulaRule`, `ColorScaleRule`, `DataBarRule`, `IconSetRule` | Read + write (write mode); modify-mode setters T1.5 |
| `DefinedName`, `DocumentProperties` | Read + write (write mode); modify-mode setters T1.5 |
| `NamedStyle`, `Protection`, `GradientFill`, `DifferentialStyle` | Stub (raises `NotImplementedError`) |
| `BarChart`, `LineChart`, `PieChart`, `Reference`, `Series` (from `wolfxl.chart`) | Stub - use modify mode to preserve existing charts |
| `Image` (from `wolfxl.drawing.image`) | Stub - preserved on modify-mode round-trip |
| `AutoFilter`, `PivotTable` | Stub - preserved on modify-mode round-trip |
| `ws.insert_rows`, `ws.delete_rows` | **Full support** (modify mode, 1.1+) — RFC-030 |
| `ws.insert_cols`, `ws.delete_cols` | **Full support** (modify mode, 1.1+) — RFC-031 |
| `ws.move_range` | **Full support** (modify mode, 1.1+) — RFC-034 |
| `wb.move_sheet` | **Full support** (modify mode, 1.1+) — RFC-036 |
| `wb.copy_worksheet` | **Full support** (modify mode only, 1.1+) — RFC-035 |

## Performance at Scale

| Scale | File size | WolfXL Read | openpyxl Read | WolfXL Write | openpyxl Write |
|-------|-----------|-------------|---------------|--------------|----------------|
| 100K cells | 400 KB | **0.11s** | 0.42s | **0.06s** | 0.28s |
| 1M cells | 3 MB | **1.1s** | 4.0s | **0.9s** | 2.9s |
| 5M cells | 25 MB | **6.0s** | 20.9s | **3.2s** | 15.5s |
| 10M cells | 45 MB | **13.0s** | 47.8s | **6.7s** | 31.8s |

Throughput stays flat as files grow — no hidden O(n^2) pathology.

## How WolfXL Compares

Every Rust-backed Python Excel project picks a different slice of the problem. WolfXL is the only one that covers all three: formatting, modify mode, and openpyxl API compatibility.

| Library | Read | Write | Modify | Styling | openpyxl API |
|---------|:----:|:-----:|:------:|:-------:|:------------:|
| [fastexcel](https://github.com/ToucanToco/fastexcel) | Yes | — | — | — | — |
| [python-calamine](https://github.com/dimastbk/python-calamine) | Yes | — | — | — | — |
| [FastXLSX](https://github.com/shuangluoxss/fastxlsx) | Yes | Yes | — | — | — |
| [rustpy-xlsxwriter](https://github.com/rahmadafandi/rustpy-xlsxwriter) | — | Yes | — | Partial | — |
| **WolfXL** | **Yes** | **Yes** | **Yes** | **Yes** | **Yes** |

- **Styling** = reads and writes fonts, fills, borders, alignment, number formats
- **Modify** = open an existing file, change cells, save back — without rebuilding from scratch
- **openpyxl API** = same `load_workbook`, `Workbook`, `Cell`, `Font`, `PatternFill` objects

Upstream [calamine](https://github.com/tafia/calamine) does not parse styles. WolfXL's read engine uses [calamine-styles](https://crates.io/crates/calamine-styles), a fork that adds Font/Fill/Border/Alignment/NumberFormat extraction from OOXML.

## Batch APIs for Maximum Speed

For write-heavy workloads, use `append()` or `write_rows()` instead of cell-by-cell access. These APIs buffer rows as raw Python lists and flush them to Rust in a single call at save time, bypassing per-cell FFI overhead entirely.

```python
from wolfxl import Workbook

wb = Workbook()
ws = wb.active

# append() — fast sequential writes (3.7x faster than cell-by-cell)
ws.append(["Name", "Amount", "Date"])
for row in data:
    ws.append(row)

# write_rows() — fast writes at arbitrary positions
ws.write_rows(header_grid, start_row=1, start_col=1)
ws.write_rows(data_grid, start_row=5, start_col=1)

wb.save("output.xlsx")
```

For reads, `iter_rows(values_only=True)` uses a fast bulk path that reads all values in a single Rust call (6.7x faster than openpyxl):

```python
wb = load_workbook("data.xlsx")
ws = wb[wb.sheetnames[0]]
for row in ws.iter_rows(values_only=True):
    process(row)  # row is a tuple of plain Python values
```

For ingestion, dataframe, or review workflows that need values plus formatting
signals, use `cell_records()`. It returns compact dictionaries without creating
one Python `Cell` object per coordinate:

```python
records = ws.cell_records(
    include_format=True,
    include_formula_blanks=False,
    include_coordinate=False,
)

for record in records:
    print(record["row"], record["column"], record["value"], record.get("number_format"))
```

| API | vs openpyxl | How |
|-----|-------------|-----|
| `ws.append(row)` | **3.7x** faster write | Buffers rows, single Rust call at save |
| `ws.write_rows(grid)` | **3.7x** faster write | Same mechanism, arbitrary start position |
| `ws.iter_rows(values_only=True)` | **6.7x** faster read | Single Rust call, no Cell objects |
| `ws.cell_records()` | Fast styled sparse read | Single Rust call, values plus compact format metadata |
| `ws.cell(r, c, value=v)` | **1.6x** faster write | Per-cell FFI (compatible but slower) |

## Formula Engine

WolfXL includes a **built-in formula evaluator** with 62 functions across 7 categories. Calculate formulas without external dependencies - no need for `formulas` or `xlcalc`.

```python
from wolfxl import Workbook
from wolfxl.calc import calculate

wb = Workbook()
ws = wb.active
ws["A1"] = 100
ws["A2"] = 200
ws["A3"] = "=SUM(A1:A2)"
ws["B1"] = "=PMT(0.05/12, 360, -300000)"  # monthly mortgage payment

results = calculate(wb)
print(results["Sheet!A3"])  # 300
print(results["Sheet!B1"])  # 1610.46...

# Recalculate after changes
ws["A1"] = 500
results = calculate(wb)
print(results["Sheet!A3"])  # 700
```

| Category | Functions |
|----------|-----------|
| **Math** (10) | SUM, ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, POWER, SQRT, SIGN |
| **Logic** (5) | IF, AND, OR, NOT, IFERROR |
| **Lookup** (7) | VLOOKUP, HLOOKUP, INDEX, MATCH, OFFSET, CHOOSE, XLOOKUP |
| **Statistical** (13) | AVERAGE, AVERAGEIF, AVERAGEIFS, COUNT, COUNTA, COUNTIF, COUNTIFS, MIN, MINIFS, MAX, MAXIFS, SUMIF, SUMIFS |
| **Financial** (7) | PV, FV, PMT, NPV, IRR, SLN, DB |
| **Text** (13) | LEFT, RIGHT, MID, LEN, CONCATENATE, UPPER, LOWER, TRIM, SUBSTITUTE, TEXT, REPT, EXACT, FIND |
| **Date** (8) | TODAY, DATE, YEAR, MONTH, DAY, EDATE, EOMONTH, DAYS |

Named ranges are resolved automatically. Error values (`#N/A`, `#VALUE!`, `#DIV/0!`, `#REF!`, `#NUM!`, `#NAME?`) propagate through formula chains like real Excel. Install `pip install wolfxl[calc]` for extended formula coverage via the `formulas` library fallback.

## Case Study: SynthGL

[SynthGL](https://github.com/SynthGL) switched from openpyxl to WolfXL for their GL journal exports (14-column financial data, 1K-50K rows). Results: **4x faster writes**, **9x faster reads** at scale. 50K-row exports dropped from 7.6s to 1.3s. [Read the full case study](docs/case-study-synthgl.md).

## How It Works

WolfXL is a thin Python layer over compiled Rust engines, connected via [PyO3](https://pyo3.rs/). The Python side uses **lazy cell proxies** — opening a 10M-cell file is instant. Values and styles are fetched from Rust only when you access them. On save, dirty cells are flushed in one batch, avoiding per-cell FFI overhead.

## License

MIT
