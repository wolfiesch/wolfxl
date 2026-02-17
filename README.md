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
  <sub>Measured with <a href="https://excelbench.vercel.app">ExcelBench</a> on Apple M1 Pro, Python 3.12, median of 3 runs.</sub>
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

| Category | Features |
|----------|----------|
| **Data** | Cell values (string, number, date, bool), formulas, hyperlinks, comments |
| **Styling** | Font (bold, italic, underline, color, size), fills, borders, number formats, alignment |
| **Structure** | Multiple sheets, merged cells, named ranges, freeze panes, tables |
| **Advanced** | Data validation, conditional formatting |

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

## How It Works

WolfXL is a thin Python layer over compiled Rust engines, connected via [PyO3](https://pyo3.rs/). The Python side uses **lazy cell proxies** — opening a 10M-cell file is instant. Values and styles are fetched from Rust only when you access them. On save, dirty cells are flushed in one batch, avoiding per-cell FFI overhead.

## License

MIT
