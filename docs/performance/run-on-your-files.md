# Run Benchmarks on Your Files

This quick harness compares read and write paths for your workload.

## Prepare

```bash
pip install wolfxl openpyxl
```

## Script

```python
from __future__ import annotations

from pathlib import Path
from time import perf_counter

from openpyxl import load_workbook as openpyxl_load
from wolfxl import Workbook, load_workbook as wolfxl_load


def timed(label: str, fn):
    t0 = perf_counter()
    fn()
    dt = perf_counter() - t0
    print(f"{label}: {dt:.4f}s")


input_file = Path("your_workbook.xlsx")

timed("openpyxl read", lambda: openpyxl_load(input_file))
timed("wolfxl read", lambda: wolfxl_load(str(input_file)))


def write_openpyxl() -> None:
    wb = openpyxl_load(input_file)
    ws = wb[wb.sheetnames[0]]
    ws["A1"] = "benchmark"
    wb.save("out_openpyxl.xlsx")


def write_wolfxl() -> None:
    wb = wolfxl_load(str(input_file), modify=True)
    ws = wb[wb.sheetnames[0]]
    ws["A1"] = "benchmark"
    wb.save("out_wolfxl.xlsx")


timed("openpyxl modify+save", write_openpyxl)
timed("wolfxl modify+save", write_wolfxl)
```

## Tips

- Run at least 3 times and compare medians.
- Keep the same input file for fair comparisons.
- Validate correctness after each run.
