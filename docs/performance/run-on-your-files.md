# Run Benchmarks on Your Files

This harness compares read / write / modify-mode and (NEW in v1.7)
chart-construction + copy-worksheet against openpyxl on your own
workbooks.

## Prepare

```bash
pip install wolfxl openpyxl
# Optional, for encrypted workbooks:
pip install wolfxl[encrypted]
```

## Single-file benchmark

```python
"""Benchmark wolfxl 1.7 vs openpyxl 3.1 on a single workbook.

Run:
    python bench.py path/to/your_workbook.xlsx
"""
from __future__ import annotations

import statistics
import sys
from pathlib import Path
from time import perf_counter

from openpyxl import Workbook as OWorkbook, load_workbook as openpyxl_load
from wolfxl import Workbook as WWorkbook, load_workbook as wolfxl_load


def timed(fn, runs: int = 5) -> float:
    """Return median wall-clock seconds across ``runs`` invocations."""
    fn()  # warm-up; discarded
    samples = []
    for _ in range(runs):
        t0 = perf_counter()
        fn()
        samples.append(perf_counter() - t0)
    return statistics.median(samples)


def report(label: str, openpyxl_t: float, wolfxl_t: float) -> None:
    speedup = openpyxl_t / wolfxl_t if wolfxl_t else float("inf")
    print(
        f"{label:<40s}  openpyxl={openpyxl_t:6.3f}s  "
        f"wolfxl={wolfxl_t:6.3f}s  speedup={speedup:5.1f}×"
    )


def main(input_file: Path) -> None:
    # ------------------------------------------------------------------
    # 1) Read
    # ------------------------------------------------------------------
    op_read = timed(lambda: openpyxl_load(input_file))
    wf_read = timed(lambda: wolfxl_load(str(input_file)))
    report("read", op_read, wf_read)

    # 2) Streaming read (wolfxl auto-engages > 50k rows; force it here)
    def wf_stream():
        wb = wolfxl_load(str(input_file), read_only=True)
        for ws in wb.worksheets:
            for _ in ws.iter_rows(values_only=True):
                pass

    def op_stream():
        wb = openpyxl_load(input_file, read_only=True)
        for ws in wb.worksheets:
            for _ in ws.iter_rows(values_only=True):
                pass

    op_s = timed(op_stream)
    wf_s = timed(wf_stream)
    report("streaming-read (iter_rows, values_only)", op_s, wf_s)

    # ------------------------------------------------------------------
    # 3) Modify+save (single-cell touch — the wolfxl differentiator)
    # ------------------------------------------------------------------
    def write_openpyxl():
        wb = openpyxl_load(input_file)
        ws = wb[wb.sheetnames[0]]
        ws["A1"] = "benchmark"
        wb.save("out_openpyxl.xlsx")

    def write_wolfxl():
        wb = wolfxl_load(str(input_file), modify=True)
        ws = wb[wb.sheetnames[0]]
        ws["A1"] = "benchmark"
        wb.save("out_wolfxl.xlsx")

    op_m = timed(write_openpyxl)
    wf_m = timed(write_wolfxl)
    report("modify+save (touch 1 cell)", op_m, wf_m)

    # ------------------------------------------------------------------
    # 4) copy_worksheet (NEW in v1.7 docs harness)
    # ------------------------------------------------------------------
    def copy_openpyxl():
        wb = openpyxl_load(input_file)
        wb.copy_worksheet(wb[wb.sheetnames[0]])
        wb.save("copy_openpyxl.xlsx")

    def copy_wolfxl():
        wb = wolfxl_load(str(input_file), modify=True)
        wb.copy_worksheet(wb[wb.sheetnames[0]])
        wb.save("copy_wolfxl.xlsx")

    op_c = timed(copy_openpyxl)
    wf_c = timed(copy_wolfxl)
    report("copy_worksheet+save", op_c, wf_c)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    main(Path(sys.argv[1]))
```

## Chart-construction microbenchmark

```python
"""Compare chart-construction throughput.

Run:
    python chart_bench.py
"""
from __future__ import annotations

import statistics
from time import perf_counter

import openpyxl
import openpyxl.chart
import wolfxl
import wolfxl.chart


def make_chart_openpyxl():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Region", "Q1", "Q2", "Q3", "Q4"])
    for i in range(1000):
        ws.append([f"row{i}", i, i + 1, i + 2, i + 3])
    chart = openpyxl.chart.BarChart()
    data = openpyxl.chart.Reference(ws, min_col=2, min_row=1, max_col=5, max_row=1001)
    chart.add_data(data, titles_from_data=True)
    ws.add_chart(chart, "G2")
    wb.save("op_chart.xlsx")


def make_chart_wolfxl():
    wb = wolfxl.Workbook()
    ws = wb.active
    ws.append(["Region", "Q1", "Q2", "Q3", "Q4"])
    for i in range(1000):
        ws.append([f"row{i}", i, i + 1, i + 2, i + 3])
    chart = wolfxl.chart.BarChart()
    data = wolfxl.chart.Reference(ws, min_col=2, min_row=1, max_col=5, max_row=1001)
    chart.add_data(data, titles_from_data=True)
    ws.add_chart(chart, "G2")
    wb.save("wf_chart.xlsx")


def timed(fn, runs: int = 5) -> float:
    fn()
    samples = []
    for _ in range(runs):
        t0 = perf_counter()
        fn()
        samples.append(perf_counter() - t0)
    return statistics.median(samples)


if __name__ == "__main__":
    op = timed(make_chart_openpyxl)
    wf = timed(make_chart_wolfxl)
    print(f"openpyxl chart construction: {op:.3f}s")
    print(f"wolfxl   chart construction: {wf:.3f}s   ({op/wf:.1f}×)")
```

## Tips

- Run at least 3 (preferably 5) iterations and compare medians.
  Single-shot runs are dominated by JIT warmup and disk cache.
- Keep the same input file across all comparisons.
- Validate correctness after each run — use openpyxl to read the
  wolfxl-produced workbook, or vice versa, and assert cell-value
  equality on a sample.
- For modify-mode benchmarks, the speedup is most pronounced on
  *large* workbooks with *small* edits — that's where wolfxl's
  surgical patcher dominates openpyxl's full-DOM rewrite.
- For chart benchmarks, the rendering of `<c:numRef>` /
  `<c:strRef>` cached values is what dominates the openpyxl
  baseline; wolfxl skips the cache write and lets Excel rebuild
  on open.

## What to send when reporting results

When sharing benchmark numbers (e.g. as a wolfxl issue or PR), include:

1. The fixture description (1k rows × 10 cols × 1 chart, etc.).
2. Hardware: CPU model + RAM.
3. OS + Python version.
4. WolfXL version (`python -c "import wolfxl; print(wolfxl.__version__)"`)
   and openpyxl version.
5. Number of runs and aggregation method (median is recommended).
6. Raw output from the harness above.
