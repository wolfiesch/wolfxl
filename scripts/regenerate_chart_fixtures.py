#!/usr/bin/env python3
"""Regenerate ``tests/fixtures/charts/*.xlsx`` from openpyxl.

Run from repo root::

    python scripts/regenerate_chart_fixtures.py

Generates one tiny xlsx per chart type — kept ≤ ~5 KB so the repo
doesn't bloat. Fixtures are reference golden files used by the parity
suite to detect regressions in wolfxl's chart emit (Sprint Μ Pod-δ,
RFC-046 §8).

The 8 chart types map 1:1 to the wolfxl.chart classes shipped in 1.6:
Bar, Line, Pie, Doughnut, Area, Scatter, Bubble, Radar.

Each file carries a 6-row × 3-column data block (header + 5 rows × 2
numeric series) so the parity tests have a stable cell layout to
compare against.
"""

from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
from openpyxl import chart as oxc

REPO_ROOT = Path(__file__).resolve().parent.parent
FIXTURES_DIR = REPO_ROOT / "tests" / "fixtures" / "charts"


def _seed(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    ws.append(["", "Series A", "Series B"])
    for i in range(1, 6):
        ws.append([f"row{i}", i * 10, i * 5])


def _build(kind: str, title: str) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    _seed(ws)

    cls_map = {
        "bar": oxc.BarChart,
        "line": oxc.LineChart,
        "pie": oxc.PieChart,
        "doughnut": oxc.DoughnutChart,
        "area": oxc.AreaChart,
        "scatter": oxc.ScatterChart,
        "bubble": oxc.BubbleChart,
        "radar": oxc.RadarChart,
    }
    chart = cls_map[kind]()
    chart.title = title

    data = oxc.Reference(ws, min_col=2, min_row=1, max_col=3, max_row=6)
    chart.add_data(data, titles_from_data=True)
    cats = oxc.Reference(ws, min_col=1, min_row=2, max_row=6)
    try:
        chart.set_categories(cats)
    except Exception:
        # Pie/Doughnut sometimes complain after add_data — non-fatal.
        pass
    ws.add_chart(chart, "E2")
    return wb


# (filename, kind, title)
SPECS: tuple[tuple[str, str, str], ...] = (
    ("bar_chart_simple.xlsx", "bar", "Sales"),
    ("line_chart_with_trendline.xlsx", "line", "Trend"),
    ("pie_chart_with_legend.xlsx", "pie", "Share"),
    ("doughnut_chart_simple.xlsx", "doughnut", "Mix"),
    ("area_chart_simple.xlsx", "area", "Volume"),
    ("scatter_chart_simple.xlsx", "scatter", "Correlation"),
    ("bubble_chart_simple.xlsx", "bubble", "Bubbles"),
    ("radar_chart_simple.xlsx", "radar", "Radar"),
)


def main() -> int:
    FIXTURES_DIR.mkdir(parents=True, exist_ok=True)
    for fname, kind, title in SPECS:
        wb = _build(kind, title)
        out = FIXTURES_DIR / fname
        wb.save(out)
        size = out.stat().st_size
        print(f"  wrote {out.relative_to(REPO_ROOT)} ({size} bytes)")
    print(f"OK — {len(SPECS)} fixtures regenerated under {FIXTURES_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
