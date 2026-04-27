"""``wolfxl.chart`` — openpyxl-shaped chart construction API.

Sprint Μ Pod-β (RFC-046) replaces the previous ``_make_stub`` placeholders
with real chart classes that mirror :mod:`openpyxl.chart`.

Eight chart kinds ship with full openpyxl per-type feature depth:

* :class:`BarChart` (column + bar)
* :class:`LineChart`
* :class:`PieChart` and :class:`DoughnutChart`
* :class:`AreaChart`
* :class:`ScatterChart`
* :class:`BubbleChart`
* :class:`RadarChart`

3-D variants (``BarChart3D``, ``LineChart3D``, ``AreaChart3D``,
``PieChart3D``), :class:`ProjectedPieChart`, ``StockChart`` and
``SurfaceChart`` are exposed as stubs raising ``NotImplementedError`` —
deferred to v1.6.1.

Construct charts exactly as you would with openpyxl::

    from wolfxl.chart import BarChart, Reference

    chart = BarChart()
    chart.title = "Sales"
    chart.style = 10
    data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=6)
    chart.add_data(data, titles_from_data=True)
    ws.add_chart(chart, "D2")
"""

from __future__ import annotations

from wolfxl._compat import _make_stub

from .area_chart import AreaChart, AreaChart3D
from .bar_chart import BarChart, BarChart3D
from .bubble_chart import BubbleChart
from .doughnut_chart import DoughnutChart
from .line_chart import LineChart, LineChart3D
from .pie_chart import PieChart, PieChart3D, ProjectedPieChart
from .radar_chart import RadarChart
from .reference import Reference
from .scatter_chart import ScatterChart
from .series import Series, SeriesFactory, SeriesLabel, XYSeries

# Stock + Surface charts are still stubbed (v1.6.1 follow-up).
_STOCK_HINT = (
    "StockChart is deferred to v1.6.1. Fall back to openpyxl for high-low-close "
    "/ open-high-low-close charts."
)
_SURFACE_HINT = (
    "SurfaceChart / SurfaceChart3D are deferred to v1.6.1. Fall back to "
    "openpyxl for surface plots."
)

StockChart = _make_stub("StockChart", _STOCK_HINT)
SurfaceChart = _make_stub("SurfaceChart", _SURFACE_HINT)
SurfaceChart3D = _make_stub("SurfaceChart3D", _SURFACE_HINT)


__all__ = [
    "AreaChart",
    "AreaChart3D",
    "BarChart",
    "BarChart3D",
    "BubbleChart",
    "DoughnutChart",
    "LineChart",
    "LineChart3D",
    "PieChart",
    "PieChart3D",
    "ProjectedPieChart",
    "RadarChart",
    "Reference",
    "ScatterChart",
    "Series",
    "SeriesFactory",
    "SeriesLabel",
    "StockChart",
    "SurfaceChart",
    "SurfaceChart3D",
    "XYSeries",
]
