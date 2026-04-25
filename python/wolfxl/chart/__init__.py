"""Shim for ``openpyxl.chart``.

Charts in existing workbooks are preserved by wolfxl's modify mode on
round-trip. Creating new charts from Python is not supported - fall back
to openpyxl for chart-creation code paths.
"""

from __future__ import annotations

from wolfxl._compat import _make_stub

_HINT = (
    "Charts can be preserved via modify mode but wolfxl cannot construct them. "
    "Use openpyxl for chart-creation workflows."
)

BarChart = _make_stub("BarChart", _HINT)
LineChart = _make_stub("LineChart", _HINT)
PieChart = _make_stub("PieChart", _HINT)
ScatterChart = _make_stub("ScatterChart", _HINT)
AreaChart = _make_stub("AreaChart", _HINT)
Reference = _make_stub("Reference", _HINT)
Series = _make_stub("Series", _HINT)

__all__ = [
    "AreaChart",
    "BarChart",
    "LineChart",
    "PieChart",
    "Reference",
    "ScatterChart",
    "Series",
]
