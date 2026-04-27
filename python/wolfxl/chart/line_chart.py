"""`LineChart` and `LineChart3D` — line plots.

Mirrors :class:`openpyxl.chart.line_chart.LineChart`. ``LineChart3D`` is
deferred to v1.6.1.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any

from ._chart import ChartBase
from .axis import ChartLines, NumericAxis, TextAxis
from .label import DataLabelList


_VALID_GROUPING = ("percentStacked", "standard", "stacked")


class _LineChartBase(ChartBase):
    """Shared state between flat and 3D line charts."""

    _series_type = "line"

    def __init__(
        self,
        grouping: str = "standard",
        varyColors: bool | None = None,
        ser: list[Any] | tuple[Any, ...] = (),
        dLbls: DataLabelList | None = None,
        dropLines: ChartLines | None = None,
        **kw: Any,
    ) -> None:
        if grouping not in _VALID_GROUPING:
            raise ValueError(f"grouping={grouping!r} not in {_VALID_GROUPING}")
        self.grouping = grouping
        self.vary_colors = varyColors
        self.dLbls = dLbls
        self.dropLines = dropLines
        super().__init__(**kw)
        self.ser = list(ser)

    @property
    def varyColors(self) -> bool | None:
        return self.vary_colors

    @varyColors.setter
    def varyColors(self, v: bool | None) -> None:
        self.vary_colors = v


class LineChart(_LineChartBase):
    """A flat (2D) line chart."""

    tagname = "lineChart"

    def __init__(
        self,
        hiLowLines: ChartLines | None = None,
        upDownBars: Any | None = None,
        marker: bool | None = None,
        smooth: bool | None = None,
        **kw: Any,
    ) -> None:
        super().__init__(**kw)
        self.hiLowLines = hiLowLines
        self.upDownBars = upDownBars
        self.marker = marker
        self.smooth = smooth
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()

    def _chart_dict_extras(self) -> dict[str, Any]:
        d: dict[str, Any] = {"grouping": self.grouping}
        if self.vary_colors is not None:
            d["varyColors"] = self.vary_colors
        if self.dropLines is not None:
            d["dropLines"] = self.dropLines.to_dict()
        if self.hiLowLines is not None:
            d["hiLowLines"] = self.hiLowLines.to_dict()
        if self.upDownBars is not None:
            d["upDownBars"] = (
                self.upDownBars.to_dict()
                if hasattr(self.upDownBars, "to_dict")
                else self.upDownBars
            )
        if self.marker is not None:
            d["marker"] = self.marker
        if self.smooth is not None:
            d["smooth"] = self.smooth
        if self.dLbls is not None:
            d["dLbls"] = self.dLbls.to_dict()
        return d


class LineChart3D(LineChart):
    """3D line chart — stub; full support deferred to v1.6.1."""

    tagname = "line3DChart"

    def __init__(self, *args: Any, **kw: Any) -> None:
        raise NotImplementedError(
            "LineChart3D is not yet implemented in wolfxl (deferred to v1.6.1). "
            "Use LineChart for 2D line plots, or fall back to openpyxl for "
            "3D variants."
        )


__all__ = ["LineChart", "LineChart3D"]
