"""`AreaChart` and `AreaChart3D` — filled area plots.

Mirrors :class:`openpyxl.chart.area_chart.AreaChart`. ``AreaChart3D`` is
deferred to v1.6.1.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any

from ._chart import ChartBase
from .axis import ChartLines, NumericAxis, TextAxis
from .label import DataLabelList


_VALID_GROUPING = ("percentStacked", "standard", "stacked")


class _AreaChartBase(ChartBase):
    """Shared state between flat and 3D area charts."""

    _series_type = "area"

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


class AreaChart(_AreaChartBase):
    """A flat (2D) area chart."""

    tagname = "areaChart"

    def __init__(self, **kw: Any) -> None:
        super().__init__(**kw)
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()

    def _chart_dict_extras(self) -> dict[str, Any]:
        d: dict[str, Any] = {"grouping": self.grouping}
        if self.vary_colors is not None:
            d["varyColors"] = self.vary_colors
        if self.dropLines is not None:
            d["dropLines"] = self.dropLines.to_dict()
        if self.dLbls is not None:
            d["dLbls"] = self.dLbls.to_dict()
        return d


class AreaChart3D(AreaChart):
    """3D area chart — stub; full support deferred to v1.6.1."""

    tagname = "area3DChart"

    def __init__(self, *args: Any, **kw: Any) -> None:
        raise NotImplementedError(
            "AreaChart3D is not yet implemented in wolfxl (deferred to v1.6.1). "
            "Use AreaChart for 2D area plots, or fall back to openpyxl for "
            "3D variants."
        )


__all__ = ["AreaChart", "AreaChart3D"]
