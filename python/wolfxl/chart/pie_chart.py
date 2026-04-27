"""`PieChart`, `DoughnutChart`, and 3D / projected variants.

Mirrors :class:`openpyxl.chart.pie_chart`. Pie + Doughnut are first-class;
``PieChart3D`` and ``ProjectedPieChart`` are deferred to v1.6.1.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any

from ._chart import ChartBase
from .axis import ChartLines
from .descriptors import NestedGapAmount
from .label import DataLabelList


class _PieChartBase(ChartBase):
    """Shared state between PieChart and DoughnutChart variants."""

    _series_type = "pie"

    def __init__(
        self,
        varyColors: bool | None = True,
        ser: list[Any] | tuple[Any, ...] = (),
        dLbls: DataLabelList | None = None,
        **kw: Any,
    ) -> None:
        self.vary_colors = varyColors
        self.dLbls = dLbls
        super().__init__(**kw)
        self.ser = list(ser)

    @property
    def varyColors(self) -> bool | None:
        return self.vary_colors

    @varyColors.setter
    def varyColors(self, v: bool | None) -> None:
        self.vary_colors = v


class PieChart(_PieChartBase):
    """A pie chart — single series, slice per category."""

    tagname = "pieChart"

    def __init__(self, firstSliceAng: int = 0, **kw: Any) -> None:
        if not (0 <= firstSliceAng <= 360):
            raise ValueError(f"firstSliceAng={firstSliceAng} must be in [0, 360]")
        super().__init__(**kw)
        self.firstSliceAng = firstSliceAng

    @property
    def first_slice_ang(self) -> int:
        return self.firstSliceAng

    @first_slice_ang.setter
    def first_slice_ang(self, v: int) -> None:
        if not (0 <= v <= 360):
            raise ValueError(f"first_slice_ang={v} must be in [0, 360]")
        self.firstSliceAng = v

    def _chart_dict_extras(self) -> dict[str, Any]:
        d: dict[str, Any] = {"firstSliceAng": self.firstSliceAng}
        if self.vary_colors is not None:
            d["varyColors"] = self.vary_colors
        if self.dLbls is not None:
            d["dLbls"] = self.dLbls.to_dict()
        return d


class DoughnutChart(_PieChartBase):
    """A doughnut chart — pie with a hole."""

    tagname = "doughnutChart"

    def __init__(
        self,
        firstSliceAng: int = 0,
        holeSize: int | None = 10,
        **kw: Any,
    ) -> None:
        if not (0 <= firstSliceAng <= 360):
            raise ValueError(f"firstSliceAng={firstSliceAng} must be in [0, 360]")
        if holeSize is not None and not (1 <= holeSize <= 90):
            raise ValueError(f"holeSize={holeSize} must be in [1, 90]")
        super().__init__(**kw)
        self.firstSliceAng = firstSliceAng
        self.holeSize = holeSize

    @property
    def hole_size(self) -> int | None:
        return self.holeSize

    @hole_size.setter
    def hole_size(self, v: int | None) -> None:
        if v is not None and not (1 <= v <= 90):
            raise ValueError(f"hole_size={v} must be in [1, 90]")
        self.holeSize = v

    def _chart_dict_extras(self) -> dict[str, Any]:
        d: dict[str, Any] = {"firstSliceAng": self.firstSliceAng}
        if self.holeSize is not None:
            d["holeSize"] = self.holeSize
        if self.vary_colors is not None:
            d["varyColors"] = self.vary_colors
        if self.dLbls is not None:
            d["dLbls"] = self.dLbls.to_dict()
        return d


class PieChart3D(PieChart):
    """3D pie chart — stub; full support deferred to v1.6.1."""

    tagname = "pie3DChart"

    def __init__(self, *args: Any, **kw: Any) -> None:
        raise NotImplementedError(
            "PieChart3D is not yet implemented in wolfxl (deferred to v1.6.1). "
            "Use PieChart for 2D pie plots, or fall back to openpyxl for "
            "3D variants."
        )


class ProjectedPieChart(PieChart):
    """`ofPieChart` (pie-of-pie / bar-of-pie) — deferred to v1.6.1."""

    tagname = "ofPieChart"

    def __init__(self, *args: Any, **kw: Any) -> None:
        raise NotImplementedError(
            "ProjectedPieChart is not yet implemented in wolfxl (deferred "
            "to v1.6.1). Fall back to openpyxl for pie-of-pie / bar-of-pie."
        )


__all__ = ["DoughnutChart", "PieChart", "PieChart3D", "ProjectedPieChart"]
