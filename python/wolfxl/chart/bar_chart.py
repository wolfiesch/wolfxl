"""`BarChart` and `BarChart3D` — column / bar plots.

Mirrors :class:`openpyxl.chart.bar_chart.BarChart`. ``BarChart3D`` is
deferred to v1.6.1 (RFC-046 §5) — exposed here as a stub that raises
``NotImplementedError`` so users discover the gap immediately.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any

from ._chart import ChartBase
from .axis import ChartLines, NumericAxis, TextAxis
from .descriptors import NestedGapAmount, NestedOverlap
from .label import DataLabelList
from .legend import Legend


_VALID_BAR_DIR = ("bar", "col")
_VALID_GROUPING = ("percentStacked", "clustered", "standard", "stacked")


class _BarChartBase(ChartBase):
    """Shared state between flat and 3D bar charts."""

    _series_type = "bar"

    gapWidth = NestedGapAmount()
    overlap = NestedOverlap()

    def __init__(
        self,
        barDir: str = "col",
        grouping: str = "clustered",
        varyColors: bool | None = None,
        ser: list[Any] | tuple[Any, ...] = (),
        dLbls: DataLabelList | None = None,
        **kw: Any,
    ) -> None:
        if barDir not in _VALID_BAR_DIR:
            raise ValueError(f"barDir={barDir!r} not in {_VALID_BAR_DIR}")
        if grouping not in _VALID_GROUPING:
            raise ValueError(f"grouping={grouping!r} not in {_VALID_GROUPING}")
        self.barDir = barDir
        self.grouping = grouping
        self.vary_colors = varyColors
        self.dLbls = dLbls
        super().__init__(**kw)
        self.ser = list(ser)

    # openpyxl aliases
    @property
    def type(self) -> str:
        return self.barDir

    @type.setter
    def type(self, value: str) -> None:
        if value not in _VALID_BAR_DIR:
            raise ValueError(f"type={value!r} not in {_VALID_BAR_DIR}")
        self.barDir = value

    @property
    def varyColors(self) -> bool | None:
        return self.vary_colors

    @varyColors.setter
    def varyColors(self, v: bool | None) -> None:
        self.vary_colors = v


class BarChart(_BarChartBase):
    """A flat (2D) bar / column chart."""

    tagname = "barChart"

    def __init__(
        self,
        gapWidth: int | None = 150,
        overlap: int | None = None,
        serLines: ChartLines | None = None,
        **kw: Any,
    ) -> None:
        super().__init__(**kw)
        self.gapWidth = gapWidth
        self.overlap = overlap
        self.serLines = serLines
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        self.legend = Legend()

    # openpyxl alias for ``y_axis``
    @property
    def gap_width(self) -> int | None:
        return self.gapWidth

    @gap_width.setter
    def gap_width(self, v: int | None) -> None:
        self.gapWidth = v

    def _chart_dict_extras(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "barDir": self.barDir,
            "grouping": self.grouping,
        }
        if self.vary_colors is not None:
            d["varyColors"] = self.vary_colors
        if self.gapWidth is not None:
            d["gapWidth"] = self.gapWidth
        if self.overlap is not None:
            d["overlap"] = self.overlap
        if self.serLines is not None:
            d["serLines"] = self.serLines.to_dict()
        if self.dLbls is not None:
            d["dLbls"] = self.dLbls.to_dict()
        return d


class BarChart3D(BarChart):
    """3D bar/column chart — stub; full support deferred to v1.6.1.

    Construction raises :class:`NotImplementedError` so users discover the
    gap synchronously rather than at chart-emit time.
    """

    tagname = "bar3DChart"

    def __init__(self, *args: Any, **kw: Any) -> None:
        raise NotImplementedError(
            "BarChart3D is not yet implemented in wolfxl (deferred to v1.6.1). "
            "Use BarChart for 2D column/bar plots, or fall back to openpyxl "
            "for 3D variants."
        )


__all__ = ["BarChart", "BarChart3D"]
