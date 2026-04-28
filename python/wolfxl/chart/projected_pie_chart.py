"""`ProjectedPieChart` — pie-of-pie / bar-of-pie (RFC-046 §11.4).

Maps to ``<c:ofPieChart>``. Carries an ``of_pie_type`` (``"bar"`` or
``"pie"``), ``split_type`` selector, and ``split_pos`` /
``second_pie_size`` per the OOXML spec.

Sprint Μ-prime Pod-β′ (RFC-046 §11.4).
"""

from __future__ import annotations

from typing import Any

from ._chart import ChartBase
from .label import DataLabelList


_VALID_OF_PIE_TYPE = ("bar", "pie")
_VALID_SPLIT_TYPE = ("auto", "pos", "percent", "val", "cust")


class ProjectedPieChart(ChartBase):
    """`<c:ofPieChart>` — bar-of-pie / pie-of-pie."""

    tagname = "ofPieChart"
    _series_type = "pie"

    def __init__(
        self,
        of_pie_type: str = "pie",
        split_type: str = "auto",
        split_pos: int | None = None,
        second_pie_size: int | None = None,
        ser: list[Any] | tuple[Any, ...] = (),
        dLbls: DataLabelList | None = None,
        varyColors: bool | None = True,
        firstSliceAng: int = 0,
        **kw: Any,
    ) -> None:
        if of_pie_type not in _VALID_OF_PIE_TYPE:
            raise ValueError(
                f"of_pie_type={of_pie_type!r} not in {_VALID_OF_PIE_TYPE}"
            )
        if split_type not in _VALID_SPLIT_TYPE:
            raise ValueError(
                f"split_type={split_type!r} not in {_VALID_SPLIT_TYPE}"
            )
        if second_pie_size is not None and not (5 <= second_pie_size <= 200):
            raise ValueError(
                f"second_pie_size={second_pie_size} must be in [5, 200]"
            )
        if not (0 <= firstSliceAng <= 360):
            raise ValueError(f"firstSliceAng={firstSliceAng} must be in [0, 360]")

        self.of_pie_type = of_pie_type
        self.split_type = split_type
        self.split_pos = split_pos
        self.second_pie_size = second_pie_size
        self.firstSliceAng = firstSliceAng
        self.dLbls = dLbls
        self.vary_colors = varyColors
        super().__init__(**kw)
        self.ser = list(ser)

    @property
    def first_slice_ang(self) -> int:
        return self.firstSliceAng

    @first_slice_ang.setter
    def first_slice_ang(self, v: int) -> None:
        if not (0 <= v <= 360):
            raise ValueError(f"first_slice_ang={v} must be in [0, 360]")
        self.firstSliceAng = v

    def _chart_type_specific_keys(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "of_pie_type": self.of_pie_type,
            "split_type": self.split_type,
            "first_slice_ang": self.firstSliceAng,
        }
        if self.split_pos is not None:
            d["split_pos"] = self.split_pos
        if self.second_pie_size is not None:
            d["second_pie_size"] = self.second_pie_size
        if self.dLbls is not None:
            from .series import _dlbls_to_snake
            d["data_labels"] = _dlbls_to_snake(self.dLbls.to_dict())
        return d


__all__ = ["ProjectedPieChart"]
