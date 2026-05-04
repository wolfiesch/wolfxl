"""`SurfaceChart` and `SurfaceChart3D`.

Both surface variants emit ``<c:surfaceChart>`` (2D) or
``<c:surface3DChart>`` (3D). The 2D form is a contour-style heatmap
projection; the 3D form is a perspective-rendered surface. Both
expose a ``wireframe: bool`` toggle.
"""

from __future__ import annotations

from typing import Any

from ._chart import ChartBase
from .axis import NumericAxis, SeriesAxis, TextAxis
from .label import DataLabelList


class _SurfaceChartBase(ChartBase):
    """Shared state between flat and 3D surface charts."""

    _series_type = "surface"

    def __init__(
        self,
        wireframe: bool = True,
        ser: list[Any] | tuple[Any, ...] = (),
        dLbls: DataLabelList | None = None,
        **kw: Any,
    ) -> None:
        self.wireframe = bool(wireframe)
        self.dLbls = dLbls
        super().__init__(**kw)
        self.ser = list(ser)
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        self.z_axis = SeriesAxis()


class SurfaceChart(_SurfaceChartBase):
    """A 2-D surface (contour) chart."""

    tagname = "surfaceChart"

    def _chart_type_specific_keys(self) -> dict[str, Any]:
        d: dict[str, Any] = {"wireframe": self.wireframe}
        if self.dLbls is not None:
            from .series import _dlbls_to_snake
            d["data_labels"] = _dlbls_to_snake(self.dLbls.to_dict())
        return d


class SurfaceChart3D(_SurfaceChartBase):
    """A 3-D surface chart.

    Defaults: rot_x=15, rot_y=20, perspective=30, depth_percent=100.
    """

    tagname = "surface3DChart"

    def __init__(
        self,
        wireframe: bool = True,
        view_3d: dict[str, Any] | None = None,
        **kw: Any,
    ) -> None:
        super().__init__(wireframe=wireframe, **kw)
        self.view_3d = {
            "rot_x": 15,
            "rot_y": 20,
            "perspective": 30,
            "right_angle_axes": False,
            "depth_percent": 100,
        }
        if view_3d is not None:
            self.view_3d.update(view_3d)

    def _chart_type_specific_keys(self) -> dict[str, Any]:
        d: dict[str, Any] = {"wireframe": self.wireframe}
        v3d = {k: v for k, v in self.view_3d.items() if v is not None}
        if v3d:
            d["view_3d"] = v3d
        if self.dLbls is not None:
            from .series import _dlbls_to_snake
            d["data_labels"] = _dlbls_to_snake(self.dLbls.to_dict())
        return d


__all__ = ["SurfaceChart", "SurfaceChart3D"]
