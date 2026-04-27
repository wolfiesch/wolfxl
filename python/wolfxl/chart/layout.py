"""`<c:layout>` and `<c:manualLayout>` — chart-element placement.

Mirrors :class:`openpyxl.chart.layout.Layout`. All fields are optional;
the Rust emitter omits the element entirely when every slot is None.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any


class ManualLayout:
    """`<c:manualLayout>` — explicit placement (xMode/yMode + x/y/w/h)."""

    __slots__ = ("layoutTarget", "xMode", "yMode", "wMode", "hMode", "x", "y", "w", "h")

    _allowed = {
        "layoutTarget": ("inner", "outer"),
        "xMode": ("edge", "factor"),
        "yMode": ("edge", "factor"),
        "wMode": ("edge", "factor"),
        "hMode": ("edge", "factor"),
    }

    def __init__(
        self,
        layoutTarget: str | None = None,
        xMode: str | None = None,
        yMode: str | None = None,
        wMode: str = "factor",
        hMode: str = "factor",
        x: float | None = None,
        y: float | None = None,
        w: float | None = None,
        h: float | None = None,
    ) -> None:
        for key, val in (
            ("layoutTarget", layoutTarget),
            ("xMode", xMode),
            ("yMode", yMode),
            ("wMode", wMode),
            ("hMode", hMode),
        ):
            if val is not None and val not in self._allowed[key]:
                raise ValueError(f"{key}={val!r} not in {self._allowed[key]}")
        self.layoutTarget = layoutTarget
        self.xMode = xMode
        self.yMode = yMode
        self.wMode = wMode
        self.hMode = hMode
        self.x = x
        self.y = y
        self.w = w
        self.h = h

    # openpyxl aliases
    @property
    def width(self) -> float | None:
        return self.w

    @width.setter
    def width(self, v: float | None) -> None:
        self.w = v

    @property
    def height(self) -> float | None:
        return self.h

    @height.setter
    def height(self, v: float | None) -> None:
        self.h = v

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        for slot in self.__slots__:
            v = getattr(self, slot)
            if v is not None:
                d[slot] = v
        return d


class Layout:
    """`<c:layout>` — wraps an optional :class:`ManualLayout`."""

    __slots__ = ("manualLayout",)

    def __init__(self, manualLayout: ManualLayout | None = None) -> None:
        self.manualLayout = manualLayout

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.manualLayout is not None:
            d["manualLayout"] = self.manualLayout.to_dict()
        return d


__all__ = ["Layout", "ManualLayout"]
