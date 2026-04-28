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

    def to_dict(self) -> dict[str, Any] | None:
        """Emit the §10.5 shape (snake_case) — flat ``{x, y, w, h, *_mode, layout_target}``.

        Returns None when every field is None so the parent's
        ``layout`` key is omitted entirely.
        """
        d: dict[str, Any] = {
            "x": self.x,
            "y": self.y,
            "w": self.w,
            "h": self.h,
            "layout_target": self.layoutTarget,
            "x_mode": self.xMode,
            "y_mode": self.yMode,
            "w_mode": self.wMode,
            "h_mode": self.hMode,
        }
        if all(v is None for v in d.values()):
            return None
        return d


class Layout:
    """`<c:layout>` — wraps an optional :class:`ManualLayout`."""

    __slots__ = ("manualLayout",)

    def __init__(self, manualLayout: ManualLayout | None = None) -> None:
        self.manualLayout = manualLayout

    def to_dict(self) -> dict[str, Any] | None:
        """Emit the §10.5 shape: pass through ManualLayout's flat fields.

        Returns None when there's no manual layout (so the parent omits
        the layout key entirely per §10.5).
        """
        if self.manualLayout is None:
            return None
        return self.manualLayout.to_dict()


__all__ = ["Layout", "ManualLayout"]
