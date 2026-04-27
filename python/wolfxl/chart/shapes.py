"""Drawing-ML graphical properties (`<c:spPr>` / `<a:spPr>`).

Mirrors :class:`openpyxl.chart.shapes.GraphicalProperties`. The DrawingML
spec for chart spPr is restrictive — no custGeom/prstGeom, no scene3d,
no bwMode — so this implementation only carries the fields that survive
chart-context serialisation.

Properties are stored as plain attributes; the ``to_dict()`` method emits
the camelCase XML names that the Rust emitter expects.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any


class LineProperties:
    """Drawing-ML `<a:ln>` line properties — width, dash, fill colour.

    Kept lightweight (vs. openpyxl's full ``LineProperties`` descriptor
    soup) because the chart emitter only reads ``w``, ``cap``, ``cmpd``,
    ``solidFill``, and ``prstDash``.
    """

    __slots__ = ("w", "cap", "cmpd", "solidFill", "prstDash", "noFill")

    def __init__(
        self,
        w: int | None = None,
        cap: str | None = None,
        cmpd: str | None = None,
        solidFill: str | None = None,
        prstDash: str | None = None,
        noFill: bool | None = None,
    ) -> None:
        self.w = w
        self.cap = cap
        self.cmpd = cmpd
        self.solidFill = solidFill
        self.prstDash = prstDash
        self.noFill = noFill

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.w is not None:
            d["w"] = self.w
        if self.cap is not None:
            d["cap"] = self.cap
        if self.cmpd is not None:
            d["cmpd"] = self.cmpd
        if self.solidFill is not None:
            d["solidFill"] = self.solidFill
        if self.prstDash is not None:
            d["prstDash"] = self.prstDash
        if self.noFill:
            d["noFill"] = True
        return d


class GraphicalProperties:
    """`<c:spPr>` — chart-side shape properties.

    Mirrors openpyxl's :class:`GraphicalProperties` but carries only the
    fields the chart-XML emitter uses: ``noFill``, ``solidFill``,
    ``ln``, plus the gradient/pattern fill placeholders (Rust side
    decides whether to emit them — for v1.5 we round-trip these as
    opaque dicts).
    """

    __slots__ = (
        "noFill",
        "solidFill",
        "gradFill",
        "pattFill",
        "ln",
    )

    def __init__(
        self,
        noFill: bool | None = None,
        solidFill: str | None = None,
        gradFill: dict[str, Any] | None = None,
        pattFill: dict[str, Any] | None = None,
        ln: LineProperties | None = None,
    ) -> None:
        self.noFill = noFill
        self.solidFill = solidFill
        self.gradFill = gradFill
        self.pattFill = pattFill
        self.ln = ln

    # openpyxl alias
    @property
    def line(self) -> LineProperties | None:
        return self.ln

    @line.setter
    def line(self, value: LineProperties | None) -> None:
        self.ln = value

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.noFill:
            d["noFill"] = True
        if self.solidFill is not None:
            d["solidFill"] = self.solidFill
        if self.gradFill is not None:
            d["gradFill"] = self.gradFill
        if self.pattFill is not None:
            d["pattFill"] = self.pattFill
        if self.ln is not None:
            d["ln"] = self.ln.to_dict()
        return d


__all__ = ["GraphicalProperties", "LineProperties"]
