"""DrawingML fill primitives — ColorChoice, SolidColorFillProperties, etc.

Mirrors :mod:`openpyxl.drawing.fill` but lightweight: the Rust chart
emitter only consumes string colour values, so most of openpyxl's
descriptor soup is collapsed to a tiny dataclass-style holder that
``GraphicalProperties.solidFill`` accepts in place of a raw string.

Sprint Μ Pod-β (RFC-046) — added during integrator finalize to satisfy
test contract surfaces that import from ``wolfxl.drawing.fill`` directly
(matching openpyxl's import path).
"""

from __future__ import annotations

from typing import Any


class ColorChoice:
    """`<a:srgbClr>` / `<a:schemeClr>` / `<a:prstClr>` colour choice.

    A user passes ``ColorChoice(srgbClr="00FF00")`` to spec a literal
    RGB colour, or ``schemeClr="accent1"`` for a theme-resolved colour.

    The Rust emitter consumes the resolved hex string (or scheme name);
    when this object is used as ``GraphicalProperties.solidFill``, the
    chart serialiser unwraps it via ``__str__``.
    """

    __slots__ = ("srgbClr", "schemeClr", "prstClr", "hslClr", "sysClr")

    def __init__(
        self,
        srgbClr: str | None = None,
        schemeClr: str | None = None,
        prstClr: str | None = None,
        hslClr: str | None = None,
        sysClr: str | None = None,
    ) -> None:
        self.srgbClr = srgbClr
        self.schemeClr = schemeClr
        self.prstClr = prstClr
        self.hslClr = hslClr
        self.sysClr = sysClr

    def __str__(self) -> str:
        # The chart emitter consumes solidFill as a string today; pick
        # the first non-None choice. Theme resolution is the writer's
        # responsibility (see Pod-α charts.rs).
        for v in (self.srgbClr, self.schemeClr, self.prstClr, self.hslClr, self.sysClr):
            if v:
                return v
        return ""

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.srgbClr is not None:
            d["srgbClr"] = self.srgbClr
        if self.schemeClr is not None:
            d["schemeClr"] = self.schemeClr
        if self.prstClr is not None:
            d["prstClr"] = self.prstClr
        if self.hslClr is not None:
            d["hslClr"] = self.hslClr
        if self.sysClr is not None:
            d["sysClr"] = self.sysClr
        return d


__all__ = ["ColorChoice"]
