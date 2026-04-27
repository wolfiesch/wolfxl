"""`<c:marker>` and `<c:dPt>` — series markers + per-point overrides.

Mirrors :mod:`openpyxl.chart.marker`.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any

from .shapes import GraphicalProperties


_VALID_SYMBOLS = (
    None,
    "none",
    "auto",
    "circle",
    "dash",
    "diamond",
    "dot",
    "picture",
    "plus",
    "square",
    "star",
    "triangle",
    "x",
)


class Marker:
    """`<c:marker>` — symbol, size, and per-marker shape properties."""

    __slots__ = ("symbol", "size", "spPr")

    def __init__(
        self,
        symbol: str | None = None,
        size: int | None = None,
        spPr: GraphicalProperties | None = None,
    ) -> None:
        if symbol not in _VALID_SYMBOLS:
            raise ValueError(f"symbol={symbol!r} not in {_VALID_SYMBOLS}")
        if size is not None and not (2 <= size <= 72):
            raise ValueError(f"size={size} must be in [2, 72]")
        self.symbol = symbol
        self.size = size
        self.spPr = spPr

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.symbol is not None:
            d["symbol"] = self.symbol
        if self.size is not None:
            d["size"] = self.size
        if self.spPr is not None:
            d["spPr"] = self.spPr.to_dict()
        return d


class DataPoint:
    """`<c:dPt>` — per-data-point override (colour, marker, explosion …)."""

    __slots__ = (
        "idx",
        "invertIfNegative",
        "marker",
        "bubble3D",
        "explosion",
        "spPr",
    )

    def __init__(
        self,
        idx: int | None = None,
        invertIfNegative: bool | None = None,
        marker: Marker | None = None,
        bubble3D: bool | None = None,
        explosion: int | None = None,
        spPr: GraphicalProperties | None = None,
    ) -> None:
        self.idx = idx
        self.invertIfNegative = invertIfNegative
        self.marker = marker
        self.bubble3D = bubble3D
        self.explosion = explosion
        self.spPr = spPr

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.idx is not None:
            d["idx"] = self.idx
        if self.invertIfNegative is not None:
            d["invertIfNegative"] = self.invertIfNegative
        if self.marker is not None:
            d["marker"] = self.marker.to_dict()
        if self.bubble3D is not None:
            d["bubble3D"] = self.bubble3D
        if self.explosion is not None:
            d["explosion"] = self.explosion
        if self.spPr is not None:
            d["spPr"] = self.spPr.to_dict()
        return d


__all__ = ["DataPoint", "Marker"]
