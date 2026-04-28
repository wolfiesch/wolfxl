"""`<c:legend>` — chart legend.

Mirrors :class:`openpyxl.chart.legend.Legend`.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any

from .layout import Layout
from .shapes import GraphicalProperties
from .text import RichText


_VALID_LEGEND_POS = ("b", "tr", "l", "r", "t")


class LegendEntry:
    """`<c:legendEntry>` — per-series legend override (delete or restyle)."""

    __slots__ = ("idx", "delete", "txPr")

    def __init__(
        self,
        idx: int = 0,
        delete: bool = False,
        txPr: RichText | None = None,
    ) -> None:
        self.idx = idx
        self.delete = delete
        self.txPr = txPr

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"idx": self.idx, "delete": self.delete}
        if self.txPr is not None:
            d["txPr"] = self.txPr.to_dict()
        return d


class Legend:
    """`<c:legend>` — placement + entry overrides + layout/text properties."""

    __slots__ = ("legendPos", "legendEntry", "layout", "overlay", "spPr", "txPr")

    def __init__(
        self,
        legendPos: str = "r",
        legendEntry: list[LegendEntry] | tuple[LegendEntry, ...] = (),
        layout: Layout | None = None,
        overlay: bool | None = None,
        spPr: GraphicalProperties | None = None,
        txPr: RichText | None = None,
    ) -> None:
        if legendPos not in _VALID_LEGEND_POS:
            raise ValueError(f"legendPos={legendPos!r} not in {_VALID_LEGEND_POS}")
        self.legendPos = legendPos
        self.legendEntry = list(legendEntry)
        self.layout = layout
        self.overlay = overlay
        self.spPr = spPr
        self.txPr = txPr

    @property
    def position(self) -> str:
        return self.legendPos

    @position.setter
    def position(self, value: str) -> None:
        if value not in _VALID_LEGEND_POS:
            raise ValueError(f"position={value!r} not in {_VALID_LEGEND_POS}")
        self.legendPos = value

    def to_dict(self) -> dict[str, Any]:
        """Emit the §10.4 shape: ``{position, overlay, layout}`` (snake_case)."""
        d: dict[str, Any] = {
            "position": self.legendPos,
            "overlay": self.overlay,
            "layout": self.layout.to_dict() if self.layout is not None else None,
        }
        return d


__all__ = ["Legend", "LegendEntry"]
