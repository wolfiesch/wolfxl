"""Anchor primitives for ``Image`` placement — Sprint Λ Pod-β (RFC-045).

Mirrors the minimum slice of ``openpyxl.drawing.spreadsheet_drawing`` and
``openpyxl.drawing.xdr`` needed by ``Worksheet.add_image``: the three
anchor flavours plus the two coordinate value objects.

These are passive containers — they do not allocate part ids and do not
touch the writer / patcher. ``Worksheet.add_image`` reads their fields
when it queues the image.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

# ---------------------------------------------------------------------------
# XDR coordinate primitives
# ---------------------------------------------------------------------------

@dataclass
class XDRPoint2D:
    """An (x, y) point in EMU (English Metric Units; 914400 EMU = 1 inch)."""

    x: int = 0
    y: int = 0


@dataclass
class XDRPositiveSize2D:
    """A positive (cx, cy) size in EMU."""

    cx: int = 0
    cy: int = 0


@dataclass
class AnchorMarker:
    """One corner of a cell-relative anchor.

    All four fields are 0-based: ``col=1, row=4`` means column B, row 5
    in 1-based Excel terms (matching openpyxl).
    """

    col: int = 0
    row: int = 0
    colOff: int = 0  # noqa: N815 — openpyxl name
    rowOff: int = 0  # noqa: N815 — openpyxl name


# ---------------------------------------------------------------------------
# Anchor types
# ---------------------------------------------------------------------------

@dataclass
class OneCellAnchor:
    """Pin the image's top-left to one cell; size set by image dims.

    ``ext`` (extent in EMU) is optional — when ``None``, the writer
    computes one from the image's pixel dimensions.
    """

    _from: AnchorMarker = None  # type: ignore[assignment]
    ext: XDRPositiveSize2D | None = None

    def __post_init__(self) -> None:
        if self._from is None:
            self._from = AnchorMarker()


@dataclass
class TwoCellAnchor:
    """Anchor the image at top-left AND bottom-right cells; image stretches."""

    _from: AnchorMarker = None  # type: ignore[assignment]
    to: AnchorMarker = None  # type: ignore[assignment]
    editAs: str = "oneCell"  # noqa: N815 — openpyxl uses "oneCell"/"twoCell"/"absolute"

    def __post_init__(self) -> None:
        if self._from is None:
            self._from = AnchorMarker()
        if self.to is None:
            self.to = AnchorMarker()


@dataclass
class AbsoluteAnchor:
    """EMU-coordinate anchor; image position is independent of cells."""

    pos: XDRPoint2D = None  # type: ignore[assignment]
    ext: XDRPositiveSize2D = None  # type: ignore[assignment]

    def __post_init__(self) -> None:
        if self.pos is None:
            self.pos = XDRPoint2D()
        if self.ext is None:
            self.ext = XDRPositiveSize2D()


@dataclass
class SpreadsheetDrawing:
    """Passive drawing container matching openpyxl's constructor shape."""

    twoCellAnchor: list[Any] = field(default_factory=list)  # noqa: N815
    oneCellAnchor: list[Any] = field(default_factory=list)  # noqa: N815
    absoluteAnchor: list[Any] = field(default_factory=list)  # noqa: N815
    charts: list[Any] = field(default_factory=list)
    images: list[Any] = field(default_factory=list)
    _rels: list[Any] = field(default_factory=list)


__all__ = [
    "AbsoluteAnchor",
    "AnchorMarker",
    "OneCellAnchor",
    "SpreadsheetDrawing",
    "TwoCellAnchor",
    "XDRPoint2D",
    "XDRPositiveSize2D",
]
