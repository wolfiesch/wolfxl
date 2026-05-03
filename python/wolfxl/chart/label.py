"""`<c:dLbl>` and `<c:dLbls>` — data labels (per-point + series-level).

Mirrors :mod:`openpyxl.chart.label`.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any

from .shapes import GraphicalProperties
from .text import (
    CharacterProperties,
    Paragraph,
    ParagraphProperties,
    RegularTextRun,
    RichText,
)


_VALID_POSITIONS = (
    None,
    "bestFit",
    "b",
    "ctr",
    "inBase",
    "inEnd",
    "l",
    "outEnd",
    "r",
    "t",
)


class _DataLabelBase:
    """Shared fields between :class:`DataLabel` and :class:`DataLabelList`."""

    __slots__ = (
        "numFmt",
        "spPr",
        "txPr",
        "dLblPos",
        "showLegendKey",
        "showVal",
        "showCatName",
        "showSerName",
        "showPercent",
        "showBubbleSize",
        "showLeaderLines",
        "separator",
    )

    def __init__(
        self,
        numFmt: str | None = None,
        spPr: GraphicalProperties | None = None,
        txPr: RichText | None = None,
        dLblPos: str | None = None,
        showLegendKey: bool | None = None,
        showVal: bool | None = None,
        showCatName: bool | None = None,
        showSerName: bool | None = None,
        showPercent: bool | None = None,
        showBubbleSize: bool | None = None,
        showLeaderLines: bool | None = None,
        separator: str | None = None,
        position: str | None = None,
        rich: Any | None = None,
    ) -> None:
        # ``position`` is an openpyxl-style alias for ``dLblPos`` -
        # accept either, prefer the one explicitly passed.
        if position is not None and dLblPos is None:
            dLblPos = position
        if dLblPos not in _VALID_POSITIONS:
            raise ValueError(f"dLblPos={dLblPos!r} not in {_VALID_POSITIONS}")
        self.numFmt = numFmt
        self.spPr = spPr
        # ``rich`` is a wolfxl convenience: accepts a CellRichText (or any
        # iterable of str/TextBlock runs) and inflates it into a wolfxl
        # ``RichText`` body that flows into ``<c:txPr>`` so the data
        # labels emit per-run formatting (bold/italic/color/size/font).
        if rich is not None and txPr is None:
            txPr = _coerce_rich_to_txpr(rich)
        self.txPr = txPr
        self.dLblPos = dLblPos
        self.showLegendKey = showLegendKey
        self.showVal = showVal
        self.showCatName = showCatName
        self.showSerName = showSerName
        self.showPercent = showPercent
        self.showBubbleSize = showBubbleSize
        self.showLeaderLines = showLeaderLines
        self.separator = separator

    @property
    def position(self) -> str | None:
        return self.dLblPos

    @position.setter
    def position(self, v: str | None) -> None:
        if v not in _VALID_POSITIONS:
            raise ValueError(f"position={v!r} not in {_VALID_POSITIONS}")
        self.dLblPos = v

    def _base_to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.numFmt is not None:
            d["numFmt"] = self.numFmt
        if self.spPr is not None:
            d["spPr"] = self.spPr.to_dict()
        if self.txPr is not None:
            d["txPr"] = self.txPr.to_dict()
        if self.dLblPos is not None:
            d["dLblPos"] = self.dLblPos
        for slot in (
            "showLegendKey",
            "showVal",
            "showCatName",
            "showSerName",
            "showPercent",
            "showBubbleSize",
            "showLeaderLines",
        ):
            v = getattr(self, slot)
            if v is not None:
                d[slot] = v
        if self.separator is not None:
            d["separator"] = self.separator
        return d


class DataLabel(_DataLabelBase):
    """`<c:dLbl>` — single per-point label override."""

    __slots__ = ("idx",)

    def __init__(self, idx: int = 0, **kwargs: Any) -> None:
        super().__init__(**kwargs)
        self.idx = idx

    def to_dict(self) -> dict[str, Any]:
        d = self._base_to_dict()
        d["idx"] = self.idx
        return d


class DataLabelList(_DataLabelBase):
    """`<c:dLbls>` — series-wide label defaults + per-point overrides."""

    __slots__ = ("dLbl", "delete")

    def __init__(
        self,
        dLbl: list[DataLabel] | tuple[DataLabel, ...] = (),
        delete: bool | None = None,
        **kwargs: Any,
    ) -> None:
        super().__init__(**kwargs)
        self.dLbl = list(dLbl)
        self.delete = delete

    def to_dict(self) -> dict[str, Any]:
        d = self._base_to_dict()
        if self.dLbl:
            d["dLbl"] = [lbl.to_dict() for lbl in self.dLbl]
        if self.delete is not None:
            d["delete"] = self.delete
        return d


def _coerce_rich_to_txpr(value: Any) -> RichText:
    """Inflate a rich-text payload into a wolfxl :class:`RichText`.

    Accepts either an existing :class:`RichText` (returned verbatim), a
    ``CellRichText`` (iterable of ``str`` + ``TextBlock`` runs), or any
    iterable of the same shape. Each ``TextBlock``'s ``InlineFont`` is
    mapped onto a chart :class:`CharacterProperties` so the run carries
    the same b/i/u/sz/color/font-name flags through the Rust emitter.
    """
    if isinstance(value, RichText):
        return value
    runs: list[RegularTextRun] = []
    iterable = value if hasattr(value, "__iter__") else [value]
    for item in iterable:
        if isinstance(item, str):
            runs.append(RegularTextRun(t=item))
            continue
        text = getattr(item, "text", None)
        if text is None:
            continue
        font = getattr(item, "font", None)
        rpr: CharacterProperties | None = None
        if font is not None:
            sz_pt = getattr(font, "sz", None)
            sz: int | None = None
            if sz_pt is not None:
                try:
                    sz = int(round(float(sz_pt) * 100))
                except (TypeError, ValueError):
                    sz = None
            u = getattr(font, "u", None)
            color = getattr(font, "color", None)
            color_hex: str | None = None
            if color is not None:
                if isinstance(color, str):
                    color_hex = color
                else:
                    color_hex = (
                        getattr(color, "rgb", None)
                        or getattr(color, "value", None)
                        or getattr(color, "srgbClr", None)
                    )
                    if color_hex is not None and not isinstance(color_hex, str):
                        color_hex = str(color_hex)
            rpr = CharacterProperties(
                b=getattr(font, "b", None),
                i=getattr(font, "i", None),
                u=u,
                strike="sngStrike" if getattr(font, "strike", None) else None,
                sz=sz,
                solidFill=color_hex,
                latin=getattr(font, "rFont", None),
            )
        runs.append(RegularTextRun(rPr=rpr, t=text))
    paragraph = Paragraph(
        pPr=ParagraphProperties(defRPr=CharacterProperties()),
        r=runs,
    )
    return RichText(p=[paragraph])


__all__ = ["DataLabel", "DataLabelList"]
