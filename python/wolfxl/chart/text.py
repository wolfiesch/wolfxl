"""Rich-text primitives for chart titles, data labels, and axis labels.

A trimmed mirror of :mod:`openpyxl.chart.text` — we keep the public API
surface (``RichText``, ``Text``, ``Paragraph``, ``RegularTextRun``,
``CharacterProperties``, ``ParagraphProperties``) but implement them as
plain attribute carriers so the Rust emitter can serialise them with
minimal Python side-validation.
"""

from __future__ import annotations

from typing import Any


class CharacterProperties:
    """`<a:rPr>` — run-level rich-text properties.

    ``lang`` (e.g. ``"en-US"``), ``sz`` (font size in 1/100 pt — 1100=11pt),
    ``b`` (bold), ``i`` (italic), ``u`` (underline style), ``strike``,
    ``solidFill`` (hex colour, e.g. ``"FF0000"``), ``latin`` (font face),
    ``baseline`` (super/subscript offset).
    """

    __slots__ = ("lang", "sz", "b", "i", "u", "strike", "solidFill", "latin", "baseline")

    def __init__(
        self,
        lang: str | None = None,
        sz: int | None = None,
        b: bool | None = None,
        i: bool | None = None,
        u: str | None = None,
        strike: str | None = None,
        solidFill: str | None = None,
        latin: str | None = None,
        baseline: int | None = None,
    ) -> None:
        self.lang = lang
        self.sz = sz
        self.b = b
        self.i = i
        self.u = u
        self.strike = strike
        self.solidFill = solidFill
        self.latin = latin
        self.baseline = baseline

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        for slot in self.__slots__:
            v = getattr(self, slot)
            if v is not None:
                d[slot] = v
        return d


class ParagraphProperties:
    """`<a:pPr>` — paragraph-level properties (alignment, default run props)."""

    __slots__ = ("algn", "defRPr")

    def __init__(
        self,
        algn: str | None = None,
        defRPr: CharacterProperties | None = None,
    ) -> None:
        self.algn = algn
        self.defRPr = defRPr

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.algn is not None:
            d["algn"] = self.algn
        if self.defRPr is not None:
            d["defRPr"] = self.defRPr.to_dict()
        return d


class RegularTextRun:
    """`<a:r>` — a single text run with optional formatting."""

    __slots__ = ("rPr", "t")

    def __init__(
        self,
        rPr: CharacterProperties | None = None,
        t: str = "",
    ) -> None:
        self.rPr = rPr
        self.t = t

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"t": self.t}
        if self.rPr is not None:
            d["rPr"] = self.rPr.to_dict()
        return d


class LineBreak:
    """`<a:br>` — explicit line break inside a paragraph."""

    __slots__ = ("rPr",)

    def __init__(self, rPr: CharacterProperties | None = None) -> None:
        self.rPr = rPr

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"_kind": "br"}
        if self.rPr is not None:
            d["rPr"] = self.rPr.to_dict()
        return d


class Paragraph:
    """`<a:p>` — a paragraph with optional ``pPr`` and a sequence of runs."""

    __slots__ = ("pPr", "r")

    def __init__(
        self,
        pPr: ParagraphProperties | None = None,
        r: list[RegularTextRun] | None = None,
    ) -> None:
        self.pPr = pPr
        self.r = list(r) if r else []

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.pPr is not None:
            d["pPr"] = self.pPr.to_dict()
        if self.r:
            d["r"] = [r.to_dict() for r in self.r]
        return d


class RichTextProperties:
    """`<a:bodyPr>` — chart-text body properties.

    Optional attributes the chart spec ships:
    ``rot``, ``spcFirstLastPara``, ``vertOverflow``, ``vert``, ``wrap``,
    ``anchor``, ``anchorCtr``. We carry them as plain attributes.
    """

    __slots__ = ("rot", "spcFirstLastPara", "vertOverflow", "vert", "wrap", "anchor", "anchorCtr")

    def __init__(self, **kwargs: Any) -> None:
        for slot in self.__slots__:
            setattr(self, slot, kwargs.get(slot))

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        for slot in self.__slots__:
            v = getattr(self, slot)
            if v is not None:
                d[slot] = v
        return d


class RichText:
    """`<c:rich>` — the rich-text body of a title or label."""

    __slots__ = ("bodyPr", "lstStyle", "p")

    def __init__(
        self,
        bodyPr: RichTextProperties | None = None,
        lstStyle: Any | None = None,
        p: list[Paragraph] | None = None,
    ) -> None:
        self.bodyPr = bodyPr if bodyPr is not None else RichTextProperties()
        self.lstStyle = lstStyle
        self.p = list(p) if p else [Paragraph()]

    # openpyxl alias
    @property
    def paragraphs(self) -> list[Paragraph]:
        return self.p

    @paragraphs.setter
    def paragraphs(self, value: list[Paragraph]) -> None:
        self.p = list(value)

    def to_dict(self) -> dict[str, Any]:
        return {
            "bodyPr": self.bodyPr.to_dict() if self.bodyPr else {},
            "p": [para.to_dict() for para in self.p],
        }


class Text:
    """`<c:tx>` — title-text container; either a strRef or a rich body."""

    __slots__ = ("strRef", "rich")

    def __init__(self, strRef: Any | None = None, rich: RichText | None = None) -> None:
        self.strRef = strRef
        self.rich = rich if rich is not None else RichText()

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.strRef is not None:
            from .data_source import StrRef  # local to avoid cycle
            if isinstance(self.strRef, StrRef):
                d["strRef"] = self.strRef.to_dict()
            else:
                d["strRef"] = self.strRef
        else:
            d["rich"] = self.rich.to_dict()
        return d


__all__ = [
    "CharacterProperties",
    "LineBreak",
    "Paragraph",
    "ParagraphProperties",
    "RegularTextRun",
    "RichText",
    "RichTextProperties",
    "Text",
]
