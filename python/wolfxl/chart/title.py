"""`<c:title>` — chart and axis titles.

Mirrors :class:`openpyxl.chart.title.Title`. The convenience descriptor
``TitleDescriptor`` accepts a string and inflates it into a
single-paragraph rich-text title (matching openpyxl's ``title_maker``).

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any

from .layout import Layout
from .shapes import GraphicalProperties
from .text import (
    CharacterProperties,
    Paragraph,
    ParagraphProperties,
    RegularTextRun,
    RichText,
    Text,
)


class Title:
    """A chart title. ``tx`` holds the rich body; ``layout`` positions it.

    Attributes
    ----------
    tx : :class:`Text` | None
        Rich body (or strRef-bound) text container.
    layout : :class:`Layout` | None
        Manual placement override.
    overlay : bool | None
        Whether the title overlays the plot area.
    spPr : :class:`GraphicalProperties` | None
        Shape properties for the title's container.
    txPr : :class:`RichText` | None
        Default text formatting if ``tx.rich`` is empty.
    """

    __slots__ = ("tx", "layout", "overlay", "spPr", "txPr")

    def __init__(
        self,
        tx: Text | None = None,
        layout: Layout | None = None,
        overlay: bool | None = None,
        spPr: GraphicalProperties | None = None,
        txPr: RichText | None = None,
    ) -> None:
        self.tx = tx if tx is not None else Text()
        self.layout = layout
        self.overlay = overlay
        self.spPr = spPr
        self.txPr = txPr

    # openpyxl aliases
    @property
    def text(self) -> Text | None:
        return self.tx

    @text.setter
    def text(self, value: Text | None) -> None:
        self.tx = value if value is not None else Text()

    @property
    def body(self) -> RichText | None:
        return self.txPr

    @body.setter
    def body(self, value: RichText | None) -> None:
        self.txPr = value

    @property
    def graphicalProperties(self) -> GraphicalProperties | None:
        return self.spPr

    @graphicalProperties.setter
    def graphicalProperties(self, value: GraphicalProperties | None) -> None:
        self.spPr = value

    def to_dict(self) -> dict[str, Any]:
        """Emit the §10.3 shape: ``{text, runs, overlay, layout}``.

        The rich-text body (``self.tx.rich``) is flattened into a list
        of runs ``[{text, font: {name, size, bold, italic, color}}, ...]``.
        If every run carries the same plain text and no formatting,
        ``text`` is also surfaced as a fast-path string.
        """
        runs: list[dict[str, Any]] = []
        text_parts: list[str] = []
        all_plain = True

        if self.tx is not None and self.tx.rich is not None:
            for para in self.tx.rich.p:
                for r in para.r:
                    text_parts.append(r.t or "")
                    font: dict[str, Any] = {}
                    if r.rPr is not None:
                        if r.rPr.latin is not None:
                            font["name"] = r.rPr.latin
                        if r.rPr.sz is not None:
                            # CharacterProperties stores size as 1/100 pt;
                            # surface as integer points.
                            try:
                                font["size"] = int(r.rPr.sz) // 100
                            except (TypeError, ValueError):
                                font["size"] = r.rPr.sz
                        if r.rPr.b is not None:
                            font["bold"] = r.rPr.b
                        if r.rPr.i is not None:
                            font["italic"] = r.rPr.i
                        if r.rPr.solidFill is not None:
                            font["color"] = r.rPr.solidFill
                        if font:
                            all_plain = False
                    runs.append({"text": r.t or "", "font": font or None})

        text_str: str | None = None
        if all_plain and text_parts:
            text_str = "\n".join(text_parts)

        d: dict[str, Any] = {
            "text": text_str,
            "runs": runs if runs else None,
            "overlay": self.overlay,
            "layout": self.layout.to_dict() if self.layout is not None else None,
        }
        # Drop wholly-empty title (no text, no runs, no overlay, no layout)
        if (
            d["text"] is None
            and d["runs"] is None
            and d["overlay"] is None
            and d["layout"] is None
        ):
            return {}
        return d


def title_maker(text: str) -> Title:
    """Inflate a bare string into a single-run rich-text :class:`Title`.

    Mirrors openpyxl's ``title_maker`` — splits on ``"\\n"`` so multi-line
    titles emit one ``<a:p>`` per line.
    """
    title = Title()
    paraprops = ParagraphProperties(defRPr=CharacterProperties())
    paras = [
        Paragraph(pPr=paraprops, r=[RegularTextRun(t=line)])
        for line in text.split("\n")
    ]
    if title.tx is None or title.tx.rich is None:
        title.tx = Text(rich=RichText())
    title.tx.rich.paragraphs = paras
    return title


class TitleDescriptor:
    """Descriptor that auto-inflates string assignments into :class:`Title`.

    ``chart.title = "Sales"`` becomes ``Title(...)`` with a single-run
    rich body — matching openpyxl's behaviour exactly.
    """

    def __set_name__(self, owner: type, name: str) -> None:
        self._attr = "_" + name

    def __get__(self, instance: Any, owner: type | None = None) -> Title | None:
        if instance is None:
            return self  # type: ignore[return-value]
        return getattr(instance, self._attr, None)

    def __set__(self, instance: Any, value: Any) -> None:
        if value is None:
            setattr(instance, self._attr, None)
            return
        if isinstance(value, str):
            value = title_maker(value)
        if not isinstance(value, Title):
            raise TypeError(f"Title must be str or Title, got {type(value).__name__}")
        setattr(instance, self._attr, value)


__all__ = ["Title", "TitleDescriptor", "title_maker"]
