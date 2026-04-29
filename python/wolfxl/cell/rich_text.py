"""Rich-text cell value support.

Provides three classes that mirror openpyxl's
``openpyxl.cell.rich_text`` surface:

* :class:`InlineFont` — a subset of font properties that may decorate a
  rich-text *run* (bold, italic, underline, strike, font name, size,
  color).  Matches openpyxl's ``InlineFont`` constructor keyword
  contract for the fields wolfxl actually round-trips.
* :class:`TextBlock` — a single styled run: ``font`` + ``text``.
* :class:`CellRichText` — an iterable container of ``str`` and
  ``TextBlock`` items, modeling a cell value that carries multiple
  styled runs.

These shims intentionally match openpyxl's iteration / equality /
``__str__`` semantics so user code that walks ``cell.value`` items via
``isinstance(item, (str, TextBlock))`` Just Works regardless of which
library produced the value.

Sprint Ι Pod-α (RFC pending) — closes the Phase 3 rich-text-reads
gap and the implicit T3 rich-text-write deferral.
"""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from typing import Optional, Union


@dataclass(eq=True)
class InlineFont:
    """Font properties for a single rich-text run.

    Field names mirror openpyxl's ``InlineFont`` keyword arguments
    (single-letter for the boolean attributes — ``b`` for bold, ``i``
    for italic, ``u`` for underline style, etc.).  All fields default
    to ``None`` so an empty ``InlineFont()`` round-trips as a run with
    no explicit ``<rPr>`` block.

    Attributes:
        rFont: Font family name.
        charset: Font character set id.
        family: Font family id.
        b: Bold flag.
        i: Italic flag.
        strike: Strikethrough flag.
        color: ARGB hex string or theme/indexed color descriptor.
        sz: Font size in points.
        u: Underline style.
        vertAlign: Vertical alignment.
        scheme: Font scheme.
    """

    rFont: Optional[str] = None
    """Font family name (openpyxl alias for ``Font.name``)."""

    charset: Optional[int] = None
    family: Optional[int] = None
    b: Optional[bool] = None
    """Bold."""
    i: Optional[bool] = None
    """Italic."""
    strike: Optional[bool] = None
    outline: Optional[bool] = None
    shadow: Optional[bool] = None
    condense: Optional[bool] = None
    extend: Optional[bool] = None
    color: Optional[str] = None
    """ARGB hex string (e.g. ``"FFFF0000"``) or theme/indexed color
    descriptor.  Stored verbatim from XML."""
    sz: Optional[float] = None
    """Font size in points."""
    u: Optional[str] = None
    """Underline style (e.g. ``"single"``, ``"double"``).  ``True``
    coerces to ``"single"`` for openpyxl parity."""
    vertAlign: Optional[str] = None
    scheme: Optional[str] = None

    def __post_init__(self) -> None:
        # openpyxl coerces the underline boolean shorthand to the
        # canonical "single" style.  Mirror that so user code that does
        # ``InlineFont(u=True)`` ends up with ``u="single"``.
        if self.u is True:
            self.u = "single"
        elif self.u is False:
            self.u = None


@dataclass
class TextBlock:
    """A styled run inside a :class:`CellRichText`."""

    font: InlineFont
    text: str

    def __init__(self, font: InlineFont, text: str) -> None:
        """Create a styled rich-text run.

        Args:
            font: Inline font applied to this run.
            text: Plain text for this run.
        """
        # openpyxl positions ``font`` first to match the ``<r><rPr>...</rPr><t>...</t></r>``
        # XML reading order.  Mirror that signature.
        self.font = font
        self.text = text

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, TextBlock):
            return NotImplemented
        return self.font == other.font and self.text == other.text

    def __hash__(self) -> int:  # pragma: no cover - dataclasses default
        return hash((self.text,))

    def __str__(self) -> str:
        # openpyxl's ``str(TextBlock)`` returns the plain text — this
        # makes ``"".join(map(str, cell_rich_text))`` produce the
        # flattened representation.
        return self.text

    def __repr__(self) -> str:
        return f"TextBlock text={self.text}, font={self.font!r}"


class CellRichText(list):
    """Sequence of ``str`` and :class:`TextBlock` runs.

    Subclassing ``list`` so existing user code that iterates,
    indexes, slices, or appends to a rich-text value Just Works
    (matches openpyxl's design — ``CellRichText`` is also a list
    subclass there).
    """

    def __init__(self, items: Optional[Iterable[Union[str, TextBlock]]] = None) -> None:
        """Create a rich-text run list.

        Args:
            items: Optional iterable of ``str`` and ``TextBlock`` runs. A
                single ``str`` or ``TextBlock`` is also accepted.
        """
        super().__init__()
        if items is None:
            return
        # Allow a single string or TextBlock as a convenience.
        if isinstance(items, (str, TextBlock)):
            self.append(items)
            return
        for item in items:
            self.append(item)

    def append(self, value: Union[str, TextBlock]) -> None:  # type: ignore[override]
        """Append one rich-text run.

        Args:
            value: Plain string run or styled ``TextBlock``.

        Raises:
            TypeError: If ``value`` is not a string or ``TextBlock``.
        """
        if not isinstance(value, (str, TextBlock)):
            raise TypeError(
                "CellRichText items must be str or TextBlock, "
                f"got {type(value).__name__}"
            )
        super().append(value)

    def __iadd__(self, other: Iterable[Union[str, TextBlock]]) -> "CellRichText":  # type: ignore[override]
        for item in other:
            self.append(item)
        return self

    def __add__(self, other: Iterable[Union[str, TextBlock]]) -> "CellRichText":  # type: ignore[override]
        out = CellRichText(self)
        out += other
        return out

    def __str__(self) -> str:
        # Flatten to plain text — same semantics as openpyxl's
        # ``CellRichText.__str__``.
        return "".join(str(item) for item in self)

    def __repr__(self) -> str:
        inner = ", ".join(repr(item) for item in self)
        return f"CellRichText([{inner}])"

    def as_list(self) -> list[Union[str, TextBlock]]:
        """Materialize a plain ``list`` copy of the runs.

        Convenience for callers that want to avoid handing the
        underlying mutable list to downstream code.

        Returns:
            A shallow list copy of the current rich-text runs.
        """
        return list(self)


__all__ = ["CellRichText", "InlineFont", "TextBlock"]
