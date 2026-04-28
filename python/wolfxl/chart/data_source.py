"""Cell-range and literal data sources for chart series.

Mirrors :mod:`openpyxl.chart.data_source`. Each class is a thin attribute
carrier with a ``to_dict()`` method matching the camelCase XML names so
the Rust emitter can serialise it.

* ``NumRef`` / ``StrRef`` — references to a cell range, with optional
  cached values (``numCache`` / ``strCache``).
* ``NumLit`` (a.k.a. ``NumData``) / ``StrLit`` (``StrData``) — embedded
  literal values, used for charts not backed by a sheet.
* ``NumDataSource`` / ``AxDataSource`` — wrappers chosen by
  ``Series.val`` / ``Series.cat`` / etc.

Sprint Μ Pod-β (RFC-046).
"""

from __future__ import annotations

from typing import Any

from .reference import Reference


class NumFmt:
    """`<c:numFmt>` — number format with optional source-link flag."""

    __slots__ = ("formatCode", "sourceLinked")

    def __init__(self, formatCode: str | None = None, sourceLinked: bool = False) -> None:
        self.formatCode = formatCode
        self.sourceLinked = sourceLinked

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"sourceLinked": self.sourceLinked}
        if self.formatCode is not None:
            d["formatCode"] = self.formatCode
        return d


class NumVal:
    """A single numeric cache point (`<c:pt idx="..."><c:v>..</c:v></c:pt>`)."""

    __slots__ = ("idx", "formatCode", "v")

    def __init__(
        self,
        idx: int | None = None,
        formatCode: str | None = None,
        v: Any | None = None,
    ) -> None:
        self.idx = idx
        self.formatCode = formatCode
        self.v = v

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.idx is not None:
            d["idx"] = self.idx
        if self.formatCode is not None:
            d["formatCode"] = self.formatCode
        if self.v is not None:
            d["v"] = self.v
        return d


class NumData:
    """`<c:numCache>` / `<c:numLit>` — sequence of numeric points."""

    __slots__ = ("formatCode", "ptCount", "pt")

    def __init__(
        self,
        formatCode: str | None = None,
        ptCount: int | None = None,
        pt: list[NumVal] | tuple[NumVal, ...] = (),
    ) -> None:
        self.formatCode = formatCode
        self.ptCount = ptCount
        self.pt = list(pt)

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.formatCode is not None:
            d["formatCode"] = self.formatCode
        if self.ptCount is not None:
            d["ptCount"] = self.ptCount
        if self.pt:
            d["pt"] = [p.to_dict() for p in self.pt]
        return d


# alias matching the openpyxl name; same type, separate identity for callers
NumLit = NumData


class StrVal:
    """A single string cache point."""

    __slots__ = ("idx", "v")

    def __init__(self, idx: int = 0, v: str | None = None) -> None:
        self.idx = idx
        self.v = v

    def to_dict(self) -> dict[str, Any]:
        return {"idx": self.idx, "v": self.v}


class StrData:
    """`<c:strCache>` / `<c:strLit>`."""

    __slots__ = ("ptCount", "pt")

    def __init__(
        self,
        ptCount: int | None = None,
        pt: list[StrVal] | tuple[StrVal, ...] = (),
    ) -> None:
        self.ptCount = ptCount
        self.pt = list(pt)

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.ptCount is not None:
            d["ptCount"] = self.ptCount
        if self.pt:
            d["pt"] = [p.to_dict() for p in self.pt]
        return d


StrLit = StrData


def _ref_to_str(f: Any) -> str | None:
    """Coerce a :class:`Reference`, raw string, or None into a formula string."""
    if f is None:
        return None
    if isinstance(f, Reference):
        return str(f)
    return str(f)


class NumRef:
    """`<c:numRef>` — a numeric data range with optional cache."""

    __slots__ = ("f", "numCache")

    def __init__(self, f: Any | None = None, numCache: NumData | None = None) -> None:
        self.f = _ref_to_str(f)
        self.numCache = numCache

    @property
    def ref(self) -> str | None:
        return self.f

    @ref.setter
    def ref(self, value: Any) -> None:
        self.f = _ref_to_str(value)

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"f": self.f}
        if self.numCache is not None:
            d["numCache"] = self.numCache.to_dict()
        return d


class StrRef:
    """`<c:strRef>` — a string data range with optional cache."""

    __slots__ = ("f", "strCache")

    def __init__(self, f: Any | None = None, strCache: StrData | None = None) -> None:
        self.f = _ref_to_str(f)
        self.strCache = strCache

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"f": self.f}
        if self.strCache is not None:
            d["strCache"] = self.strCache.to_dict()
        return d


class NumDataSource:
    """`<c:val>` / `<c:bubbleSize>` / `<c:yVal>` — wraps numRef or numLit."""

    __slots__ = ("numRef", "numLit")

    def __init__(self, numRef: NumRef | None = None, numLit: NumData | None = None) -> None:
        self.numRef = numRef
        self.numLit = numLit

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.numRef is not None:
            d["numRef"] = self.numRef.to_dict()
        if self.numLit is not None:
            d["numLit"] = self.numLit.to_dict()
        return d


class AxDataSource:
    """`<c:cat>` / `<c:xVal>` — wraps any of {numRef, numLit, strRef, strLit, multiLvlStrRef}."""

    __slots__ = ("numRef", "numLit", "strRef", "strLit", "multiLvlStrRef")

    def __init__(
        self,
        numRef: NumRef | None = None,
        numLit: NumData | None = None,
        strRef: StrRef | None = None,
        strLit: StrData | None = None,
        multiLvlStrRef: Any | None = None,
    ) -> None:
        if not any([numRef, numLit, strRef, strLit, multiLvlStrRef]):
            raise TypeError("AxDataSource requires at least one source")
        self.numRef = numRef
        self.numLit = numLit
        self.strRef = strRef
        self.strLit = strLit
        self.multiLvlStrRef = multiLvlStrRef

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.numRef is not None:
            d["numRef"] = self.numRef.to_dict()
        if self.numLit is not None:
            d["numLit"] = self.numLit.to_dict()
        if self.strRef is not None:
            d["strRef"] = self.strRef.to_dict()
        if self.strLit is not None:
            d["strLit"] = self.strLit.to_dict()
        if self.multiLvlStrRef is not None:
            d["multiLvlStrRef"] = (
                self.multiLvlStrRef.to_dict()
                if hasattr(self.multiLvlStrRef, "to_dict")
                else self.multiLvlStrRef
            )
        return d


__all__ = [
    "AxDataSource",
    "NumData",
    "NumDataSource",
    "NumFmt",
    "NumLit",
    "NumRef",
    "NumVal",
    "StrData",
    "StrLit",
    "StrRef",
    "StrVal",
]
