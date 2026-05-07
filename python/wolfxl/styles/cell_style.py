"""Cell-style array compatibility helpers."""

from __future__ import annotations

from array import array
from copy import copy
from typing import Iterable


class _ArrayDescriptor:
    def __init__(self, key: int) -> None:
        self.key = key

    def __get__(self, instance: "StyleArray", cls: type["StyleArray"]) -> int:
        return instance[self.key]

    def __set__(self, instance: "StyleArray", value: int) -> None:
        instance[self.key] = value


class StyleArray(array):
    """Compact nine-slot style-id tuple used by openpyxl-style callers."""

    __slots__ = ()
    tagname = "xf"

    fontId = _ArrayDescriptor(0)
    fillId = _ArrayDescriptor(1)
    borderId = _ArrayDescriptor(2)
    numFmtId = _ArrayDescriptor(3)
    protectionId = _ArrayDescriptor(4)
    alignmentId = _ArrayDescriptor(5)
    pivotButton = _ArrayDescriptor(6)
    quotePrefix = _ArrayDescriptor(7)
    xfId = _ArrayDescriptor(8)

    def __new__(cls, args: Iterable[int] = (0, 0, 0, 0, 0, 0, 0, 0, 0)) -> "StyleArray":
        return array.__new__(cls, "i", args)

    def __hash__(self) -> int:
        return hash(tuple(self))

    def __copy__(self) -> "StyleArray":
        return StyleArray(self)

    def __deepcopy__(self, memo: dict[int, object]) -> "StyleArray":
        return copy(self)


__all__ = ["StyleArray"]
