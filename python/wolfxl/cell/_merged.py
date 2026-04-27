"""MergedCell — placeholder for non-anchor cells in a merged range.

openpyxl exposes a ``MergedCell`` type whose ``.value`` is always
``None`` and whose value setter raises :class:`AttributeError`.
The anchor (top-left) cell of the merged range remains a regular
:class:`Cell`; every other coordinate inside the range becomes a
``MergedCell``.

Wolfxl manages merges differently (the rows/columns continue to
hold real :class:`Cell` proxies and the merge is tracked via the
worksheet-level ``merged_cells`` collection), so this class is
a thin compatibility shim — user code that constructs or
``isinstance``-checks ``MergedCell`` to detect non-anchor
positions can migrate unchanged.

Reference: ``openpyxl.cell.cell.MergedCell`` (openpyxl 3.1.x).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from wolfxl._utils import rowcol_to_a1

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


class MergedCell:
    """Placeholder for cells inside a merged range that aren't
    the top-left anchor.

    ``value`` is always ``None``; assignment raises
    :class:`AttributeError` to match openpyxl's contract.
    """

    __slots__ = ("_parent", "_row", "_col")

    def __init__(self, parent: Worksheet | None, row: int, column: int) -> None:
        self._parent = parent
        self._row = row
        self._col = column

    @property
    def parent(self) -> Worksheet | None:
        return self._parent

    @property
    def row(self) -> int:
        return self._row

    @property
    def column(self) -> int:
        return self._col

    @property
    def coordinate(self) -> str:
        return rowcol_to_a1(self._row, self._col)

    @property
    def value(self) -> Any:
        return None

    @value.setter
    def value(self, _: Any) -> None:
        raise AttributeError("Cell is part of a merged range")

    def __repr__(self) -> str:
        return f"<MergedCell {self.coordinate}>"


__all__ = ["MergedCell"]
