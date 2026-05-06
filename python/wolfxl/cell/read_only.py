"""Read-only cell helpers compatible with ``openpyxl.cell.read_only``."""

from __future__ import annotations

from typing import Any

from wolfxl._cell import Cell
from wolfxl.utils import get_column_letter


class ReadOnlyCell:
    """Immutable cell proxy used by openpyxl's read-only worksheet surface."""

    __slots__ = ("parent", "row", "column", "_value", "data_type", "_style_id")

    def __init__(
        self,
        sheet: Any,
        row: int,
        column: int,
        value: Any,
        data_type: str = "n",
        style_id: int = 0,
    ) -> None:
        self.parent = sheet
        self.row = row
        self.column = column
        self._value = value
        self.data_type = data_type
        self._style_id = style_id

    def __repr__(self) -> str:
        title = getattr(self.parent, "title", "")
        return f"<ReadOnlyCell {title!r}.{self.coordinate}>"

    @property
    def coordinate(self) -> str:
        return f"{get_column_letter(self.column)}{self.row}"

    @property
    def column_letter(self) -> str:
        return get_column_letter(self.column)

    @property
    def style_array(self) -> Any:
        return self.parent.parent._cell_styles[self._style_id]

    @property
    def has_style(self) -> bool:
        return self._style_id != 0

    @property
    def number_format(self) -> Any:
        styles = getattr(self.parent.parent, "_cell_styles", None)
        if styles is None:
            return None
        try:
            style = styles[self._style_id]
            return getattr(style, "numFmtId", 0)
        except Exception:
            return None

    @property
    def font(self) -> Any:
        return getattr(Cell, "font", None).__get__(self) if hasattr(Cell, "font") else None

    @property
    def fill(self) -> Any:
        return getattr(Cell, "fill", None).__get__(self) if hasattr(Cell, "fill") else None

    @property
    def border(self) -> Any:
        return getattr(Cell, "border", None).__get__(self) if hasattr(Cell, "border") else None

    @property
    def alignment(self) -> Any:
        return getattr(Cell, "alignment", None).__get__(self) if hasattr(Cell, "alignment") else None

    @property
    def protection(self) -> Any:
        return getattr(Cell, "protection", None).__get__(self) if hasattr(Cell, "protection") else None

    @property
    def is_date(self) -> bool:
        return bool(getattr(Cell, "is_date", False))

    @property
    def internal_value(self) -> Any:
        return self._value

    @property
    def value(self) -> Any:
        return self._value

    @value.setter
    def value(self, value: Any) -> None:
        if self._value is not None:
            raise AttributeError("Cell is read only")
        self._value = value


class EmptyCell:
    """Singleton placeholder for missing read-only cells."""

    __slots__ = ()

    value = None
    is_date = False
    font = None
    border = None
    fill = None
    number_format = None
    alignment = None
    data_type = "n"

    def __repr__(self) -> str:
        return "<EmptyCell>"


EMPTY_CELL = EmptyCell()

__all__ = ["EMPTY_CELL", "EmptyCell", "ReadOnlyCell"]
