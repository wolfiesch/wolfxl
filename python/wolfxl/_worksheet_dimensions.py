"""Worksheet row and column dimension proxy objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from wolfxl._utils import column_index

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


class RowDimensionProxy:
    """Dict-like proxy for row metadata.

    Examples:
        Set row height using the openpyxl-shaped API:

        >>> ws.row_dimensions[1].height = 30
    """

    __slots__ = ("_ws",)

    def __init__(self, ws: Worksheet) -> None:
        self._ws = ws

    def __getitem__(self, row: int) -> RowDimension:
        return RowDimension(self._ws, row)

    def get(self, row: int, default: Any = None) -> RowDimension | Any:
        if not isinstance(row, int):
            return default
        dimension = RowDimension(self._ws, row)
        if dimension.height is not None or dimension.hidden or dimension.outline_level:
            return dimension
        return default


class RowDimension:
    """Single row dimension with openpyxl-shaped metadata properties."""

    __slots__ = ("_ws", "_row")

    def __init__(self, ws: Worksheet, row: int) -> None:
        self._ws = ws
        self._row = row

    @property
    def height(self) -> float | None:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            return wb._rust_reader.read_row_height(self._ws._title, self._row)  # noqa: SLF001
        return self._ws._row_heights.get(self._row)  # noqa: SLF001

    @height.setter
    def height(self, value: float | None) -> None:
        self._ws._row_heights[self._row] = value  # noqa: SLF001

    @property
    def hidden(self) -> bool:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            return self._row in self._ws.sheet_visibility()["hidden_rows"]
        return False

    @property
    def outlineLevel(self) -> int:  # noqa: N802 - openpyxl public API
        return self.outline_level

    @property
    def outline_level(self) -> int:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            return int(self._ws.sheet_visibility()["row_outline_levels"].get(self._row, 0))
        return 0


class ColumnDimensionProxy:
    """Dict-like proxy for column metadata.

    Examples:
        Set column width using the openpyxl-shaped API:

        >>> ws.column_dimensions["A"].width = 15
    """

    __slots__ = ("_ws",)

    def __init__(self, ws: Worksheet) -> None:
        self._ws = ws

    def __getitem__(self, col_letter: str) -> ColumnDimension:
        return ColumnDimension(self._ws, col_letter.upper())

    def get(self, col_letter: str, default: Any = None) -> ColumnDimension | Any:
        if not isinstance(col_letter, str):
            return default
        dimension = ColumnDimension(self._ws, col_letter.upper())
        if dimension.width is not None or dimension.hidden or dimension.outline_level:
            return dimension
        return default


class ColumnDimension:
    """Single column dimension with openpyxl-shaped metadata properties."""

    __slots__ = ("_ws", "_col_letter")

    def __init__(self, ws: Worksheet, col_letter: str) -> None:
        self._ws = ws
        self._col_letter = col_letter

    @property
    def width(self) -> float | None:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            return wb._rust_reader.read_column_width(self._ws._title, self._col_letter)  # noqa: SLF001
        return self._ws._col_widths.get(self._col_letter)  # noqa: SLF001

    @width.setter
    def width(self, value: float | None) -> None:
        self._ws._col_widths[self._col_letter] = value  # noqa: SLF001

    @property
    def hidden(self) -> bool:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            return column_index(self._col_letter) in self._ws.sheet_visibility()["hidden_columns"]
        return False

    @property
    def outlineLevel(self) -> int:  # noqa: N802 - openpyxl public API
        return self.outline_level

    @property
    def outline_level(self) -> int:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            col = column_index(self._col_letter)
            return int(self._ws.sheet_visibility()["column_outline_levels"].get(col, 0))
        return 0
