"""Worksheet row and column dimension proxy objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from wolfxl._utils import column_index

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def _default_style_value(name: str) -> Any:
    if name == "font":
        from wolfxl import Font

        return Font()
    if name == "fill":
        from wolfxl import PatternFill

        return PatternFill()
    if name == "border":
        from wolfxl import Border

        return Border()
    if name == "alignment":
        from wolfxl import Alignment

        return Alignment()
    if name == "protection":
        from wolfxl.styles import Protection

        return Protection()
    if name == "number_format":
        return "General"
    return None


class _DimensionStyleMixin:
    @property
    def parent(self) -> Worksheet:
        return self._ws  # type: ignore[attr-defined]

    @property
    def has_style(self) -> bool:
        return False

    @property
    def style_id(self) -> int:
        return 0

    @property
    def number_format(self) -> str:
        return _default_style_value("number_format")

    @property
    def font(self) -> Any:
        return _default_style_value("font")

    @property
    def fill(self) -> Any:
        return _default_style_value("fill")

    @property
    def border(self) -> Any:
        return _default_style_value("border")

    @property
    def alignment(self) -> Any:
        return _default_style_value("alignment")

    @property
    def protection(self) -> Any:
        return _default_style_value("protection")

    @property
    def quotePrefix(self) -> bool:  # noqa: N802
        return False

    @property
    def pivotButton(self) -> bool:  # noqa: N802
        return False


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


class RowDimension(_DimensionStyleMixin):
    """Single row dimension with openpyxl-shaped metadata properties."""

    __slots__ = ("_ws", "_row")

    def __init__(self, ws: Worksheet, row: int) -> None:
        self._ws = ws
        self._row = row

    @property
    def index(self) -> int:
        return self._row

    @property
    def r(self) -> int:
        return self._row

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
    def ht(self) -> float | None:
        return self.height

    @ht.setter
    def ht(self, value: float | None) -> None:
        self.height = value

    @property
    def customHeight(self) -> bool:  # noqa: N802
        return self.height is not None

    @property
    def hidden(self) -> bool:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            return self._row in self._ws.sheet_visibility()["hidden_rows"]
        return False

    @property
    def collapsed(self) -> bool:
        return False

    @property
    def style(self) -> int:
        return 0

    @property
    def s(self) -> int:
        return self.style

    @property
    def customFormat(self) -> bool:  # noqa: N802
        return False

    @property
    def thickTop(self) -> bool:  # noqa: N802
        return False

    @property
    def thickBot(self) -> bool:  # noqa: N802
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


class ColumnDimension(_DimensionStyleMixin):
    """Single column dimension with openpyxl-shaped metadata properties."""

    __slots__ = ("_ws", "_col_letter")

    def __init__(self, ws: Worksheet, col_letter: str) -> None:
        self._ws = ws
        self._col_letter = col_letter

    @property
    def index(self) -> str:
        return self._col_letter

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
    def customWidth(self) -> bool:  # noqa: N802
        return self.width is not None

    @property
    def bestFit(self) -> bool:  # noqa: N802
        return False

    @bestFit.setter
    def bestFit(self, value: bool) -> None:  # noqa: N802, ARG002
        # Stored for openpyxl surface compatibility; writer support is not modeled yet.
        return None

    @property
    def auto_size(self) -> bool:
        return self.bestFit

    @auto_size.setter
    def auto_size(self, value: bool) -> None:
        self.bestFit = value

    @property
    def hidden(self) -> bool:
        wb = self._ws._workbook  # noqa: SLF001
        if wb._rust_reader is not None:  # noqa: SLF001
            return column_index(self._col_letter) in self._ws.sheet_visibility()["hidden_columns"]
        return False

    @property
    def collapsed(self) -> bool:
        return False

    @property
    def style(self) -> int:
        return 0

    @property
    def min(self) -> int:
        return column_index(self._col_letter)

    @property
    def max(self) -> int:
        return column_index(self._col_letter)

    @property
    def range(self) -> str:
        return f"{self._col_letter}:{self._col_letter}"

    def reindex(self) -> None:
        """Openpyxl compatibility hook; single-column proxies are already indexed."""
        return None

    def to_tree(self) -> Any:
        from xml.etree import ElementTree as ET

        node = ET.Element("col")
        node.set("min", str(self.min))
        node.set("max", str(self.max))
        if self.width is not None:
            node.set("width", str(self.width))
            node.set("customWidth", "1")
        return node

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
