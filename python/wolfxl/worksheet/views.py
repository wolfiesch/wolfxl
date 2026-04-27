"""Sheet view classes (RFC-055 §2.5).

Backs ``ws.sheet_view``. Provides ``Pane``, ``Selection``, ``SheetView``,
and ``SheetViewList`` per OOXML CT_SheetView.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


_VALID_PANE_NAMES = ("bottomLeft", "bottomRight", "topLeft", "topRight")
_VALID_PANE_STATES = ("frozen", "split", "frozenSplit")
_VALID_VIEWS = ("normal", "pageBreakPreview", "pageLayout")


@dataclass
class Pane:
    """A pane within a sheet view (CT_Pane)."""

    xSplit: float = 0.0  # noqa: N815
    ySplit: float = 0.0  # noqa: N815
    topLeftCell: str = "A1"  # noqa: N815
    activePane: str = "topLeft"  # noqa: N815
    state: str = "frozen"

    @property
    def x_split(self) -> float:
        return self.xSplit

    @x_split.setter
    def x_split(self, value: float) -> None:
        self.xSplit = float(value)

    @property
    def y_split(self) -> float:
        return self.ySplit

    @y_split.setter
    def y_split(self, value: float) -> None:
        self.ySplit = float(value)

    @property
    def top_left_cell(self) -> str:
        return self.topLeftCell

    @top_left_cell.setter
    def top_left_cell(self, value: str) -> None:
        self.topLeftCell = value

    @property
    def active_pane(self) -> str:
        return self.activePane

    @active_pane.setter
    def active_pane(self, value: str) -> None:
        if value not in _VALID_PANE_NAMES:
            raise ValueError(
                f"active_pane must be one of {_VALID_PANE_NAMES}, got {value!r}"
            )
        self.activePane = value

    def __post_init__(self) -> None:
        if self.activePane not in _VALID_PANE_NAMES:
            raise ValueError(
                f"activePane must be one of {_VALID_PANE_NAMES}, got {self.activePane!r}"
            )
        if self.state not in _VALID_PANE_STATES:
            raise ValueError(
                f"state must be one of {_VALID_PANE_STATES}, got {self.state!r}"
            )

    def to_rust_dict(self) -> dict[str, Any]:
        return {
            "x_split": float(self.xSplit),
            "y_split": float(self.ySplit),
            "top_left_cell": self.topLeftCell,
            "active_pane": self.activePane,
            "state": self.state,
        }


@dataclass
class Selection:
    """Cell selection within a pane (CT_Selection)."""

    activeCell: str = "A1"  # noqa: N815
    sqref: str = "A1"
    pane: str | None = None
    activeCellId: int | None = None  # noqa: N815

    @property
    def active_cell(self) -> str:
        return self.activeCell

    @active_cell.setter
    def active_cell(self, value: str) -> None:
        self.activeCell = value

    def __post_init__(self) -> None:
        if self.pane is not None and self.pane not in _VALID_PANE_NAMES:
            raise ValueError(
                f"pane must be one of {_VALID_PANE_NAMES} or None, got {self.pane!r}"
            )

    def to_rust_dict(self) -> dict[str, Any]:
        return {
            "active_cell": self.activeCell,
            "sqref": self.sqref,
            "pane": self.pane,
        }


@dataclass
class SheetView:
    """Single sheet view (CT_SheetView)."""

    zoomScale: int = 100  # noqa: N815
    zoomScaleNormal: int = 100  # noqa: N815
    view: str = "normal"
    showGridLines: bool = True  # noqa: N815
    showRowColHeaders: bool = True  # noqa: N815
    showOutlineSymbols: bool = True  # noqa: N815
    showZeros: bool = True  # noqa: N815
    rightToLeft: bool = False  # noqa: N815
    tabSelected: bool = False  # noqa: N815
    topLeftCell: str | None = None  # noqa: N815
    workbookViewId: int = 0  # noqa: N815
    pane: Pane | None = None
    selection: list[Selection] = field(default_factory=list)

    # snake_case aliases
    @property
    def zoom_scale(self) -> int:
        return self.zoomScale

    @zoom_scale.setter
    def zoom_scale(self, value: int) -> None:
        if not (10 <= int(value) <= 400):
            raise ValueError(f"zoom_scale must be between 10 and 400, got {value}")
        self.zoomScale = int(value)

    @property
    def zoom_scale_normal(self) -> int:
        return self.zoomScaleNormal

    @zoom_scale_normal.setter
    def zoom_scale_normal(self, value: int) -> None:
        self.zoomScaleNormal = int(value)

    @property
    def show_grid_lines(self) -> bool:
        return self.showGridLines

    @show_grid_lines.setter
    def show_grid_lines(self, value: bool) -> None:
        self.showGridLines = bool(value)

    @property
    def show_row_col_headers(self) -> bool:
        return self.showRowColHeaders

    @show_row_col_headers.setter
    def show_row_col_headers(self, value: bool) -> None:
        self.showRowColHeaders = bool(value)

    @property
    def show_outline_symbols(self) -> bool:
        return self.showOutlineSymbols

    @show_outline_symbols.setter
    def show_outline_symbols(self, value: bool) -> None:
        self.showOutlineSymbols = bool(value)

    @property
    def show_zeros(self) -> bool:
        return self.showZeros

    @show_zeros.setter
    def show_zeros(self, value: bool) -> None:
        self.showZeros = bool(value)

    @property
    def right_to_left(self) -> bool:
        return self.rightToLeft

    @right_to_left.setter
    def right_to_left(self, value: bool) -> None:
        self.rightToLeft = bool(value)

    @property
    def tab_selected(self) -> bool:
        return self.tabSelected

    @tab_selected.setter
    def tab_selected(self, value: bool) -> None:
        self.tabSelected = bool(value)

    @property
    def top_left_cell(self) -> str | None:
        return self.topLeftCell

    @top_left_cell.setter
    def top_left_cell(self, value: str | None) -> None:
        self.topLeftCell = value

    def __post_init__(self) -> None:
        if self.view not in _VALID_VIEWS:
            raise ValueError(
                f"view must be one of {_VALID_VIEWS}, got {self.view!r}"
            )
        if not (10 <= int(self.zoomScale) <= 400):
            raise ValueError(f"zoomScale must be between 10 and 400, got {self.zoomScale}")

    def to_rust_dict(self) -> dict[str, Any]:
        return {
            "zoom_scale": int(self.zoomScale),
            "zoom_scale_normal": int(self.zoomScaleNormal),
            "view": self.view,
            "show_grid_lines": bool(self.showGridLines),
            "show_row_col_headers": bool(self.showRowColHeaders),
            "show_outline_symbols": bool(self.showOutlineSymbols),
            "show_zeros": bool(self.showZeros),
            "right_to_left": bool(self.rightToLeft),
            "tab_selected": bool(self.tabSelected),
            "top_left_cell": self.topLeftCell,
            "pane": self.pane.to_rust_dict() if self.pane is not None else None,
            "selection": [s.to_rust_dict() for s in self.selection],
        }

    def is_default(self) -> bool:
        return (
            self.zoomScale == 100
            and self.zoomScaleNormal == 100
            and self.view == "normal"
            and self.showGridLines
            and self.showRowColHeaders
            and self.showOutlineSymbols
            and self.showZeros
            and not self.rightToLeft
            and not self.tabSelected
            and self.topLeftCell is None
            and self.pane is None
            and not self.selection
        )


class SheetViewList:
    """Container for a worksheet's sheet views.

    openpyxl exposes this as ``ws.sheet_view`` (singular property
    returning the first / only view) plus
    ``ws.views.sheetView`` (a list). Wolfxl mirrors both: the
    ``views`` accessor returns a SheetViewList, while ``sheet_view``
    returns the first SheetView (creating one if absent).
    """

    __slots__ = ("sheetView",)

    def __init__(self, views: list[SheetView] | None = None) -> None:
        # openpyxl naming: list attribute is ``sheetView`` (singular).
        self.sheetView = list(views or [SheetView()])  # noqa: N815

    def __iter__(self):
        return iter(self.sheetView)

    def __len__(self) -> int:
        return len(self.sheetView)

    def __getitem__(self, idx: int) -> SheetView:
        return self.sheetView[idx]


__all__ = [
    "Pane",
    "Selection",
    "SheetView",
    "SheetViewList",
]
