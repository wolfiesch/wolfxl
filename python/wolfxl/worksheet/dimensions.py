"""``openpyxl.worksheet.dimensions`` — re-export shim plus dimension helpers.

Wolfxl's row / column dimension proxies live in :mod:`wolfxl._worksheet`
under underscore-prefixed names (the public interaction is via
``ws.row_dimensions[…]`` / ``ws.column_dimensions[…]``). This module
surfaces the same classes under the openpyxl-shaped names so
``from openpyxl.worksheet.dimensions import RowDimension`` ports
mechanically.

Sprint Π Pod Π-α (RFC-062) replaces the construction stubs for
:class:`DimensionHolder`, :class:`SheetFormatProperties`, and
:class:`SheetDimension` with real dataclasses.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterator

from wolfxl._worksheet import (
    _ColumnDimension as ColumnDimension,
    _RowDimension as RowDimension,
)


# ---------------------------------------------------------------------------
# DimensionHolder — dict-like view over ws.row_dimensions / column_dimensions
# ---------------------------------------------------------------------------

@dataclass
class DimensionHolder:
    """openpyxl-shape mapping wrapper over ``ws.row_dimensions`` /
    ``ws.column_dimensions``.

    openpyxl's ``DimensionHolder`` doubles as a dict-like view AND as
    the type its ``Worksheet.row_dimensions`` / ``column_dimensions``
    return. Wolfxl already exposes those via dedicated proxies, so
    this class is a thin wrapper that flows ``__getitem__`` /
    ``__setitem__`` / ``__iter__`` / ``__len__`` calls to the
    matching worksheet proxy.

    Attributes:
        worksheet: Backing :class:`~wolfxl._worksheet.Worksheet`.
        default_factory: Optional callable that produces a default
            dimension instance when a key is missing. Compatibility
            shim; wolfxl's proxies already auto-create entries on
            first access, so this slot is informational unless the
            caller iterates the holder explicitly.
        max_outline: openpyxl-shape outline-level cache. Populated by
            wolfxl's own outline-level computation when requested.
    """

    worksheet: Any
    default_factory: Any = None
    max_outline: int = 0

    # ---- mapping protocol ------------------------------------------------

    def _proxy(self) -> Any:
        # Default routing: the row-dimensions proxy. Sub-classing or
        # passing a different proxy via ``worksheet`` lets the user
        # bind a column-dimension wrapper if needed.
        return self.worksheet.row_dimensions

    def __getitem__(self, key: Any) -> Any:
        return self._proxy()[key]

    def __setitem__(self, key: Any, value: Any) -> None:
        self._proxy()[key] = value

    def __iter__(self) -> Iterator[Any]:
        proxy = self._proxy()
        # Wolfxl's row/column-dimension proxies are sparse maps with
        # no explicit backing-store iterator (dimensions are
        # auto-materialised on access), so we can't enumerate them
        # the way openpyxl can. Return an empty iterator when the
        # proxy doesn't natively support iteration.
        try:
            it = proxy.__iter__()  # type: ignore[attr-defined]
        except (AttributeError, TypeError):
            return iter(())
        return it

    def __len__(self) -> int:
        proxy = self._proxy()
        try:
            return proxy.__len__()  # type: ignore[attr-defined]
        except (AttributeError, TypeError):
            return 0

    def __contains__(self, key: Any) -> bool:
        proxy = self._proxy()
        try:
            return proxy.__contains__(key)  # type: ignore[attr-defined]
        except (AttributeError, TypeError):
            return False


# ---------------------------------------------------------------------------
# SheetFormatProperties — <sheetFormatPr> defaults
# ---------------------------------------------------------------------------

@dataclass
class SheetFormatProperties:
    """``<sheetFormatPr>`` defaults (CT_SheetFormatPr §18.3.1.81).

    Attributes:
        baseColWidth: Default column width in characters when no
            explicit width is set on a column.
        defaultColWidth: Optional override for the per-cell-default
            column width.
        defaultRowHeight: Default row height in points (Excel's
            default is 15.0pt).
        customHeight: When ``True``, Excel writes a per-row
            ``customHeight`` flag for rows whose height equals the
            default — almost never user-set.
        zeroHeight: When ``True``, the sheet defaults to "rows
            collapsed to 0 height" (used when generating
            data-only sheets).
        thickTop, thickBottom: When ``True``, Excel renders the
            corresponding sheet edge with a thick rule.
        outlineLevelRow, outlineLevelCol: max outline-grouping
            depth for rows / columns.
    """

    baseColWidth: int = 8                 # noqa: N815
    defaultColWidth: float | None = None  # noqa: N815
    defaultRowHeight: float = 15.0        # noqa: N815
    customHeight: bool = False            # noqa: N815
    zeroHeight: bool = False              # noqa: N815
    thickTop: bool = False                # noqa: N815
    thickBottom: bool = False             # noqa: N815
    outlineLevelRow: int = 0              # noqa: N815
    outlineLevelCol: int = 0              # noqa: N815

    def is_default(self) -> bool:
        """Return ``True`` iff every attribute is at its construction default."""
        return (
            self.baseColWidth == 8
            and self.defaultColWidth is None
            and self.defaultRowHeight == 15.0
            and not self.customHeight
            and not self.zeroHeight
            and not self.thickTop
            and not self.thickBottom
            and self.outlineLevelRow == 0
            and self.outlineLevelCol == 0
        )

    def to_rust_dict(self) -> dict[str, Any]:
        """Return the §10 dict shape consumed by the Rust patcher / writer."""
        return {
            "base_col_width": self.baseColWidth,
            "default_col_width": self.defaultColWidth,
            "default_row_height": self.defaultRowHeight,
            "custom_height": self.customHeight,
            "zero_height": self.zeroHeight,
            "thick_top": self.thickTop,
            "thick_bottom": self.thickBottom,
            "outline_level_row": self.outlineLevelRow,
            "outline_level_col": self.outlineLevelCol,
        }


# ---------------------------------------------------------------------------
# SheetDimension — <dimension ref="..."/>
# ---------------------------------------------------------------------------

@dataclass
class SheetDimension:
    """``<dimension ref="A1:Z100"/>`` (CT_SheetDimension §18.3.1.35).

    Auto-computed by wolfxl from cell bounds at emit time; users can
    override via construction (rare). Pure metadata — Excel re-computes
    on file open if the stored ref disagrees with the actual cell
    bounding box.
    """

    ref: str = "A1"

    def to_rust_dict(self) -> dict[str, Any]:
        return {"ref": self.ref}


# ---------------------------------------------------------------------------
# Dimension — abstract base advisory only
# ---------------------------------------------------------------------------

class Dimension:
    """openpyxl's abstract base for row / column dimensions.

    Direct construction is unusual; the class exists so user code that
    does ``isinstance(d, Dimension)`` against either a row-dimension
    or column-dimension returns ``True``. Wolfxl's :class:`RowDimension`
    and :class:`ColumnDimension` do *not* derive from this class — the
    ``isinstance`` contract is therefore advisory only.
    """


__all__ = [
    "ColumnDimension",
    "Dimension",
    "DimensionHolder",
    "RowDimension",
    "SheetDimension",
    "SheetFormatProperties",
]
