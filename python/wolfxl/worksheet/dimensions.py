"""``openpyxl.worksheet.dimensions`` — re-export shim.

Wolfxl's row / column dimension proxies live in :mod:`wolfxl._worksheet`
under underscore-prefixed names (the public interaction is via
``ws.row_dimensions[…]`` / ``ws.column_dimensions[…]``).  This module
surfaces the same classes under the openpyxl-shaped names so
``from openpyxl.worksheet.dimensions import RowDimension`` ports
mechanically.

Pod 2 (RFC-060 §2.1).
"""

from __future__ import annotations

from wolfxl._compat import _make_stub
from wolfxl._worksheet import (
    _ColumnDimension as ColumnDimension,
    _RowDimension as RowDimension,
)


# Containers that openpyxl exposes.  Wolfxl tracks dimensions inline on
# the Worksheet proxy, so these are stubs — instantiating them raises
# with a hint pointing at ``ws.row_dimensions`` / ``ws.column_dimensions``.
DimensionHolder = _make_stub(
    "DimensionHolder",
    "openpyxl's DimensionHolder is wolfxl's ``ws.row_dimensions`` / "
    "``ws.column_dimensions`` proxy — interact through those properties.",
)
SheetFormatProperties = _make_stub(
    "SheetFormatProperties",
    "Wolfxl tracks default row/column dimensions inline on the Worksheet; "
    "construction is not exposed.",
)
SheetDimension = _make_stub(
    "SheetDimension",
    "Wolfxl computes the sheet's used range automatically; "
    "SheetDimension is not constructable.",
)


class Dimension:
    """openpyxl's abstract base for row / column dimensions.

    Direct construction is unusual; the class exists so user code that
    does ``isinstance(d, Dimension)`` against either a row-dimension
    or column-dimension returns ``True``.  Wolfxl's
    :class:`RowDimension` and :class:`ColumnDimension` do *not* derive
    from this class — the ``isinstance`` contract is therefore
    advisory only.
    """


__all__ = [
    "ColumnDimension",
    "Dimension",
    "DimensionHolder",
    "RowDimension",
    "SheetDimension",
    "SheetFormatProperties",
]
