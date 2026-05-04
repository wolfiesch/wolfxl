"""``openpyxl.pivot.table`` — re-export shim.

Wolfxl's pivot-table classes live in :mod:`wolfxl.pivot._table`; the
package init re-exports them.  This module surfaces the same surface
under the openpyxl-shaped explicit-module path so
``from openpyxl.pivot.table import PivotTable`` ports mechanically.
"""

from __future__ import annotations

from wolfxl.pivot._table import (
    ColumnField,
    DataField,
    DataFunction,
    Location,
    PageField,
    PivotField,
    PivotItem,
    PivotSource,
    PivotTable,
    PivotTableStyleInfo,
    RowField,
)


# openpyxl exposes ``TableDefinition`` as the underlying name; alias for
# import-compat parity.
TableDefinition = PivotTable


__all__ = [
    "ColumnField",
    "DataField",
    "DataFunction",
    "Location",
    "PageField",
    "PivotField",
    "PivotItem",
    "PivotSource",
    "PivotTable",
    "PivotTableStyleInfo",
    "RowField",
    "TableDefinition",
]
