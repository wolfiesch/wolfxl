"""``openpyxl.workbook.child`` — internal ``_WorkbookChild`` mixin.

Pod 2 (RFC-060).  openpyxl uses ``_WorkbookChild`` as a base for
sheet-like objects; wolfxl's :class:`Worksheet` does not derive from
it, so this module surfaces a stub for import-compat parity.
"""

from __future__ import annotations

from wolfxl._compat import _make_stub

_WorkbookChild = _make_stub(
    "_WorkbookChild",
    "openpyxl's _WorkbookChild is an internal mixin; wolfxl's Worksheet "
    "does not derive from it.",
)


__all__ = ["_WorkbookChild"]
