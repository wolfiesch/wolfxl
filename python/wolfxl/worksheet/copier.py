"""``openpyxl.worksheet.copier`` — :class:`WorksheetCopy` shim.

Wolfxl exposes worksheet copying via :meth:`Workbook.copy_worksheet` and
the :class:`CopyOptions` value type; the openpyxl-shaped
:class:`WorksheetCopy` lives here as a stub for import-compat parity.

Pod 2 (RFC-060).
"""

from __future__ import annotations

from wolfxl._compat import _make_stub

WorksheetCopy = _make_stub(
    "WorksheetCopy",
    "Use ``Workbook.copy_worksheet(source)`` instead of constructing "
    "WorksheetCopy directly.",
)


__all__ = ["WorksheetCopy"]
