"""``openpyxl.worksheet.properties`` — sheet-level property containers.

Pod 2 (RFC-060 §2.1).  Wolfxl preserves these on round-trip but does
not yet expose Python-side construction; the names land as stubs so
import statements port mechanically.
"""

from __future__ import annotations

from wolfxl._compat import _make_stub

WorksheetProperties = _make_stub(
    "WorksheetProperties",
    "Wolfxl preserves worksheet properties on round-trip; construction "
    "lands in a future release.",
)
PageSetupProperties = _make_stub(
    "PageSetupProperties",
    "Wolfxl preserves page-setup properties on round-trip; construction "
    "lands in a future release.",
)
Outline = _make_stub(
    "Outline",
    "Wolfxl tracks row/column outline levels via "
    "``ws.row_dimensions[r].outline_level`` accessors; the Outline "
    "value type is not yet exposed.",
)


__all__ = ["Outline", "PageSetupProperties", "WorksheetProperties"]
