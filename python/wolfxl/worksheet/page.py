"""``openpyxl.worksheet.page`` — page-margin / print-options / page-setup.

Pod 2 (RFC-060 §2.1).  Wolfxl preserves these on round-trip but does
not yet expose Python-side construction; the names land as stubs so
import statements port mechanically.
"""

from __future__ import annotations

from wolfxl._compat import _make_stub

PageMargins = _make_stub(
    "PageMargins",
    "Wolfxl preserves page margins on round-trip; construction lands "
    "in a future release.",
)
PrintOptions = _make_stub(
    "PrintOptions",
    "Wolfxl preserves print options on round-trip; construction lands "
    "in a future release.",
)
PrintPageSetup = _make_stub(
    "PrintPageSetup",
    "Wolfxl preserves page-setup state on round-trip; construction "
    "lands in a future release.",
)


__all__ = ["PageMargins", "PrintOptions", "PrintPageSetup"]
