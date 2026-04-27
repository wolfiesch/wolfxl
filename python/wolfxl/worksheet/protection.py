"""``openpyxl.worksheet.protection`` — :class:`SheetProtection`.

Pod 2 (RFC-060 §2.1).  Wolfxl preserves sheet-protection state on
round-trip but does not yet expose Python-side construction; the
class lands as a stub so import statements port mechanically.
"""

from __future__ import annotations

from wolfxl._compat import _make_stub

SheetProtection = _make_stub(
    "SheetProtection",
    "Wolfxl preserves sheet-protection state on round-trip; "
    "construction lands in a future release.",
)


__all__ = ["SheetProtection"]
