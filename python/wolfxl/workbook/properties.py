"""``openpyxl.workbook.properties`` — :class:`CalcProperties` shim.

Pod 2 (RFC-060 §2.6).  Wolfxl preserves calc-mode flags on round-trip
but does not yet expose Python-side construction.
"""

from __future__ import annotations

from wolfxl._compat import _make_stub

CalcProperties = _make_stub(
    "CalcProperties",
    "Wolfxl preserves calc-properties on round-trip; construction "
    "lands in a future release.",
)
WorkbookProperties = _make_stub(
    "WorkbookProperties",
    "Wolfxl preserves workbook properties on round-trip; construction "
    "lands in a future release.",
)


__all__ = ["CalcProperties", "WorkbookProperties"]
