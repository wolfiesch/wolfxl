"""``openpyxl.worksheet.header_footer`` — header / footer value types.

Pod 2 (RFC-060 §2.1).  Wolfxl preserves header/footer state on
round-trip but does not yet expose Python-side construction.
"""

from __future__ import annotations

from wolfxl._compat import _make_stub

HeaderFooter = _make_stub(
    "HeaderFooter",
    "Wolfxl preserves header/footer state on round-trip; construction "
    "lands in a future release.",
)
HeaderFooterItem = _make_stub(
    "HeaderFooterItem",
    "Wolfxl preserves header/footer items on round-trip; construction "
    "lands in a future release.",
)


__all__ = ["HeaderFooter", "HeaderFooterItem"]
