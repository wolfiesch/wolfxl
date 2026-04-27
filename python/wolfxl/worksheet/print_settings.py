"""``openpyxl.worksheet.print_settings`` — print-area / print-titles types.

Pod 2 (RFC-060 §2.1).  Wolfxl preserves print-area + print-titles state
on round-trip but does not yet expose construction; the names land as
stubs so import statements port mechanically.
"""

from __future__ import annotations

from wolfxl._compat import _make_stub

PrintArea = _make_stub(
    "PrintArea",
    "Wolfxl preserves print-area definitions on round-trip; "
    "construction lands in a future release.",
)
PrintTitles = _make_stub(
    "PrintTitles",
    "Wolfxl preserves print titles on round-trip; construction lands "
    "in a future release.",
)
ColRange = _make_stub(
    "ColRange",
    "Wolfxl uses plain strings for print-title column ranges.",
)
RowRange = _make_stub(
    "RowRange",
    "Wolfxl uses plain strings for print-title row ranges.",
)


__all__ = ["ColRange", "PrintArea", "PrintTitles", "RowRange"]
