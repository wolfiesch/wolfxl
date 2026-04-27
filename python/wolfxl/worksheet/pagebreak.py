"""``openpyxl.worksheet.pagebreak`` — page-break value types.

Pod 2 (RFC-060 §2.1).  Wolfxl preserves page breaks on round-trip but
does not author them from Python yet; the classes land as stubs so
import statements port mechanically.
"""

from __future__ import annotations

from wolfxl._compat import _make_stub

Break = _make_stub(
    "Break",
    "Wolfxl preserves page breaks on round-trip but does not yet "
    "expose construction.",
)
ColBreak = _make_stub(
    "ColBreak",
    "Wolfxl preserves column page breaks on round-trip but does not "
    "yet expose construction.",
)
RowBreak = _make_stub(
    "RowBreak",
    "Wolfxl preserves row page breaks on round-trip but does not "
    "yet expose construction.",
)
PageBreak = RowBreak


__all__ = ["Break", "ColBreak", "PageBreak", "RowBreak"]
