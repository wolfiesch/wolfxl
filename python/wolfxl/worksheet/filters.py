"""Shim for ``openpyxl.worksheet.filters``."""

from __future__ import annotations

from wolfxl._compat import _make_stub

AutoFilter = _make_stub(
    "AutoFilter",
    "AutoFilter construction is not supported. Set ws.auto_filter.ref directly.",
)
FilterColumn = _make_stub(
    "FilterColumn",
    "FilterColumn construction is not supported.",
)
Filters = _make_stub(
    "Filters",
    "Filters construction is not supported.",
)

__all__ = ["AutoFilter", "FilterColumn", "Filters"]
