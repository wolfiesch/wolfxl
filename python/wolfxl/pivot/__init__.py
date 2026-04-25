"""Shim for ``openpyxl.pivot``."""

from __future__ import annotations

from wolfxl._compat import _make_stub

PivotTable = _make_stub(
    "PivotTable",
    "Pivot tables are preserved on modify-mode round-trip but cannot be constructed.",
)

__all__ = ["PivotTable"]
