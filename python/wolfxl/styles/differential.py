"""Shim for ``openpyxl.styles.differential``."""

from __future__ import annotations

from wolfxl._compat import _make_stub

DifferentialStyle = _make_stub(
    "DifferentialStyle",
    "Differential styles are tied to conditional formatting, which wolfxl preserves "
    "in modify mode but does not construct.",
)

__all__ = ["DifferentialStyle"]
