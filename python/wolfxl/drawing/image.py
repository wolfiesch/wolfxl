"""Shim for ``openpyxl.drawing.image``."""

from __future__ import annotations

from wolfxl._compat import _make_stub

Image = _make_stub(
    "Image",
    "Images are preserved on modify-mode round-trip but cannot be added programmatically.",
)

__all__ = ["Image"]
