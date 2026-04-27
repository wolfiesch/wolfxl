"""``openpyxl.styles.protection`` — cell-protection re-export.

Wolfxl exposes :class:`Protection` as a ``NotImplementedError`` stub on the
:mod:`wolfxl.styles` package; this module surfaces it under the
openpyxl-shaped path so import statements port mechanically.

Pod 2 (RFC-060).
"""

from __future__ import annotations

from wolfxl.styles import Protection

__all__ = ["Protection"]
