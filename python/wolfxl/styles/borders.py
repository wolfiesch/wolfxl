"""``openpyxl.styles.borders`` — re-export shim for ``Border`` + ``Side``.

Pod 2 (RFC-060).
"""

from __future__ import annotations

from wolfxl._styles import Border, Side

# openpyxl exposes a frozen tuple of valid border-style names at module
# scope; mirror it for callers that introspect against it.
BORDER_STYLES = (
    "dashDot",
    "dashDotDot",
    "dashed",
    "dotted",
    "double",
    "hair",
    "medium",
    "mediumDashDot",
    "mediumDashDotDot",
    "mediumDashed",
    "slantDashDot",
    "thick",
    "thin",
)

__all__ = ["BORDER_STYLES", "Border", "Side"]
