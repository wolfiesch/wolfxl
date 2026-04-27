"""``openpyxl.styles.colors`` — re-export shim for ``Color`` + the legacy palette.

Pod 2 (RFC-060).
"""

from __future__ import annotations

from wolfxl._styles import COLOR_INDEX, Color

# Aliases openpyxl exposes for the most common palette positions.
BLACK = "00000000"
WHITE = "00FFFFFF"
RED = "00FF0000"
DARKRED = "00800000"
BLUE = "000000FF"
DARKBLUE = "00000080"
GREEN = "0000FF00"
DARKGREEN = "00008000"
YELLOW = "00FFFF00"
DARKYELLOW = "00808000"


__all__ = [
    "BLACK",
    "BLUE",
    "COLOR_INDEX",
    "Color",
    "DARKBLUE",
    "DARKGREEN",
    "DARKRED",
    "DARKYELLOW",
    "GREEN",
    "RED",
    "WHITE",
    "YELLOW",
]
