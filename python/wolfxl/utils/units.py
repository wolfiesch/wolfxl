"""``openpyxl.utils.units``-shaped pixel/EMU/point conversion helpers.

Constants and helpers below mirror openpyxl's ``openpyxl/utils/units.py``
verbatim under its MIT license.  Drawing/image code converts between
EMU (English Metric Units, OOXML's universal unit, 914 400 per inch),
pixels (72 dpi assumption), points (1/72"), and column-width units.

Pod 2 (RFC-060) — these helpers were inlined inside the drawing/image
modules previously; this module is the openpyxl-shaped public path.
"""

from __future__ import annotations

EMU_PER_PIXEL = 9525
EMU_PER_POINT = 12700
EMU_PER_CM = 360000
EMU_PER_INCH = 914400
EMU_PER_MM = 36000


def pixels_to_EMU(value: float) -> int:  # noqa: N802 — openpyxl public name
    """Pixels (96 dpi) → EMU."""
    return int(round(value * EMU_PER_PIXEL))


def EMU_to_pixels(value: float) -> int:  # noqa: N802
    """EMU → pixels (96 dpi, integer-rounded)."""
    return int(round(value / EMU_PER_PIXEL))


def points_to_pixels(value: float) -> int:
    """Points (1/72") → pixels (96 dpi, integer-rounded)."""
    return int(round(value * 96 / 72))


def pixels_to_points(value: float) -> float:
    """Pixels (96 dpi) → points (1/72")."""
    return value * 72 / 96


def cm_to_EMU(value: float) -> int:  # noqa: N802
    return int(round(value * EMU_PER_CM))


def inch_to_EMU(value: float) -> int:  # noqa: N802
    return int(round(value * EMU_PER_INCH))


def mm_to_EMU(value: float) -> int:  # noqa: N802
    return int(round(value * EMU_PER_MM))


__all__ = [
    "EMU_PER_CM",
    "EMU_PER_INCH",
    "EMU_PER_MM",
    "EMU_PER_PIXEL",
    "EMU_PER_POINT",
    "EMU_to_pixels",
    "cm_to_EMU",
    "inch_to_EMU",
    "mm_to_EMU",
    "pixels_to_EMU",
    "pixels_to_points",
    "points_to_pixels",
]
