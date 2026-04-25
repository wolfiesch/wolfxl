"""``wolfxl.styles`` — openpyxl-shape style surface.

Real re-exports for classes wolfxl implements (``Font``, ``PatternFill``,
``Border``, ``Side``, ``Alignment``, ``Color``). Stubs for everything else
— instantiating a stub raises ``NotImplementedError`` with a clear hint
so migrating users see where the gap is rather than a silent failure.
"""

from __future__ import annotations

from wolfxl._compat import _make_stub
from wolfxl._styles import Alignment, Border, Color, Font, PatternFill, Side

NamedStyle = _make_stub(
    "NamedStyle",
    "Named styles are not yet supported; assign Font/Fill/Border directly on cells.",
)
Protection = _make_stub(
    "Protection",
    "Cell protection is not yet supported.",
)
GradientFill = _make_stub(
    "GradientFill",
    "Gradient fills are not yet supported; use PatternFill for solid fills.",
)

__all__ = [
    "Alignment",
    "Border",
    "Color",
    "Font",
    "GradientFill",
    "NamedStyle",
    "PatternFill",
    "Protection",
    "Side",
]
