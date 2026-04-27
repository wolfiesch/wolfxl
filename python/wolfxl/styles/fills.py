"""``openpyxl.styles.fills`` — re-export shim.

Wolfxl ships :class:`~wolfxl._styles.PatternFill` natively; ``GradientFill``
and the abstract :class:`Fill` base are stubs that raise
:class:`NotImplementedError` on instantiation (matching the existing
behaviour of :mod:`wolfxl.styles`).

Pod 2 (RFC-060).
"""

from __future__ import annotations

from wolfxl._compat import _make_stub
from wolfxl._styles import PatternFill
from wolfxl.styles import GradientFill  # already a stub at the package init.

# Pattern-type vocabulary mirrored from openpyxl for callers that
# introspect against it (e.g. validation pre-processors).
fills = (
    "none",
    "solid",
    "darkGray",
    "mediumGray",
    "lightGray",
    "gray125",
    "gray0625",
    "darkHorizontal",
    "darkVertical",
    "darkDown",
    "darkUp",
    "darkGrid",
    "darkTrellis",
    "lightHorizontal",
    "lightVertical",
    "lightDown",
    "lightUp",
    "lightGrid",
    "lightTrellis",
)

# ``Fill`` is openpyxl's abstract base for every fill kind.  Wolfxl
# never needed it (PatternFill is the only concrete fill we author),
# so it lands here as a stub — instantiation raises with a clear hint.
Fill = _make_stub(
    "Fill",
    "Fill is openpyxl's abstract fill base; construct PatternFill or "
    "GradientFill directly.",
)


__all__ = ["Fill", "GradientFill", "PatternFill", "fills"]
