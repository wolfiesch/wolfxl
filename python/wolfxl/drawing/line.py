"""Re-export of :class:`wolfxl.chart.shapes.LineProperties`.

openpyxl exposes ``LineProperties`` at ``openpyxl.drawing.line.LineProperties``;
mirroring that import path is required for source compatibility with
existing openpyxl-based code.

Sprint Μ Pod-β (RFC-046) — added during integrator finalize.
"""

from __future__ import annotations

from wolfxl.chart.shapes import LineProperties

__all__ = ["LineProperties"]
