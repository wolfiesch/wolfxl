"""``openpyxl.worksheet.views`` — sheet-view value types.

Wolfxl tracks freeze-panes, zoom, and selection state internally on
the Worksheet without exposing dedicated value classes.  The
openpyxl-shaped names land here as stubs so import statements port
mechanically; instantiation raises :class:`NotImplementedError`
with a hint pointing at the equivalent ``ws`` attributes
(``ws.freeze_panes``, ``ws.sheet_view``, etc.) when those land.

Pod 2 (RFC-060 §2.1).
"""

from __future__ import annotations

from wolfxl._compat import _make_stub

SheetView = _make_stub(
    "SheetView",
    "Wolfxl tracks sheet-view state inline; use ``ws.freeze_panes`` / "
    "``ws.sheet_view`` accessors when authoring views.",
)
Pane = _make_stub(
    "Pane",
    "Wolfxl manages split panes via ``ws.freeze_panes``; direct Pane "
    "construction is not supported.",
)
Selection = _make_stub(
    "Selection",
    "Wolfxl preserves selection state on round-trip but does not expose "
    "Selection construction.",
)
SheetViewList = _make_stub(
    "SheetViewList",
    "Wolfxl exposes the active view via ``ws.sheet_view``; multi-view "
    "construction is not supported.",
)


__all__ = ["Pane", "Selection", "SheetView", "SheetViewList"]
