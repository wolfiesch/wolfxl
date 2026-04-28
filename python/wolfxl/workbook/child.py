"""``openpyxl.workbook.child`` — internal ``_WorkbookChild`` mixin."""

from __future__ import annotations

from typing import Any


class _WorkbookChild:
    """Small openpyxl-compatible base for sheet-like helper objects."""

    _default_title = "Sheet"

    def __init__(self, parent: Any = None, title: str | None = None) -> None:
        self._parent = parent
        self.title = title or ""
        self.HeaderFooter = None

    @property
    def parent(self) -> Any:
        return self._parent

    @property
    def encoding(self) -> str:
        return "utf-8"

    @property
    def path(self) -> str:
        return ""

    def __repr__(self) -> str:
        return f'<{self.__class__.__name__} "{self.title}">'


__all__ = ["_WorkbookChild"]
