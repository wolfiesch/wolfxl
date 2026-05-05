"""Workbook view containers compatible with ``openpyxl.workbook.views``."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class BookView:
    """Workbook window view metadata from ``<workbookView>``."""

    visibility: str = "visible"
    minimized: bool = False
    showHorizontalScroll: bool = True  # noqa: N815
    showVerticalScroll: bool = True  # noqa: N815
    showSheetTabs: bool = True  # noqa: N815
    xWindow: int | None = None  # noqa: N815
    yWindow: int | None = None  # noqa: N815
    windowWidth: int | None = None  # noqa: N815
    windowHeight: int | None = None  # noqa: N815
    tabRatio: int = 600  # noqa: N815
    firstSheet: int = 0  # noqa: N815
    activeTab: int = 0  # noqa: N815
    autoFilterDateGrouping: bool = True  # noqa: N815


__all__ = ["BookView"]
