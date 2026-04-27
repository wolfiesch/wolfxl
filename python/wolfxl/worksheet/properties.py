"""Worksheet properties (RFC-055 §2.x — Pod 2 re-export targets).

Provides ``WorksheetProperties``, ``PageSetupProperties``, ``Outline``
classes that mirror openpyxl's
``openpyxl.worksheet.properties.{WorksheetProperties, PageSetupProperties, Outline}``.

These are container classes; the actual page-setup contract lives in
``wolfxl.worksheet.page_setup``.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Outline:
    """Outline display properties for sheet rows/columns."""

    summaryBelow: bool = True  # noqa: N815
    summaryRight: bool = True  # noqa: N815
    applyStyles: bool = False  # noqa: N815
    showOutlineSymbols: bool = True  # noqa: N815


@dataclass
class PageSetupProperties:
    """`<pageSetUpPr>` toggles inside `<sheetPr>`."""

    autoPageBreaks: bool = True  # noqa: N815
    fitToPage: bool = False  # noqa: N815


@dataclass
class WorksheetProperties:
    """Container for `<sheetPr>` child elements (CT_SheetPr)."""

    codeName: str | None = None  # noqa: N815
    enableFormatConditionsCalculation: bool | None = None  # noqa: N815
    filterMode: bool | None = None  # noqa: N815
    published: bool | None = None
    syncHorizontal: bool | None = None  # noqa: N815
    syncRef: str | None = None  # noqa: N815
    syncVertical: bool | None = None  # noqa: N815
    transitionEvaluation: bool | None = None  # noqa: N815
    transitionEntry: bool | None = None  # noqa: N815
    tabColor: str | None = None  # noqa: N815  - "RRGGBB" hex
    outlinePr: Outline = field(default_factory=Outline)  # noqa: N815
    pageSetUpPr: PageSetupProperties = field(default_factory=PageSetupProperties)  # noqa: N815


__all__ = [
    "WorksheetProperties",
    "PageSetupProperties",
    "Outline",
]
