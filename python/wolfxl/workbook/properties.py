"""``openpyxl.workbook.properties`` — :class:`CalcProperties` and
:class:`WorkbookProperties` dataclasses (RFC-065).

These two dataclasses back ``wb.calc_properties`` and
``wb.workbook_properties`` respectively. They carry the per-workbook
calc-engine flags (``<calcPr>``) and the workbook-wide configuration
(``<workbookPr>``); both are spliced into ``xl/workbook.xml`` by the
patcher's Phase 2.5q (workbook security drain, extended for RFC-065).

Field names mirror openpyxl exactly (camelCase XML attributes) so
existing user code that pokes at ``wb.calc_properties.calcId`` keeps
working unchanged. Only the ``to_rust_dict`` boundary uses snake_case
to match the §10 contract carried across the PyO3 wall.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass
class CalcProperties:
    """`<calcPr>` element (CT_CalcPr §18.2.2).

    Backs ``wb.calc_properties``. Field defaults match the values Excel
    writes for a freshly-created workbook (``calcId=124519``,
    ``calcMode="auto"``, etc.).
    """

    calcId: int = 124519             # noqa: N815 — openpyxl XML name
    calcMode: str = "auto"           # noqa: N815 — auto | autoNoTable | manual
    fullCalcOnLoad: bool = False     # noqa: N815
    refMode: str = "A1"              # noqa: N815 — A1 | R1C1
    iterate: bool = False
    iterateCount: int = 100          # noqa: N815
    iterateDelta: float = 0.001      # noqa: N815
    fullPrecision: bool = True       # noqa: N815
    calcCompleted: bool = True       # noqa: N815
    calcOnSave: bool = True          # noqa: N815
    concurrentCalc: bool = True      # noqa: N815
    concurrentManualCount: int | None = None  # noqa: N815
    forceFullCalc: bool = False      # noqa: N815

    def to_rust_dict(self) -> dict[str, Any]:
        """Return the §10 snake_case dict consumed by the Rust patcher."""
        return {
            "calc_id": self.calcId,
            "calc_mode": self.calcMode,
            "full_calc_on_load": self.fullCalcOnLoad,
            "ref_mode": self.refMode,
            "iterate": self.iterate,
            "iterate_count": self.iterateCount,
            "iterate_delta": self.iterateDelta,
            "full_precision": self.fullPrecision,
            "calc_completed": self.calcCompleted,
            "calc_on_save": self.calcOnSave,
            "concurrent_calc": self.concurrentCalc,
            "concurrent_manual_count": self.concurrentManualCount,
            "force_full_calc": self.forceFullCalc,
        }


@dataclass
class WorkbookProperties:
    """`<workbookPr>` element (CT_WorkbookPr §18.2.28).

    Backs ``wb.workbook_properties``. Carries the date1904 epoch flag,
    VBA codeName, and a handful of UI / compatibility toggles that
    Excel persists in the workbook part.
    """

    date1904: bool = False
    dateCompatibility: bool = True            # noqa: N815
    showObjects: str = "all"                  # noqa: N815 — all|placeholders|none
    showBorderUnselectedTables: bool = True   # noqa: N815
    filterPrivacy: bool = False               # noqa: N815
    promptedSolutions: bool = False           # noqa: N815
    showInkAnnotation: bool = True            # noqa: N815
    backupFile: bool = False                  # noqa: N815
    saveExternalLinkValues: bool = True       # noqa: N815
    updateLinks: str = "userSet"              # noqa: N815
    codeName: str | None = None               # noqa: N815
    hidePivotFieldList: bool = False          # noqa: N815
    showPivotChartFilter: bool = False        # noqa: N815
    allowRefreshQuery: bool = False           # noqa: N815
    publishItems: bool = False                # noqa: N815
    checkCompatibility: bool = False          # noqa: N815
    autoCompressPictures: bool = True         # noqa: N815
    refreshAllConnections: bool = False       # noqa: N815
    defaultThemeVersion: int = 124226         # noqa: N815

    def to_rust_dict(self) -> dict[str, Any]:
        """Return the §10 snake_case dict consumed by the Rust patcher."""
        return {
            "date1904": self.date1904,
            "date_compatibility": self.dateCompatibility,
            "show_objects": self.showObjects,
            "show_border_unselected_tables": self.showBorderUnselectedTables,
            "filter_privacy": self.filterPrivacy,
            "prompted_solutions": self.promptedSolutions,
            "show_ink_annotation": self.showInkAnnotation,
            "backup_file": self.backupFile,
            "save_external_link_values": self.saveExternalLinkValues,
            "update_links": self.updateLinks,
            "code_name": self.codeName,
            "hide_pivot_field_list": self.hidePivotFieldList,
            "show_pivot_chart_filter": self.showPivotChartFilter,
            "allow_refresh_query": self.allowRefreshQuery,
            "publish_items": self.publishItems,
            "check_compatibility": self.checkCompatibility,
            "auto_compress_pictures": self.autoCompressPictures,
            "refresh_all_connections": self.refreshAllConnections,
            "default_theme_version": self.defaultThemeVersion,
        }


__all__ = ["CalcProperties", "WorkbookProperties"]
