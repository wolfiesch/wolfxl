"""Worksheet setup, print, view, protection, and page-break helpers."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def get_freeze_panes(ws: Worksheet) -> str | None:
    """Return the worksheet freeze-pane top-left cell, if any."""
    workbook = ws._workbook  # noqa: SLF001
    if workbook._rust_reader is not None and ws._freeze_panes is None:  # noqa: SLF001
        info = workbook._rust_reader.read_freeze_panes(ws._title)  # noqa: SLF001
        if info and info.get("mode"):
            return info.get("top_left_cell")
        return None
    return ws._freeze_panes  # noqa: SLF001


def set_freeze_panes(ws: Worksheet, value: str | None) -> None:
    """Set freeze panes and mirror the state into ``sheet_view.pane``."""
    ws._freeze_panes = value  # noqa: SLF001
    if ws._sheet_view is None:  # noqa: SLF001
        return

    from wolfxl.worksheet.views import Pane

    if value is None:
        ws._sheet_view.pane = None  # noqa: SLF001
        return

    from wolfxl._utils import a1_to_rowcol

    try:
        row, col = a1_to_rowcol(value)
    except Exception:
        return
    ws._sheet_view.pane = Pane(  # noqa: SLF001
        xSplit=float(col - 1),
        ySplit=float(row - 1),
        topLeftCell=value,
        activePane="bottomRight",
        state="frozen",
    )


def get_auto_filter(ws: Worksheet) -> Any:
    """Return the worksheet auto-filter proxy, loading reader metadata once."""
    workbook = ws._workbook  # noqa: SLF001
    auto_filter = ws._auto_filter  # noqa: SLF001
    if (
        auto_filter.ref is None
        and not auto_filter.filter_columns
        and auto_filter.sort_state is None
        and workbook._rust_reader is not None  # noqa: SLF001
        and hasattr(workbook._rust_reader, "read_auto_filter")  # noqa: SLF001
    ):
        payload = workbook._rust_reader.read_auto_filter(ws._title)  # noqa: SLF001
        if isinstance(payload, dict):
            auto_filter.ref = payload.get("ref")
            auto_filter.filter_columns = [
                _filter_column_from_payload(column)
                for column in payload.get("filter_columns", [])
                if isinstance(column, dict)
            ]
            auto_filter.sort_state = _sort_state_from_payload(payload.get("sort_state"))
    return auto_filter


def _filter_column_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.worksheet.filters import FilterColumn

    return FilterColumn(
        col_id=int(payload.get("col_id", 0)),
        hidden_button=bool(payload.get("hidden_button", False)),
        show_button=bool(payload.get("show_button", True)),
        filter=_filter_from_payload(payload.get("filter")),
        date_group_items=[
            _date_group_item_from_payload(item)
            for item in payload.get("date_group_items", [])
            if isinstance(item, dict)
        ],
    )


def _filter_from_payload(payload: Any) -> Any:
    if not isinstance(payload, dict):
        return None
    from wolfxl.worksheet.filters import (
        BlankFilter,
        ColorFilter,
        CustomFilter,
        CustomFilters,
        DynamicFilter,
        IconFilter,
        StringFilter,
        Top10,
    )

    kind = payload.get("kind")
    if kind == "blank":
        return BlankFilter()
    if kind == "color":
        return ColorFilter(
            dxf_id=int(payload.get("dxf_id", 0)),
            cell_color=bool(payload.get("cell_color", True)),
        )
    if kind == "custom":
        return CustomFilters(
            customFilter=[
                CustomFilter(
                    operator=str(item.get("operator", "equal")),
                    val=str(item.get("val", "")),
                )
                for item in payload.get("filters", [])
                if isinstance(item, dict)
            ],
            and_=bool(payload.get("and_", False)),
        )
    if kind == "dynamic":
        return DynamicFilter(
            type=str(payload.get("type", "null")),
            val=payload.get("val"),
            val_iso=payload.get("val_iso"),
            max_val_iso=payload.get("max_val_iso"),
        )
    if kind == "icon":
        return IconFilter(
            icon_set=str(payload.get("icon_set", "3Arrows")),
            icon_id=int(payload.get("icon_id", 0)),
        )
    if kind == "string":
        return StringFilter(values=[str(value) for value in payload.get("values", [])])
    if kind == "top10":
        return Top10(
            top=bool(payload.get("top", True)),
            percent=bool(payload.get("percent", False)),
            val=float(payload.get("val", 10.0)),
            filter_val=payload.get("filter_val"),
        )
    return None


def _date_group_item_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.worksheet.filters import DateGroupItem

    return DateGroupItem(
        year=int(payload.get("year", 0)),
        month=payload.get("month"),
        day=payload.get("day"),
        hour=payload.get("hour"),
        minute=payload.get("minute"),
        second=payload.get("second"),
        date_time_grouping=str(payload.get("date_time_grouping", "year")),
    )


def _sort_state_from_payload(payload: Any) -> Any:
    if not isinstance(payload, dict):
        return None
    from wolfxl.worksheet.filters import SortState

    return SortState(
        sort_conditions=[
            _sort_condition_from_payload(condition)
            for condition in payload.get("sort_conditions", [])
            if isinstance(condition, dict)
        ],
        column_sort=bool(payload.get("column_sort", False)),
        case_sensitive=bool(payload.get("case_sensitive", False)),
        ref=payload.get("ref"),
    )


def _sort_condition_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.worksheet.filters import SortCondition

    return SortCondition(
        ref=str(payload.get("ref", "")),
        descending=bool(payload.get("descending", False)),
        sort_by=str(payload.get("sort_by", "value")),
        custom_list=payload.get("custom_list"),
        dxf_id=payload.get("dxf_id"),
        icon_set=payload.get("icon_set"),
        icon_id=payload.get("icon_id"),
    )


def get_page_setup(ws: Worksheet) -> Any:
    """Return the lazy page setup object."""
    if ws._page_setup is None:  # noqa: SLF001
        from wolfxl.worksheet.page_setup import PageSetup

        ws._page_setup = PageSetup()  # noqa: SLF001
    return ws._page_setup  # noqa: SLF001


def get_page_margins(ws: Worksheet) -> Any:
    """Return the lazy page margins object."""
    if ws._page_margins is None:  # noqa: SLF001
        from wolfxl.worksheet.page_setup import PageMargins

        ws._page_margins = PageMargins()  # noqa: SLF001
    return ws._page_margins  # noqa: SLF001


def get_header_footer(ws: Worksheet) -> Any:
    """Return the lazy header/footer object."""
    if ws._header_footer is None:  # noqa: SLF001
        from wolfxl.worksheet.header_footer import HeaderFooter

        ws._header_footer = HeaderFooter()  # noqa: SLF001
    return ws._header_footer  # noqa: SLF001


def get_sheet_view(ws: Worksheet) -> Any:
    """Return the lazy sheet view object, carrying pending freeze panes."""
    if ws._sheet_view is None:  # noqa: SLF001
        from wolfxl.worksheet.views import Pane, SheetView

        sheet_view = SheetView()
        if ws._freeze_panes is not None:  # noqa: SLF001
            from wolfxl._utils import a1_to_rowcol

            try:
                row, col = a1_to_rowcol(ws._freeze_panes)  # noqa: SLF001
                sheet_view.pane = Pane(
                    xSplit=float(col - 1),
                    ySplit=float(row - 1),
                    topLeftCell=ws._freeze_panes,  # noqa: SLF001
                    activePane="bottomRight",
                    state="frozen",
                )
            except Exception:
                pass
        ws._sheet_view = sheet_view  # noqa: SLF001
    return ws._sheet_view  # noqa: SLF001


def get_protection(ws: Worksheet) -> Any:
    """Return the lazy sheet protection object."""
    if ws._protection is None:  # noqa: SLF001
        workbook = ws._workbook  # noqa: SLF001
        payload = None
        reader = getattr(workbook, "_rust_reader", None)
        if reader is not None and hasattr(reader, "read_sheet_protection"):
            payload = reader.read_sheet_protection(ws._title)  # noqa: SLF001
        ws._protection = _sheet_protection_from_payload(payload)  # noqa: SLF001
    return ws._protection  # noqa: SLF001


def _sheet_protection_from_payload(payload: Any) -> Any:
    from wolfxl.worksheet.protection import SheetProtection

    if not isinstance(payload, dict):
        return SheetProtection()
    return SheetProtection(
        sheet=bool(payload.get("sheet", False)),
        objects=bool(payload.get("objects", False)),
        scenarios=bool(payload.get("scenarios", False)),
        formatCells=bool(payload.get("format_cells", True)),
        formatColumns=bool(payload.get("format_columns", True)),
        formatRows=bool(payload.get("format_rows", True)),
        insertColumns=bool(payload.get("insert_columns", True)),
        insertRows=bool(payload.get("insert_rows", True)),
        insertHyperlinks=bool(payload.get("insert_hyperlinks", True)),
        deleteColumns=bool(payload.get("delete_columns", True)),
        deleteRows=bool(payload.get("delete_rows", True)),
        selectLockedCells=bool(payload.get("select_locked_cells", False)),
        sort=bool(payload.get("sort", True)),
        autoFilter=bool(payload.get("auto_filter", True)),
        pivotTables=bool(payload.get("pivot_tables", True)),
        selectUnlockedCells=bool(payload.get("select_unlocked_cells", False)),
        password=payload.get("password_hash"),
    )


def get_row_breaks(ws: Worksheet) -> Any:
    """Return the lazy row page-break list."""
    if ws._row_breaks is None:  # noqa: SLF001
        from wolfxl.worksheet.pagebreak import PageBreakList

        ws._row_breaks = PageBreakList()  # noqa: SLF001
    return ws._row_breaks  # noqa: SLF001


def get_col_breaks(ws: Worksheet) -> Any:
    """Return the lazy column page-break list."""
    if ws._col_breaks is None:  # noqa: SLF001
        from wolfxl.worksheet.pagebreak import PageBreakList

        ws._col_breaks = PageBreakList()  # noqa: SLF001
    return ws._col_breaks  # noqa: SLF001


def get_sheet_format(ws: Worksheet) -> Any:
    """Return the lazy sheet format properties object."""
    if ws._sheet_format is None:  # noqa: SLF001
        from wolfxl.worksheet.dimensions import SheetFormatProperties

        ws._sheet_format = SheetFormatProperties()  # noqa: SLF001
    return ws._sheet_format  # noqa: SLF001


def get_dimension_holder(ws: Worksheet) -> Any:
    """Return a fresh dimension holder bound to a worksheet."""
    from wolfxl.worksheet.dimensions import DimensionHolder

    return DimensionHolder(ws)


def to_rust_page_breaks_dict(ws: Worksheet) -> dict[str, Any]:
    """Return the Rust payload shape for row and column page breaks."""
    return {
        "row_breaks": (
            ws._row_breaks.to_rust_dict()  # noqa: SLF001
            if ws._row_breaks is not None and len(ws._row_breaks) > 0  # noqa: SLF001
            else None
        ),
        "col_breaks": (
            ws._col_breaks.to_rust_dict()  # noqa: SLF001
            if ws._col_breaks is not None and len(ws._col_breaks) > 0  # noqa: SLF001
            else None
        ),
    }


def to_rust_sheet_format_dict(ws: Worksheet) -> dict[str, Any] | None:
    """Return the Rust payload shape for sheet format properties."""
    if ws._sheet_format is None or ws._sheet_format.is_default():  # noqa: SLF001
        return None
    return ws._sheet_format.to_rust_dict()  # noqa: SLF001


def set_print_title_rows(ws: Worksheet, value: str | None) -> None:
    """Set repeat rows for printed pages."""
    if value is not None:
        from wolfxl.worksheet.print_settings import RowRange

        ws._print_title_rows = str(RowRange.from_string(value))  # noqa: SLF001
    else:
        ws._print_title_rows = None  # noqa: SLF001


def set_print_title_cols(ws: Worksheet, value: str | None) -> None:
    """Set repeat columns for printed pages."""
    if value is not None:
        from wolfxl.worksheet.print_settings import ColRange

        ws._print_title_cols = str(ColRange.from_string(value))  # noqa: SLF001
    else:
        ws._print_title_cols = None  # noqa: SLF001


def to_rust_setup_dict(ws: Worksheet) -> dict[str, Any]:
    """Return the Rust payload shape for worksheet setup metadata."""
    payload: dict[str, Any] = {}
    payload["page_setup"] = (
        ws._page_setup.to_rust_dict()  # noqa: SLF001
        if ws._page_setup is not None and not ws._page_setup.is_default()  # noqa: SLF001
        else None
    )
    payload["page_margins"] = (
        ws._page_margins.to_rust_dict()  # noqa: SLF001
        if ws._page_margins is not None and not ws._page_margins.is_default()  # noqa: SLF001
        else None
    )
    payload["header_footer"] = (
        ws._header_footer.to_rust_dict()  # noqa: SLF001
        if ws._header_footer is not None and not ws._header_footer.is_default()  # noqa: SLF001
        else None
    )
    payload["sheet_view"] = (
        ws._sheet_view.to_rust_dict()  # noqa: SLF001
        if ws._sheet_view is not None and not ws._sheet_view.is_default()  # noqa: SLF001
        else None
    )
    payload["sheet_protection"] = (
        ws._protection.to_rust_dict()  # noqa: SLF001
        if ws._protection is not None and not ws._protection.is_default()  # noqa: SLF001
        else None
    )
    if ws._print_title_rows is not None or ws._print_title_cols is not None:  # noqa: SLF001
        payload["print_titles"] = {
            "rows": ws._print_title_rows,  # noqa: SLF001
            "cols": ws._print_title_cols,  # noqa: SLF001
        }
    else:
        payload["print_titles"] = None
    return payload
