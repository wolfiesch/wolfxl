"""Write-mode worksheet flush helpers."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def flush_worksheet(ws: Worksheet) -> None:
    """Write all pending worksheet changes to the active Rust backend."""
    from wolfxl._cell import (
        alignment_to_format_dict,
        border_to_rust_dict,
        fill_to_format_dict,
        font_to_format_dict,
        protection_to_format_dict,
        python_value_to_payload,
    )

    wb = ws._workbook  # noqa: SLF001
    patcher = wb._rust_patcher  # noqa: SLF001
    writer = wb._rust_writer  # noqa: SLF001

    if writer is not None:
        ws._flush_compat_properties(writer)  # noqa: SLF001

    if patcher is not None:
        if ws._append_buffer:  # noqa: SLF001
            ws._materialize_append_buffer()  # noqa: SLF001
        if ws._bulk_writes:  # noqa: SLF001
            ws._materialize_bulk_writes()  # noqa: SLF001
        ws._flush_to_patcher(  # noqa: SLF001
            patcher,
            python_value_to_payload,
            font_to_format_dict,
            fill_to_format_dict,
            alignment_to_format_dict,
            border_to_rust_dict,
            protection_to_format_dict,
        )
    elif writer is not None:
        ws._flush_to_writer(  # noqa: SLF001
            writer,
            python_value_to_payload,
            font_to_format_dict,
            fill_to_format_dict,
            alignment_to_format_dict,
            border_to_rust_dict,
            protection_to_format_dict,
        )
        ws._flush_autofilter_post_cells(writer)  # noqa: SLF001

    ws._dirty.clear()  # noqa: SLF001


def flush_autofilter_post_cells(ws: Worksheet, writer: Any) -> None:
    """Flush the auto-filter after cell values have reached the writer."""
    sheet = ws._title  # noqa: SLF001
    af = ws._auto_filter  # noqa: SLF001
    if _autofilter_has_state(af) and hasattr(writer, "set_autofilter_native"):
        _set_autofilter_native(writer, sheet, af)


def _autofilter_has_state(autofilter: Any) -> bool:
    """Return whether an autofilter contains state worth writing."""
    return (
        autofilter.ref is not None
        or bool(autofilter.filter_columns)
        or autofilter.sort_state is not None
    )


def _set_autofilter_native(writer: Any, sheet: str, autofilter: Any) -> None:
    """Set one native autofilter, preserving the defensive save path."""
    try:
        writer.set_autofilter_native(sheet, autofilter.to_rust_dict())
    except Exception:
        # Defensive: do not poison the save path on a malformed filter spec.
        pass


def flush_compat_properties(ws: Worksheet, writer: Any) -> None:
    """Flush openpyxl compatibility metadata to the write-mode backend."""
    sheet = ws._title  # noqa: SLF001

    _flush_sheet_layout(ws, writer, sheet)
    _flush_sheet_setup(ws, writer, sheet)
    _flush_page_breaks(ws, writer, sheet)
    _flush_pending_hyperlinks(ws, writer, sheet)
    _flush_pending_comments(ws, writer, sheet)
    _flush_pending_threaded_comments(ws, writer, sheet)
    _flush_pending_tables(ws, writer, sheet)
    _flush_pending_data_validations(ws, writer, sheet)
    _flush_pending_conditional_formats(ws, writer, sheet)
    _flush_pending_images(ws, writer, sheet)
    _flush_pending_charts(ws, writer, sheet)


def _flush_sheet_layout(ws: Worksheet, writer: Any, sheet: str) -> None:
    """Flush freeze panes, dimensions, and print area metadata."""
    if ws._freeze_panes is not None:  # noqa: SLF001
        writer.set_freeze_panes(
            sheet, {"mode": "freeze", "top_left_cell": ws._freeze_panes},  # noqa: SLF001
        )

    for row_num, height in ws._row_heights.items():  # noqa: SLF001
        if height is not None:
            writer.set_row_height(sheet, row_num, height)

    for col_letter, width in ws._col_widths.items():  # noqa: SLF001
        if width is not None:
            writer.set_column_width(sheet, col_letter, width)

    if ws._print_area is not None and hasattr(writer, "set_print_area"):  # noqa: SLF001
        writer.set_print_area(sheet, ws._print_area)  # noqa: SLF001


def _flush_sheet_setup(ws: Worksheet, writer: Any, sheet: str) -> None:
    """Flush page setup, margins, headers, views, protection, and titles."""
    if not hasattr(writer, "set_sheet_setup_native"):
        return
    if not _has_sheet_setup(ws):
        return
    _set_sheet_setup_native(writer, sheet, ws)


def _has_sheet_setup(ws: Worksheet) -> bool:
    """Return whether a worksheet has pending sheet setup state."""
    return (
        ws._page_setup is not None  # noqa: SLF001
        or ws._page_margins is not None  # noqa: SLF001
        or ws._header_footer is not None  # noqa: SLF001
        or ws._sheet_view is not None  # noqa: SLF001
        or ws._protection is not None  # noqa: SLF001
        or getattr(ws, "_print_title_rows", None) is not None
        or getattr(ws, "_print_title_cols", None) is not None
    )


def _sheet_setup_payload(ws: Worksheet) -> dict[str, Any] | None:
    """Build the native writer payload for worksheet setup metadata."""
    payload = ws.to_rust_setup_dict()
    if any(value is not None for value in payload.values()):
        return payload
    return None


def _set_sheet_setup_native(writer: Any, sheet: str, ws: Worksheet) -> None:
    """Set one native sheet setup payload, preserving the defensive save path."""
    try:
        payload = _sheet_setup_payload(ws)
        if payload is not None:
            writer.set_sheet_setup_native(sheet, payload)
    except Exception:
        # Defensive: Python wrapper validators should already reject bad specs.
        pass


def _flush_page_breaks(ws: Worksheet, writer: Any, sheet: str) -> None:
    """Flush page breaks and sheet format metadata."""
    if not hasattr(writer, "set_page_breaks_native"):
        return
    if not _has_page_breaks(ws):
        return
    _set_page_breaks_native(writer, sheet, ws)


def _has_page_breaks(ws: Worksheet) -> bool:
    """Return whether a worksheet has pending page-break or sheet-format state."""
    return (
        ws._row_breaks is not None  # noqa: SLF001
        or ws._col_breaks is not None  # noqa: SLF001
        or ws._sheet_format is not None  # noqa: SLF001
    )


def _page_breaks_payload(ws: Worksheet) -> dict[str, Any] | None:
    """Build the native writer payload for page breaks and sheet format."""
    breaks_dict = ws.to_rust_page_breaks_dict()
    payload = {
        "row_breaks": breaks_dict.get("row_breaks"),
        "col_breaks": breaks_dict.get("col_breaks"),
        "sheet_format": ws.to_rust_sheet_format_dict(),
    }
    if any(value is not None for value in payload.values()):
        return payload
    return None


def _set_page_breaks_native(writer: Any, sheet: str, ws: Worksheet) -> None:
    """Set one native page-break payload, preserving the defensive save path."""
    try:
        payload = _page_breaks_payload(ws)
        if payload is not None:
            writer.set_page_breaks_native(sheet, payload)
    except Exception:
        # Defensive: do not poison the save path.
        pass


def _flush_pending_hyperlinks(ws: Worksheet, writer: Any, sheet: str) -> None:
    """Flush queued write-mode hyperlinks."""
    if not ws._pending_hyperlinks:  # noqa: SLF001
        return
    for coord, hl in ws._pending_hyperlinks.items():  # noqa: SLF001
        if hl is None:
            continue
        target = hl.target
        internal = False
        if target is None and hl.location is not None:
            target = hl.location
            internal = True
        if not target:
            continue
        writer.add_hyperlink(sheet, _hyperlink_payload(coord, hl, target, internal))
    ws._pending_hyperlinks.clear()  # noqa: SLF001


def _hyperlink_payload(
    coord: str,
    hyperlink: Any,
    target: str,
    internal: bool,
) -> dict[str, Any]:
    """Build the native writer payload for a worksheet hyperlink."""
    return {
        "cell": coord,
        "target": target,
        "display": hyperlink.display,
        "tooltip": hyperlink.tooltip,
        "internal": internal,
    }


def _flush_pending_comments(ws: Worksheet, writer: Any, sheet: str) -> None:
    """Flush queued write-mode comments."""
    if not ws._pending_comments:  # noqa: SLF001
        return
    for coord, comment in ws._pending_comments.items():  # noqa: SLF001
        if comment is None:
            continue
        writer.add_comment(sheet, _comment_payload(coord, comment))
    ws._pending_comments.clear()  # noqa: SLF001


def _comment_payload(coord: str, comment: Any) -> dict[str, Any]:
    """Build the native writer payload for a worksheet comment."""
    return {
        "cell": coord,
        "text": comment.text,
        "author": comment.author,
    }


def _flush_pending_threaded_comments(ws: Worksheet, writer: Any, sheet: str) -> None:
    """Flush queued write-mode threaded comments (RFC-068 / G08).

    Walks every top-level thread and emits one ``add_threaded_comment`` call
    per top-level + reply, with parent_id resolved to the parent's GUID.
    Person GUIDs are emitted via the workbook-level person flush; this helper
    only references them.
    """
    pending = ws._pending_threaded_comments  # noqa: SLF001
    if not pending:
        return
    if not hasattr(writer, "add_threaded_comment"):
        # Older native backend (or modify-mode patcher) does not yet expose
        # the entry point. Drop silently — Step 5 lights up modify mode.
        return

    for coord, top in pending.items():
        if top is None:
            continue
        # Ensure GUID + timestamp on the top-level thread before emitting.
        top.ensure_id()
        top.ensure_created()
        writer.add_threaded_comment(
            sheet,
            _threaded_comment_payload(coord, top, parent_id=None),
        )
        for reply in top.replies:
            reply.ensure_id()
            reply.ensure_created()
            writer.add_threaded_comment(
                sheet,
                _threaded_comment_payload(coord, reply, parent_id=top.id),
            )
    pending.clear()


def _threaded_comment_payload(
    coord: str, tc: Any, *, parent_id: str | None
) -> dict[str, Any]:
    """Build the native writer payload for one ThreadedComment.

    The Rust backend wants `created` as an ISO 8601 string (Excel writes
    millisecond precision, e.g. ``2024-09-12T15:31:01.42``). The
    ``ensure_created()`` call on the caller side guarantees we never see
    None here.
    """
    person_id = tc.person.id if tc.person is not None else ""
    return {
        "id": tc.id,
        "cell": coord,
        "person_id": person_id,
        "created": tc.created.isoformat(timespec="milliseconds"),
        "parent_id": parent_id,
        "text": tc.text,
        "done": tc.done,
    }


def _flush_pending_tables(ws: Worksheet, writer: Any, sheet: str) -> None:
    """Flush queued write-mode tables."""
    if not ws._pending_tables:  # noqa: SLF001
        return
    for table in ws._pending_tables:  # noqa: SLF001
        writer.add_table(sheet, _table_payload(table))
    ws._pending_tables.clear()  # noqa: SLF001


def _table_payload(table: Any) -> dict[str, Any]:
    """Build the native writer payload for a worksheet table."""
    style_name = table.tableStyleInfo.name if table.tableStyleInfo else None
    col_names = [col.name for col in table.tableColumns] if table.tableColumns else []
    return {
        "name": table.name,
        "ref": table.ref,
        "style": style_name,
        "columns": col_names,
        "header_row": table.headerRowCount > 0,
        "totals_row": table.totalsRowCount > 0,
    }


def _flush_pending_data_validations(ws: Worksheet, writer: Any, sheet: str) -> None:
    """Flush queued write-mode data validations."""
    if not ws._pending_data_validations:  # noqa: SLF001
        return
    for dv in ws._pending_data_validations:  # noqa: SLF001
        writer.add_data_validation(sheet, _data_validation_payload(dv))
    ws._pending_data_validations.clear()  # noqa: SLF001


def _data_validation_payload(data_validation: Any) -> dict[str, Any]:
    """Build the native writer payload for a worksheet data validation."""
    return {
        "range": str(data_validation.sqref),
        "validation_type": data_validation.type,
        "operator": data_validation.operator,
        "formula1": data_validation.formula1,
        "formula2": data_validation.formula2,
        "allow_blank": data_validation.allowBlank,
        "error_title": data_validation.errorTitle,
        "error": data_validation.error,
    }


def _flush_pending_conditional_formats(ws: Worksheet, writer: Any, sheet: str) -> None:
    """Flush queued write-mode conditional formats."""
    if not ws._pending_conditional_formats:  # noqa: SLF001
        return
    for range_string, rule in ws._pending_conditional_formats:  # noqa: SLF001
        writer.add_conditional_format(
            sheet, _conditional_format_payload(range_string, rule)
        )
    ws._pending_conditional_formats.clear()  # noqa: SLF001


def _conditional_format_payload(range_string: str, rule: Any) -> dict[str, Any]:
    """Build the native writer payload for a worksheet conditional format."""
    formula = rule.formula[0] if rule.formula else None
    return {
        "range": range_string,
        "rule_type": rule.type,
        "operator": rule.operator,
        "formula": formula,
        "stop_if_true": rule.stopIfTrue,
    }


def _flush_pending_images(ws: Worksheet, writer: Any, sheet: str) -> None:
    """Flush queued write-mode images."""
    if not ws._pending_images or not hasattr(writer, "add_image"):  # noqa: SLF001
        return
    from wolfxl._images import image_to_writer_payload

    for image in ws._pending_images:  # noqa: SLF001
        writer.add_image(sheet, image_to_writer_payload(image))
    ws._pending_images.clear()  # noqa: SLF001


def _flush_pending_charts(ws: Worksheet, writer: Any, sheet: str) -> None:
    """Flush queued write-mode charts."""
    if not ws._pending_charts:  # noqa: SLF001
        return
    if hasattr(writer, "add_chart_native"):
        for chart in ws._pending_charts:  # noqa: SLF001
            writer.add_chart_native(sheet, chart.to_rust_dict(), chart._anchor)  # noqa: SLF001
    else:
        import warnings

        warnings.warn(
            "wolfxl.chart: native chart write requires Pod-alpha's "
            "add_chart_native binding (not yet available). "
            f"Dropping {len(ws._pending_charts)} chart(s) on sheet {sheet!r}.",  # noqa: SLF001
            RuntimeWarning,
            stacklevel=2,
        )
    ws._pending_charts.clear()  # noqa: SLF001
