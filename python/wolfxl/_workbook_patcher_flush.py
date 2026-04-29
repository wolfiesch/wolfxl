"""Modify-mode workbook flush helpers for the Rust patcher."""

from __future__ import annotations

from typing import Any


def flush_pending_hyperlinks_to_patcher(wb: Any) -> None:
    """Drain pending worksheet hyperlinks into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = ws._pending_hyperlinks  # noqa: SLF001
        if not pending:
            continue
        for coord, hl in pending.items():
            if hl is None:
                patcher.queue_hyperlink_delete(ws.title, coord)
                continue
            payload: dict[str, Any] = {}
            if hl.target is not None:
                payload["target"] = hl.target
            if hl.location is not None:
                payload["location"] = hl.location
            if hl.tooltip is not None:
                payload["tooltip"] = hl.tooltip
            if hl.display is not None:
                payload["display"] = hl.display
            patcher.queue_hyperlink(ws.title, coord, payload)
        pending.clear()


def flush_pending_tables_to_patcher(wb: Any) -> None:
    """Drain pending worksheet tables into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = ws._pending_tables  # noqa: SLF001
        if not pending:
            continue
        for table in pending:
            payload: dict[str, Any] = {
                "name": table.name,
                "ref": table.ref,
                "columns": (
                    [column.name for column in table.tableColumns]
                    if table.tableColumns
                    else []
                ),
                "header_row_count": int(table.headerRowCount or 0),
                "totals_row_shown": bool(
                    table.totalsRowCount and table.totalsRowCount > 0
                ),
                "autofilter": True,
            }
            if table.displayName and table.displayName != table.name:
                payload["display_name"] = table.displayName
            if table.tableStyleInfo is not None and table.tableStyleInfo.name:
                payload["style"] = {
                    "name": table.tableStyleInfo.name,
                    "show_first_column": bool(table.tableStyleInfo.showFirstColumn),
                    "show_last_column": bool(table.tableStyleInfo.showLastColumn),
                    "show_row_stripes": bool(table.tableStyleInfo.showRowStripes),
                    "show_column_stripes": bool(table.tableStyleInfo.showColumnStripes),
                }
            patcher.queue_table(ws.title, payload)
        pending.clear()


def flush_pending_images_to_patcher(wb: Any) -> None:
    """Drain pending worksheet images into the Rust patcher."""
    from wolfxl._images import image_to_writer_payload

    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = ws._pending_images  # noqa: SLF001
        if not pending:
            continue
        for image in pending:
            patcher.queue_image_add(ws.title, image_to_writer_payload(image))
        pending.clear()


def flush_pending_comments_to_patcher(wb: Any) -> None:
    """Drain pending worksheet comments into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = ws._pending_comments  # noqa: SLF001
        if not pending:
            continue
        for coord, comment in pending.items():
            if comment is None:
                patcher.queue_comment_delete(ws.title, coord)
                continue
            payload: dict[str, Any] = {
                "text": comment.text,
                "author": comment.author or "wolfxl",
            }
            if getattr(comment, "width", None) is not None:
                payload["width_pt"] = float(comment.width)
            if getattr(comment, "height", None) is not None:
                payload["height_pt"] = float(comment.height)
            patcher.queue_comment(ws.title, coord, payload)
        pending.clear()


def flush_pending_data_validations_to_patcher(wb: Any) -> None:
    """Drain pending worksheet data validations into the Rust patcher."""
    from wolfxl.worksheet.datavalidation import _dv_to_patcher_dict

    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = ws._pending_data_validations  # noqa: SLF001
        if not pending:
            continue
        for data_validation in pending:
            patcher.queue_data_validation(
                ws.title,
                _dv_to_patcher_dict(data_validation),
            )
        pending.clear()


def flush_pending_conditional_formats_to_patcher(wb: Any) -> None:
    """Drain pending worksheet conditional formatting into the Rust patcher."""
    from wolfxl.formatting import _cf_to_patcher_dict

    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = ws._pending_conditional_formats  # noqa: SLF001
        if not pending:
            continue
        by_sqref: dict[str, list[Any]] = {}
        order: list[str] = []
        for sqref, rule in pending:
            if sqref not in by_sqref:
                by_sqref[sqref] = []
                order.append(sqref)
            by_sqref[sqref].append(rule)
        for sqref in order:
            patcher.queue_conditional_formatting(
                ws.title,
                _cf_to_patcher_dict(sqref, by_sqref[sqref]),
            )
        pending.clear()
