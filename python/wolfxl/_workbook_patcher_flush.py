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
            patcher.queue_hyperlink(ws.title, coord, _hyperlink_payload(hl))
        pending.clear()


def _hyperlink_payload(hyperlink: Any) -> dict[str, Any]:
    """Build the Rust patcher payload for a worksheet hyperlink."""
    payload: dict[str, Any] = {}
    if hyperlink.target is not None:
        payload["target"] = hyperlink.target
    if hyperlink.location is not None:
        payload["location"] = hyperlink.location
    if hyperlink.tooltip is not None:
        payload["tooltip"] = hyperlink.tooltip
    if hyperlink.display is not None:
        payload["display"] = hyperlink.display
    return payload


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
            patcher.queue_table(ws.title, _table_payload(table))
        pending.clear()


def _table_payload(table: Any) -> dict[str, Any]:
    """Build the Rust patcher payload for a worksheet table."""
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
    return payload


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
            patcher.queue_comment(ws.title, coord, _comment_payload(comment))
        pending.clear()


def _comment_payload(comment: Any) -> dict[str, Any]:
    """Build the Rust patcher payload for a worksheet comment."""
    payload: dict[str, Any] = {
        "text": comment.text,
        "author": comment.author or "wolfxl",
    }
    if getattr(comment, "width", None) is not None:
        payload["width_pt"] = float(comment.width)
    if getattr(comment, "height", None) is not None:
        payload["height_pt"] = float(comment.height)
    return payload


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


def flush_pending_axis_shifts_to_patcher(wb: Any) -> None:
    """Drain pending row/column structural shifts into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None or not wb._pending_axis_shifts:  # noqa: SLF001
        return
    for sheet_title, axis, idx, n in wb._pending_axis_shifts:  # noqa: SLF001
        patcher.queue_axis_shift(sheet_title, axis, idx, n)
    wb._pending_axis_shifts.clear()  # noqa: SLF001


def flush_pending_range_moves_to_patcher(wb: Any) -> None:
    """Drain pending range moves into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None or not wb._pending_range_moves:  # noqa: SLF001
        return
    for (
        sheet_title,
        src_min_col,
        src_min_row,
        src_max_col,
        src_max_row,
        d_row,
        d_col,
        translate,
    ) in wb._pending_range_moves:  # noqa: SLF001
        patcher.queue_range_move(
            sheet_title,
            src_min_col,
            src_min_row,
            src_max_col,
            src_max_row,
            d_row,
            d_col,
            translate,
        )
    wb._pending_range_moves.clear()  # noqa: SLF001


def flush_pending_sheet_copies_to_patcher(wb: Any) -> None:
    """Drain pending sheet-copy operations into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None or not wb._pending_sheet_copies:  # noqa: SLF001
        return
    for src_title, dst_title, deep_copy_images in wb._pending_sheet_copies:  # noqa: SLF001
        patcher.queue_sheet_copy(src_title, dst_title, deep_copy_images)
    wb._pending_sheet_copies.clear()  # noqa: SLF001


def queue_sheet_move_to_patcher(wb: Any, name: str, offset: int) -> None:
    """Queue one sheet move into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    patcher.queue_sheet_move(name, offset)


def flush_defined_names_to_patcher(wb: Any) -> None:
    """Drain pending workbook defined-name writes into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None or not wb._pending_defined_names:  # noqa: SLF001
        return
    for defined_name in wb._pending_defined_names.values():  # noqa: SLF001
        patcher.queue_defined_name(_defined_name_payload(defined_name))
    wb._pending_defined_names.clear()  # noqa: SLF001


def _defined_name_payload(defined_name: Any) -> dict[str, Any]:
    """Build the Rust patcher payload for a workbook defined name."""
    payload: dict[str, Any] = {
        "name": defined_name.name,
        "formula": defined_name.value,
    }
    if defined_name.localSheetId is not None:
        payload["local_sheet_id"] = defined_name.localSheetId
    if defined_name.hidden:
        payload["hidden"] = True
    if defined_name.comment is not None:
        payload["comment"] = defined_name.comment
    return payload


def build_security_dict(wb: Any) -> dict[str, Any]:
    """Build the Rust patcher payload for workbook security blocks."""
    return {
        "workbook_protection": (
            wb._security.to_dict() if wb._security is not None else None  # noqa: SLF001
        ),
        "file_sharing": (
            wb._file_sharing.to_dict()  # noqa: SLF001
            if wb._file_sharing is not None  # noqa: SLF001
            else None
        ),
    }


def flush_security_to_patcher(wb: Any) -> None:
    """Drain workbook security metadata into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None or not wb._pending_security_update:  # noqa: SLF001
        return
    patcher.queue_workbook_security(build_security_dict(wb))
    wb._pending_security_update = False  # noqa: SLF001


def flush_properties_to_patcher(wb: Any) -> None:
    """Drain dirty workbook document properties into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    props = wb._properties_cache  # noqa: SLF001
    if props is None:
        wb._properties_dirty = False  # noqa: SLF001
        return

    patcher.queue_properties(_properties_payload(wb, props))
    wb._properties_dirty = False  # noqa: SLF001


def _properties_payload(wb: Any, props: Any) -> dict[str, Any]:
    """Build the Rust patcher payload for workbook document properties."""
    user_set: set[str] = getattr(props, "_user_set", set())
    modified_iso: str | None = None
    if "modified" in user_set and props.modified is not None:
        modified_iso = props.modified.isoformat()
    payload: dict[str, Any] = {
        "title": props.title,
        "subject": props.subject,
        "creator": props.creator,
        "keywords": props.keywords,
        "description": props.description,
        "last_modified_by": props.lastModifiedBy,
        "category": props.category,
        "content_status": props.contentStatus,
        "created_iso": props.created.isoformat() if props.created else None,
        "modified_iso": modified_iso,
        "sheet_names": list(wb._sheet_names),  # noqa: SLF001
    }
    return {key: value for key, value in payload.items() if value is not None}


def flush_pending_sheet_setup_to_patcher(wb: Any) -> None:
    """Drain pending worksheet setup metadata into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        if (
            ws._page_setup is None  # noqa: SLF001
            and ws._page_margins is None  # noqa: SLF001
            and ws._header_footer is None  # noqa: SLF001
            and ws._sheet_view is None  # noqa: SLF001
            and ws._protection is None  # noqa: SLF001
            and getattr(ws, "_print_title_rows", None) is None
            and getattr(ws, "_print_title_cols", None) is None
        ):
            continue
        payload = ws.to_rust_setup_dict()
        if all(value is None for value in payload.values()):
            continue
        patcher.queue_sheet_setup_update(ws.title, payload)


def flush_pending_page_breaks_to_patcher(wb: Any) -> None:
    """Drain pending page-break and sheet-format metadata into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        row_breaks = getattr(ws, "_row_breaks", None)
        col_breaks = getattr(ws, "_col_breaks", None)
        sheet_format = getattr(ws, "_sheet_format", None)
        if row_breaks is None and col_breaks is None and sheet_format is None:
            continue
        try:
            breaks_dict = ws.to_rust_page_breaks_dict()
            fmt_dict = ws.to_rust_sheet_format_dict()
        except Exception:
            continue
        payload = {
            "row_breaks": breaks_dict.get("row_breaks"),
            "col_breaks": breaks_dict.get("col_breaks"),
            "sheet_format": fmt_dict,
        }
        if all(value is None for value in payload.values()):
            continue
        try:
            patcher.queue_page_breaks_update(ws.title, payload)
        except Exception:
            continue


def flush_pending_autofilters_to_patcher(wb: Any) -> None:
    """Drain pending worksheet autofilter state into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        autofilter = ws._auto_filter  # noqa: SLF001
        has_state = (
            autofilter.ref is not None
            or bool(autofilter.filter_columns)
            or autofilter.sort_state is not None
        )
        if not has_state:
            continue
        try:
            payload = autofilter.to_rust_dict()
        except Exception:
            continue
        try:
            patcher.queue_autofilter(ws.title, payload)
        except Exception:
            continue


def flush_pending_charts_to_patcher(wb: Any) -> None:
    """Drain pending chart additions into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return

    _flush_pending_chart_bytes(wb, patcher)
    _flush_pending_chart_objects(wb, patcher)


def _flush_pending_chart_bytes(wb: Any, patcher: Any) -> None:
    """Drain pre-serialized chart XML payloads into the Rust patcher."""
    pending_bytes = getattr(wb, "_pending_chart_adds", None)
    if pending_bytes:
        for sheet_title, items in pending_bytes.items():
            if not items:
                continue
            for chart_xml, anchor_a1, width_emu, height_emu in items:
                patcher.queue_chart_add(
                    sheet_title,
                    chart_xml,
                    anchor_a1,
                    int(width_emu),
                    int(height_emu),
                )
        pending_bytes.clear()


def _flush_pending_chart_objects(wb: Any, patcher: Any) -> None:
    """Serialize queued chart objects and drain them into the Rust patcher."""
    any_pending = any(ws._pending_charts for ws in wb._sheets.values())  # noqa: SLF001
    if not any_pending:
        return
    try:
        from wolfxl._rust import serialize_chart_dict  # type: ignore[attr-defined]
    except ImportError as exc:  # pragma: no cover - defensive
        raise NotImplementedError(
            "Modify-mode high-level Worksheet.add_chart() requires "
            "Sprint Μ-prime Pod-α′'s serialize_chart_dict PyO3 export. "
            "Build the wolfxl wheel from a branch that includes the "
            "Pod-α′ commits, or fall back to "
            "Workbook.add_chart_modify_mode(sheet, chart_xml_bytes, anchor) "
            "with pre-serialised XML."
        ) from exc

    cm_to_emu = 360_000
    for ws in wb._sheets.values():  # noqa: SLF001
        pending_objs = ws._pending_charts  # noqa: SLF001
        if not pending_objs:
            continue
        for chart in pending_objs:
            chart_dict = chart.to_rust_dict()
            anchor = chart._anchor or "E15"  # noqa: SLF001
            chart_xml = serialize_chart_dict(chart_dict, anchor)
            width_emu, height_emu = _chart_size_emu(chart, cm_to_emu)
            patcher.queue_chart_add(
                ws.title,
                chart_xml,
                anchor,
                width_emu,
                height_emu,
            )
        pending_objs.clear()


def _chart_size_emu(chart: Any, cm_to_emu: int) -> tuple[int, int]:
    """Return chart width and height in EMUs for patcher queueing."""
    return int(chart.width * cm_to_emu), int(chart.height * cm_to_emu)


def flush_pending_slicers_to_patcher(wb: Any) -> None:
    """Drain pending slicer additions into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = getattr(ws, "_pending_slicers", None)
        if not pending:
            continue
        for slicer in pending:
            cache = slicer.cache
            try:
                cache_dict = cache.to_rust_dict()
            except Exception:
                continue
            try:
                slicer_dict = slicer.to_rust_dict()
            except Exception:
                continue
            try:
                patcher.queue_slicer_add(ws.title, cache_dict, slicer_dict)
            except Exception:
                continue
        ws._pending_slicers = []  # noqa: SLF001
    wb._pending_slicer_caches = []  # noqa: SLF001


def flush_pending_pivots_to_patcher(wb: Any) -> None:
    """Drain pending pivot caches and tables into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return

    any_caches = bool(wb._pending_pivot_caches)  # noqa: SLF001
    any_tables = any(
        getattr(ws, "_pending_pivot_tables", None)
        for ws in wb._sheets.values()  # noqa: SLF001
    )
    if not any_caches and not any_tables:
        return

    try:
        from wolfxl._rust import (  # type: ignore[attr-defined]
            serialize_pivot_cache_dict,
            serialize_pivot_records_dict,
            serialize_pivot_table_dict,
        )
    except ImportError as exc:  # pragma: no cover - defensive
        raise NotImplementedError(
            "Modify-mode Workbook.add_pivot_cache() / "
            "Worksheet.add_pivot_table() require Sprint Ν Pod-γ's "
            "serialize_pivot_*_dict PyO3 exports. Build the wolfxl wheel "
            "from a branch that includes the Pod-γ commits."
        ) from exc

    cache_dicts: dict[int, dict[str, Any]] = {}
    for cache in wb._pending_pivot_caches:  # noqa: SLF001
        definition_dict = cache.to_rust_dict()
        records_dict = cache.to_rust_records_dict()
        definition_xml = serialize_pivot_cache_dict(definition_dict)
        records_xml = serialize_pivot_records_dict(definition_dict, records_dict)
        cache_dicts[int(cache._cache_id)] = definition_dict
        allocated = patcher.queue_pivot_cache_add(definition_xml, records_xml)
        if allocated != cache._cache_id:
            raise RuntimeError(
                f"Pivot cache id mismatch: python={cache._cache_id} "
                f"vs patcher={allocated}. This indicates a queue-ordering "
                f"bug in _flush_pending_pivots_to_patcher."
            )
    wb._pending_pivot_caches.clear()  # noqa: SLF001

    for ws in wb._sheets.values():  # noqa: SLF001
        pending = getattr(ws, "_pending_pivot_tables", None)
        if not pending:
            continue
        for pivot_table in pending:
            if pivot_table.cache._cache_id is None:
                raise ValueError(
                    f"PivotTable on sheet {ws.title!r} references a PivotCache "
                    f"that was not registered via Workbook.add_pivot_cache() - "
                    f"register the cache before calling save()."
                )
            if hasattr(pivot_table, "_compute_layout"):
                pivot_table._compute_layout()
            table_dict = pivot_table.to_rust_dict()
            cache_id = int(pivot_table.cache._cache_id)
            cache_dict = cache_dicts.get(cache_id)
            if cache_dict is None:
                cache_dict = pivot_table.cache.to_rust_dict()
            table_xml = serialize_pivot_table_dict(cache_dict, table_dict)
            patcher.queue_pivot_table_add(
                ws.title,
                table_xml,
                cache_id,
            )
        pending.clear()
