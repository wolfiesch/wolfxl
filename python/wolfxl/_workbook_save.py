"""Workbook save orchestration helpers.

This module keeps the public :class:`wolfxl.Workbook` methods thin while
preserving the exact writer/patcher flush order used by the save pipeline.
"""

from __future__ import annotations

import os
from typing import Any, BinaryIO

from wolfxl._workbook_state import same_existing_path


def normalize_openpyxl_package_shape(wb: Any, filename: str) -> None:
    """Apply source-backed openpyxl package-shape cleanup when relevant."""
    from wolfxl._openpyxl_package_shape import normalize_openpyxl_package_shape as _normalize

    keep_vba = bool(getattr(wb, "_keep_vba", False))
    if keep_vba:
        _normalize(filename, keep_vba=True)


def save_workbook(
    wb: Any,
    filename: str | os.PathLike[str] | BinaryIO,
    *,
    password: str | bytes | None = None,
) -> None:
    """Flush workbook state and save it through the active backend."""
    if hasattr(filename, "write") and not isinstance(filename, (str, bytes, os.PathLike)):
        save_workbook_to_fileobj(wb, filename, password=password)
        return
    filename = str(filename)
    # G20: write-only mode is consumed-on-save. A second save raises
    # WorkbookAlreadySaved (matches openpyxl's `_write_only.py`).
    # Eager-mode workbooks remain re-savable.
    if getattr(wb, "_saved", False) and getattr(wb, "_write_only", False):
        from wolfxl.utils.exceptions import WorkbookAlreadySaved

        raise WorkbookAlreadySaved(
            "Workbook(write_only=True) is consumed-on-save; "
            "open a new workbook to write again"
        )

    if password is not None:
        # Validate password early so we don't write a plaintext tempfile that
        # we'd then have to throw away.
        from wolfxl._encryption import _coerce_password

        _coerce_password(password)  # raises ValueError on empty
        save_encrypted(wb, filename, password)
        return

    if wb._rust_patcher is not None:  # noqa: SLF001
        save_modify_mode(wb, filename)
    elif wb._rust_writer is not None:  # noqa: SLF001
        if getattr(wb, "_write_only", False):
            save_write_only_mode(wb, filename)
        else:
            save_write_mode(wb, filename)
    elif getattr(wb, "_rust_reader", None) is not None and getattr(wb, "_source_path", None):
        save_read_mode(wb, filename)
    else:
        raise RuntimeError("save requires write or modify mode")
    # Mark consumed AFTER save succeeds so a write failure leaves the
    # workbook in a re-tryable state for the eager path. Write-only
    # mode flips this once and never re-saves.
    wb._saved = True  # noqa: SLF001
    # Mark every WriteOnlyWorksheet closed so subsequent appends raise
    # WorkbookAlreadySaved cleanly.
    if getattr(wb, "_write_only", False):
        for ws in wb._sheets.values():  # noqa: SLF001
            close = getattr(ws, "close", None)
            if close is not None:
                close()


def save_workbook_to_fileobj(
    wb: Any,
    fileobj: BinaryIO,
    *,
    password: str | bytes | None = None,
) -> None:
    """Save to a binary file-like object using the path-oriented backends."""
    import tempfile

    tmp = tempfile.NamedTemporaryFile(prefix="wolfxl-save-", suffix=".xlsx", delete=False)
    tmp_path = tmp.name
    tmp.close()
    try:
        save_workbook(wb, tmp_path, password=password)
        with open(tmp_path, "rb") as src:
            data = src.read()
        try:
            fileobj.seek(0)
            fileobj.truncate()
        except Exception:
            pass
        fileobj.write(data)
        try:
            fileobj.flush()
            fileobj.seek(0)
        except Exception:
            pass
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def save_read_mode(wb: Any, filename: str) -> None:
    """Save an unmodified path-backed read workbook by copying the source package."""
    import shutil

    source_path = getattr(wb, "_source_path", None)
    if source_path is None:
        raise RuntimeError("save requires write or modify mode")
    if _read_mode_has_pending_changes(wb):
        if getattr(wb, "_read_only", False):
            raise RuntimeError(
                "save() on a read_only=True workbook would discard pending changes; "
                "reopen with modify=True before editing"
            )
        _promote_read_mode_to_patcher(wb, source_path)
        save_modify_mode(wb, filename)
        return
    shutil.copyfile(source_path, filename)
    normalize_openpyxl_package_shape(wb, filename)


def _promote_read_mode_to_patcher(wb: Any, source_path: str) -> None:
    from wolfxl import _rust

    wb._rust_patcher = _rust.XlsxPatcher.open(source_path, False)


def _read_mode_has_pending_changes(wb: Any) -> bool:
    workbook_pending_attrs = (
        "_chartsheets_dirty",
        "_properties_dirty",
        "_pending_defined_names",
        "_pending_security_update",
        "_pending_axis_shifts",
        "_pending_range_moves",
        "_pending_sheet_copies",
        "_pending_chart_adds",
        "_pending_source_chart_ops",
        "_pending_pivot_caches",
        "_pending_slicer_caches",
        "_strip_external_links_on_save",
    )
    if any(bool(getattr(wb, attr, None)) for attr in workbook_pending_attrs):
        return True
    links = getattr(wb, "_external_links", None)
    if links is not None and getattr(links, "dirty", False):
        return True

    worksheet_pending_attrs = (
        "_dirty",
        "_append_buffer",
        "_bulk_writes",
        "_pending_comments",
        "_pending_threaded_comments",
        "_pending_hyperlinks",
        "_pending_tables",
        "_pending_data_validations",
        "_pending_conditional_formats",
        "_pending_rich_text",
        "_pending_array_formulas",
        "_pending_images",
        "_pending_charts",
        "_pending_chart_deletions",
        "_pending_pivot_tables",
        "_pending_slicers",
        "_print_titles_dirty",
    )
    from wolfxl._worksheet_media import has_pending_image_deletions

    for ws in getattr(wb, "_sheets", {}).values():
        if any(bool(getattr(ws, attr, None)) for attr in worksheet_pending_attrs):
            return True
        if getattr(ws, "_header_footer", None) is not None:
            return True
        if has_pending_image_deletions(ws):
            return True
        for handle in getattr(ws, "_pivot_handles_cache", None) or []:
            if getattr(handle, "_dirty", False) or getattr(handle, "_layout_dirty", False):
                return True
    return False


def save_write_only_mode(wb: Any, filename: str) -> None:
    """Save path for ``Workbook(write_only=True)`` (G20).

    The streaming temp files have been accumulating row XML throughout
    the session. This flushes each per-sheet ``BufWriter`` then drives
    the standard ``emit_xlsx`` pipeline — the only difference from
    eager save is that ``sheet_xml::emit`` splices the temp file into
    the ``<sheetData>`` slot instead of walking ``Worksheet.rows``.
    """
    # Workbook-level metadata flush (defined names, properties) still
    # composes through the writer-side path; sheet-level flush is a
    # no-op for write-only sheets because the temp files have been
    # written incrementally.
    wb._flush_workbook_writes()  # noqa: SLF001
    wb._rust_writer.finalize_streaming_sheets()  # noqa: SLF001
    wb._rust_writer.save(filename)  # noqa: SLF001
    flush_chartsheets_authoring(wb, filename)


def save_modify_mode(wb: Any, filename: str) -> None:
    """Flush pending modify-mode queues and write through ``XlsxPatcher``."""
    if not _modify_mode_has_pending_changes(wb):
        if same_existing_path(filename, wb._source_path):  # noqa: SLF001
            wb._rust_patcher.save_in_place()  # noqa: SLF001
        else:
            wb._rust_patcher.save(filename)  # noqa: SLF001
        return

    # Workbook-level metadata flushes before per-sheet drains so the patcher
    # composes workbook.xml once, with all pending workbook-scoped edits.
    if wb._properties_dirty:  # noqa: SLF001
        wb._flush_properties_to_patcher()  # noqa: SLF001
    wb._flush_defined_names_to_patcher()  # noqa: SLF001
    # Workbook-level security also targets workbook.xml and must precede
    # sheet-scoped patch queues.
    if wb._pending_security_update:  # noqa: SLF001
        wb._flush_security_to_patcher()  # noqa: SLF001
    for ws in wb._sheets.values():  # noqa: SLF001
        ws._flush()  # noqa: SLF001
    # Sheet copies must flush before every per-sheet phase so cloned sheets are
    # visible to downstream drains as if they had always been part of the
    # source workbook.
    wb._flush_pending_sheet_copies_to_patcher()  # noqa: SLF001
    # Hyperlinks share the sheet rels graph with tables and comments. Flush
    # them first so validations/conditional formats run afterward against an
    # already-stable rels graph.
    wb._flush_pending_hyperlinks_to_patcher()  # noqa: SLF001
    # Tables also touch the rels graph, add ZIP parts, and add content-type
    # overrides. Flush after hyperlinks so external-hyperlink rIds are stable.
    wb._flush_pending_tables_to_patcher()  # noqa: SLF001
    # Threaded comments + person list (RFC-068 G08). Drained BEFORE the
    # legacy comments flush so the Rust patcher's threaded-comments phase
    # (which synthesizes `tc={topId}` placeholders) can pre-populate
    # queued_comments before apply_comments_phase runs.
    wb._flush_pending_threaded_comments_to_patcher()  # noqa: SLF001
    wb._flush_pending_persons_to_patcher()  # noqa: SLF001
    # Comments and VML drawings.
    wb._flush_pending_comments_to_patcher()  # noqa: SLF001
    # Worksheet-level data validation setters.
    wb._flush_pending_data_validations_to_patcher()  # noqa: SLF001
    # Conditional formatting sibling blocks.
    wb._flush_pending_conditional_formats_to_patcher()  # noqa: SLF001
    # Structural axis shifts. Drained last among core sheet-data phases so they
    # see earlier per-cell and per-block rewrites.
    wb._flush_pending_axis_shifts_to_patcher()  # noqa: SLF001
    # Range moves. Drained after axis shifts so coordinate space is post-shift.
    wb._flush_pending_range_moves_to_patcher()  # noqa: SLF001
    # Images.
    wb._flush_pending_images_to_patcher()  # noqa: SLF001
    # Chart additions.
    wb._flush_pending_charts_to_patcher()  # noqa: SLF001
    # Pivot caches and tables.
    wb._flush_pending_pivots_to_patcher()  # noqa: SLF001
    # G17 / RFC-070 — pivot source-range edits on existing pivots.
    # Sequenced AFTER the adds flush so a session that both adds and
    # edits goes through the patcher's pivot phases in source order.
    wb._flush_pending_pivot_source_edits_to_patcher()  # noqa: SLF001
    # Sheet-setup blocks.
    wb._flush_pending_sheet_setup_to_patcher()  # noqa: SLF001
    # Page breaks and sheetFormatPr.
    wb._flush_pending_page_breaks_to_patcher()  # noqa: SLF001
    # Slicers.
    wb._flush_pending_slicers_to_patcher()  # noqa: SLF001
    # AutoFilter dicts.
    wb._flush_pending_autofilters_to_patcher()  # noqa: SLF001

    if same_existing_path(filename, wb._source_path):  # noqa: SLF001
        wb._rust_patcher.save_in_place()  # noqa: SLF001
    else:
        wb._rust_patcher.save(filename)  # noqa: SLF001
    flush_pivot_layout_authoring(wb, filename)
    flush_external_links_authoring(wb, filename)
    flush_source_chart_authoring(wb, filename)
    flush_chartsheets_authoring(wb, filename)
    normalize_openpyxl_package_shape(wb, filename)


def _modify_mode_has_pending_changes(wb: Any) -> bool:
    if _read_mode_has_pending_changes(wb):
        return True
    patcher = getattr(wb, "_rust_patcher", None)
    if patcher is not None:
        has_pending = getattr(patcher, "_has_pending_save_work", None)
        if has_pending is None:
            return True
        if bool(has_pending()):
            return True
    if _has_pending_source_chart_authoring(wb):
        return True
    return False


def _has_pending_source_chart_authoring(wb: Any) -> bool:
    if getattr(wb, "_pending_source_chart_ops", None):
        return True
    for ws in getattr(wb, "_sheets", {}).values():
        for chart in getattr(ws, "_charts_cache", None) or []:
            meta = getattr(chart, "_wolfxl_source_chart", None)
            if not meta:
                continue
            original_title = getattr(chart, "_wolfxl_source_title", None)
            if _source_chart_title_signature(chart) != original_title:
                return True
    return False


def save_write_mode(wb: Any, filename: str) -> None:
    """Flush pending write-mode queues and save through ``NativeWorkbook``.

    Write-mode pivot construction (G17 / RFC-070 §8.7 reach-extension):
    when pending pivot caches or pivot tables are queued, a two-phase
    save runs — the native writer emits cell data + sheet structure
    to a tempfile, then the same tempfile is reopened in modify mode
    and the patcher's ``apply_pivot_adds_phase`` stamps in the pivot
    parts. Final bytes copy onto ``filename``.
    """
    has_pending_pivots = bool(getattr(wb, "_pending_pivot_caches", None)) or any(
        getattr(ws, "_pending_pivot_tables", None) for ws in wb._sheets.values()  # noqa: SLF001
    )
    if not has_pending_pivots:
        wb._flush_workbook_writes()  # noqa: SLF001
        for ws in wb._sheets.values():  # noqa: SLF001
            ws._flush()  # noqa: SLF001
        wb._rust_writer.save(filename)  # noqa: SLF001
        flush_external_links_authoring(wb, filename)
        flush_chartsheets_authoring(wb, filename)
        return

    _save_write_mode_with_pivots(wb, filename)
    flush_external_links_authoring(wb, filename)
    flush_chartsheets_authoring(wb, filename)


def flush_external_links_authoring(wb: Any, filename: str) -> None:
    links = getattr(wb, "_external_links_cache", None)
    strip_links = bool(getattr(wb, "_strip_external_links_on_save", False))
    if links is None and strip_links:
        links = wb._external_links  # noqa: SLF001
    if links is None or (not getattr(links, "dirty", False) and not strip_links):
        return
    from wolfxl import _external_links as _el

    if strip_links:
        links._mark_dirty()  # noqa: SLF001
    _el.apply_authoring_to_xlsx(filename, links)
    wb._strip_external_links_on_save = False  # noqa: SLF001


def flush_chartsheets_authoring(wb: Any, filename: str) -> None:
    chartsheets = getattr(wb, "_chartsheets", None)
    if not chartsheets or not any(
        not getattr(cs, "_source_chartsheet", False) for cs in chartsheets.values()
    ):
        return
    from wolfxl import _chartsheets

    _chartsheets.apply_chartsheets_to_xlsx(filename, wb)


def flush_source_chart_authoring(wb: Any, filename: str) -> None:
    ops = list(getattr(wb, "_pending_source_chart_ops", []))
    touched = {
        op.get("meta", {}).get("chart_path")
        for op in ops
        if isinstance(op.get("meta"), dict)
    }
    for ws in getattr(wb, "_sheets", {}).values():
        for chart in getattr(ws, "_charts_cache", None) or []:
            meta = getattr(chart, "_wolfxl_source_chart", None)
            if not meta or meta.get("chart_path") in touched:
                continue
            original_title = getattr(chart, "_wolfxl_source_title", None)
            current_title = _source_chart_title_signature(chart)
            if current_title != original_title:
                ops.append({"op": "title", "meta": meta, "chart": chart})
                touched.add(meta.get("chart_path"))

    if not ops:
        return

    from wolfxl import _source_charts

    materialized = []
    for op in ops:
        if op["op"] in {"replace", "title"}:
            op = dict(op)
            op["chart_xml"] = _serialize_chart_xml(op["chart"])
        materialized.append(op)
    _source_charts.apply_source_chart_authoring_to_xlsx(filename, materialized)
    wb._pending_source_chart_ops.clear()  # noqa: SLF001


def _serialize_chart_xml(chart: Any) -> bytes:
    from wolfxl._rust import serialize_chart_dict  # type: ignore[attr-defined]

    original_anchor = getattr(chart, "_anchor", None)
    chart._anchor = "E15"  # noqa: SLF001
    try:
        return serialize_chart_dict(chart.to_rust_dict(), "E15")
    finally:
        chart._anchor = original_anchor  # noqa: SLF001


def _source_chart_title_signature(chart: Any) -> Any:
    title = getattr(chart, "title", None)
    if title is None:
        return None
    to_dict = getattr(title, "to_dict", None)
    if to_dict is not None:
        try:
            return to_dict()
        except Exception:
            pass
    return str(title)


def flush_pivot_layout_authoring(wb: Any, filename: str) -> None:
    from wolfxl.pivot._handle import apply_pivot_layout_authoring_to_xlsx

    apply_pivot_layout_authoring_to_xlsx(filename, wb)


def _save_write_mode_with_pivots(wb: Any, filename: str) -> None:
    """Two-phase save: writer → tempfile → patcher → final destination.

    This unblocks the openpyxl-shaped pattern used by the G17 oracle
    probe:

    .. code-block:: python

        wb = wolfxl.Workbook()
        wb.add_pivot_cache(cache)
        ws.add_pivot_table(pt, "A1")
        wb.save(path)

    by emitting cell data through the native writer, then surgically
    grafting the pivot parts on via the existing modify-mode patcher
    pipeline. No new pivot emit code is introduced.
    """
    import tempfile

    # Stash the pending pivot queues — the writer-side flush would
    # otherwise observe them and (in some future hardening) try to
    # double-emit. Re-attach them onto the modify-mode workbook below.
    pending_caches = list(getattr(wb, "_pending_pivot_caches", []))
    pending_tables_by_sheet: dict[str, list[Any]] = {}
    for ws in wb._sheets.values():  # noqa: SLF001
        title = ws.title
        pending_tables_by_sheet[title] = list(
            getattr(ws, "_pending_pivot_tables", [])
        )
        ws._pending_pivot_tables = []  # noqa: SLF001
    wb._pending_pivot_caches = []  # noqa: SLF001
    # Reset cache id allocator and clear cache._cache_id stamps so the
    # modify-mode add_pivot_cache call re-allocates cleanly.
    wb._next_pivot_cache_id = 0  # noqa: SLF001
    for cache in pending_caches:
        cache._cache_id = None  # noqa: SLF001

    # Stage 1 — writer emits the workbook (cells, sheets, formats, etc.).
    wb._flush_workbook_writes()  # noqa: SLF001
    for ws in wb._sheets.values():  # noqa: SLF001
        ws._flush()  # noqa: SLF001

    tmp_fd, tmp_name = tempfile.mkstemp(prefix=".wolfxl-pivot-", suffix=".xlsx")
    os.close(tmp_fd)
    try:
        wb._rust_writer.save(tmp_name)  # noqa: SLF001

        # Stage 2 — reopen the tempfile in modify mode and replay the
        # pivot adds.
        import wolfxl as _wolfxl

        modify_wb = _wolfxl.load_workbook(tmp_name, modify=True)
        for cache in pending_caches:
            modify_wb.add_pivot_cache(cache)
        for sheet_title, pts in pending_tables_by_sheet.items():
            if not pts:
                continue
            target = modify_wb._sheets.get(sheet_title)  # noqa: SLF001
            if target is None:
                continue
            for pt in pts:
                # add_pivot_table on Worksheet expects a single arg in
                # write/modify mode; the anchor is captured on the
                # PivotTable's `location` attr at construction time.
                target.add_pivot_table(pt)
        modify_wb.save(filename)
    finally:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass


def save_encrypted(wb: Any, filename: str, password: str | bytes) -> None:
    """Save plaintext to a tempfile, then encrypt it atomically."""
    import tempfile

    from wolfxl._encryption import encrypt_xlsx_to_path

    tmp_fd, tmp_name = tempfile.mkstemp(
        prefix=".wolfxl-plain-",
        suffix=".xlsx",
    )
    os.close(tmp_fd)
    try:
        # Re-enter the normal plaintext save path so writer and patcher modes
        # exercise the same pipeline before encryption.
        save_workbook(wb, tmp_name)
        with open(tmp_name, "rb") as fp:
            plaintext_bytes = fp.read()
        encrypt_xlsx_to_path(plaintext_bytes, password, filename)
    finally:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass
