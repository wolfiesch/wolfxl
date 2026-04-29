"""Workbook save orchestration helpers.

This module keeps the public :class:`wolfxl.Workbook` methods thin while
preserving the exact writer/patcher flush order used by the save pipeline.
"""

from __future__ import annotations

import os
from typing import Any

from wolfxl._workbook_state import same_existing_path


def save_workbook(
    wb: Any,
    filename: str | os.PathLike[str],
    *,
    password: str | bytes | None = None,
) -> None:
    """Flush workbook state and save it through the active backend."""
    filename = str(filename)
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
        save_write_mode(wb, filename)
    else:
        raise RuntimeError("save requires write or modify mode")


def save_modify_mode(wb: Any, filename: str) -> None:
    """Flush pending modify-mode queues and write through ``XlsxPatcher``."""
    # Workbook-level metadata flushes before per-sheet drains so the patcher
    # composes workbook.xml once, with all pending workbook-scoped edits.
    if wb._properties_dirty:  # noqa: SLF001
        wb._flush_properties_to_patcher()  # noqa: SLF001
    if wb._pending_defined_names:  # noqa: SLF001
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


def save_write_mode(wb: Any, filename: str) -> None:
    """Flush pending write-mode queues and save through ``NativeWorkbook``."""
    wb._flush_workbook_writes()  # noqa: SLF001
    for ws in wb._sheets.values():  # noqa: SLF001
        ws._flush()  # noqa: SLF001
    wb._rust_writer.save(filename)  # noqa: SLF001


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
