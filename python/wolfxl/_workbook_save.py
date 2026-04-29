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
    # Modify mode: workbook-level metadata writes don't have a patcher path
    # yet (T1.5 follow-up). Surface the limitation before mutating the file
    # rather than silently dropping the user's edits.
    # RFC-020: properties round-trip (Phase 2.5d in the patcher).
    # Workbook-level, so it flushes before the per-sheet drains.
    if wb._properties_dirty:  # noqa: SLF001
        wb._flush_properties_to_patcher()  # noqa: SLF001
    if wb._pending_defined_names:  # noqa: SLF001
        wb._flush_defined_names_to_patcher()  # noqa: SLF001
    # RFC-058: workbook-level security (workbookProtection + fileSharing).
    # Drained BEFORE sheet flushes so the patcher's Phase 2.5q composes
    # against the source workbook.xml.
    if wb._pending_security_update:  # noqa: SLF001
        wb._flush_security_to_patcher()  # noqa: SLF001
    for ws in wb._sheets.values():  # noqa: SLF001
        ws._flush()  # noqa: SLF001
    # RFC-035: sheet copies must flush BEFORE every per-sheet phase so cloned
    # sheets are visible to downstream drains as if they had always been part
    # of the source workbook.
    wb._flush_pending_sheet_copies_to_patcher()  # noqa: SLF001
    # RFC-022: hyperlinks share the sheet rels graph with future rels-touching
    # writers (RFC-024 tables, RFC-023 comments). Flush them first so DV/CF run
    # afterward against an already-stable rels graph.
    wb._flush_pending_hyperlinks_to_patcher()  # noqa: SLF001
    # RFC-024: tables also touch the rels graph + add new ZIP parts +
    # content-type Overrides. Flush after hyperlinks so the rels graph already
    # carries any external-hyperlink rIds.
    wb._flush_pending_tables_to_patcher()  # noqa: SLF001
    # RFC-023: comments + VML drawings.
    wb._flush_pending_comments_to_patcher()  # noqa: SLF001
    # RFC-025: worksheet-level setters that the patcher accepts.
    wb._flush_pending_data_validations_to_patcher()  # noqa: SLF001
    # RFC-026: conditional formatting sibling blocks.
    wb._flush_pending_conditional_formats_to_patcher()  # noqa: SLF001
    # RFC-030 / RFC-031: structural axis shifts. Drained LAST among core
    # sheet-data phases so it sees earlier per-cell + per-block rewrites.
    wb._flush_pending_axis_shifts_to_patcher()  # noqa: SLF001
    # RFC-034: range moves. Drained after axis shifts so coordinate space is
    # post-shift.
    wb._flush_pending_range_moves_to_patcher()  # noqa: SLF001
    # Sprint Lambda Pod-beta (RFC-045): drain pending images.
    wb._flush_pending_images_to_patcher()  # noqa: SLF001
    # Sprint Mu Pod-gamma (RFC-046): drain pending chart adds.
    wb._flush_pending_charts_to_patcher()  # noqa: SLF001
    # Sprint Nu Pod-gamma (RFC-047 / RFC-048): drain pivot caches/tables.
    wb._flush_pending_pivots_to_patcher()  # noqa: SLF001
    # Sprint Omicron Pod 1A.5 (RFC-055): sheet-setup blocks.
    wb._flush_pending_sheet_setup_to_patcher()  # noqa: SLF001
    # Sprint Pi Pod Pi-alpha (RFC-062): page breaks + sheetFormatPr.
    wb._flush_pending_page_breaks_to_patcher()  # noqa: SLF001
    # Sprint Omicron Pod 3.5 (RFC-061 section 3.1): slicers.
    wb._flush_pending_slicers_to_patcher()  # noqa: SLF001
    # Sprint Omicron Pod 1B (RFC-056): autoFilter dicts.
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
