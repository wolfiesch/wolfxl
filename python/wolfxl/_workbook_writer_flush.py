"""Write-mode workbook flush helpers for the native Rust writer."""

from __future__ import annotations

from typing import Any


def flush_workbook_writes(wb: Any) -> None:
    """Push workbook metadata, names, and security into the native writer."""
    writer = wb._rust_writer  # noqa: SLF001
    if writer is None:
        return

    if wb._properties_dirty and wb._properties_cache is not None:  # noqa: SLF001
        props = wb._properties_cache  # noqa: SLF001
        payload = {
            "title": props.title,
            "subject": props.subject,
            "creator": props.creator,
            "keywords": props.keywords,
            "description": props.description,
            "lastModifiedBy": props.lastModifiedBy,
            "category": props.category,
            "contentStatus": props.contentStatus,
            "identifier": props.identifier,
            "language": props.language,
            "revision": props.revision,
            "version": props.version,
            "created": props.created.isoformat() if props.created else None,
            "modified": props.modified.isoformat() if props.modified else None,
        }
        writer.set_properties(payload)
        wb._properties_dirty = False  # noqa: SLF001

    if wb._pending_defined_names:  # noqa: SLF001
        primary_sheet = wb._sheet_names[0] if wb._sheet_names else "Sheet"  # noqa: SLF001
        for defined_name in wb._pending_defined_names.values():  # noqa: SLF001
            if defined_name.localSheetId is not None:
                if 0 <= defined_name.localSheetId < len(wb._sheet_names):  # noqa: SLF001
                    sheet_hint = wb._sheet_names[defined_name.localSheetId]  # noqa: SLF001
                else:
                    sheet_hint = primary_sheet
                scope = "sheet"
            else:
                sheet_hint = primary_sheet
                scope = "workbook"
            writer.add_named_range(
                sheet_hint,
                {
                    "name": defined_name.name,
                    "refers_to": defined_name.value,
                    "scope": scope,
                    "comment": defined_name.comment,
                    "local_sheet_id": defined_name.localSheetId,
                    "hidden": defined_name.hidden,
                },
            )
        wb._pending_defined_names.clear()  # noqa: SLF001

    if wb._pending_security_update:  # noqa: SLF001
        payload = wb._build_security_dict()  # noqa: SLF001
        if hasattr(writer, "set_workbook_security"):
            writer.set_workbook_security(payload)
        wb._pending_security_update = False  # noqa: SLF001
