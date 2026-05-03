"""Write-mode workbook flush helpers for the native Rust writer."""

from __future__ import annotations

from typing import Any


def flush_workbook_writes(wb: Any) -> None:
    """Push workbook metadata, names, and security into the native writer."""
    writer = wb._rust_writer  # noqa: SLF001
    if writer is None:
        return

    _flush_properties(wb, writer)
    _flush_defined_names(wb, writer)
    _flush_named_styles(wb, writer)
    _flush_security(wb, writer)
    _flush_persons(wb, writer)


def _flush_properties(wb: Any, writer: Any) -> None:
    """Push dirty document properties to the native writer."""
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


def _flush_defined_names(wb: Any, writer: Any) -> None:
    """Push pending defined names to the native writer."""
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


def _flush_named_styles(wb: Any, writer: Any) -> None:
    """Push workbook named-style registrations to the native writer."""
    if not hasattr(writer, "add_named_style"):
        return
    registry = getattr(wb, "_named_styles_registry", None)  # noqa: SLF001
    if registry is None:
        return
    for style in registry.user_styles():
        writer.add_named_style(style.name)


def _flush_security(wb: Any, writer: Any) -> None:
    """Push pending workbook security metadata to the native writer."""
    if wb._pending_security_update:  # noqa: SLF001
        payload = wb._build_security_dict()  # noqa: SLF001
        if hasattr(writer, "set_workbook_security"):
            writer.set_workbook_security(payload)
        wb._pending_security_update = False  # noqa: SLF001


def _flush_persons(wb: Any, writer: Any) -> None:
    """Push the workbook-scope threaded-comment person registry (RFC-068).

    The Rust ``add_person`` entry point is idempotent on Person ``id`` so
    repeated flush calls produce a stable personList. The registry is only
    seeded lazily; if the user never accessed ``wb.persons`` there is
    nothing to flush.
    """
    if not hasattr(writer, "add_person"):
        return
    registry = getattr(wb, "_persons_registry", None)  # noqa: SLF001
    if registry is None or len(registry) == 0:
        return
    for person in registry:
        if person.id is None:
            # PersonRegistry should always allocate a GUID, but guard the
            # contract so a misbehaving caller can't poison the save path.
            continue
        writer.add_person(
            {
                "id": person.id,
                "name": person.name,
                "user_id": person.user_id,
                "provider_id": person.provider_id,
            }
        )
