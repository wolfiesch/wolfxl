"""Workbook metadata, defined-name, and security helper APIs."""

from __future__ import annotations

from typing import Any


def get_properties(wb: Any) -> Any:
    """Return a lazy-loaded workbook document properties object.

    Args:
        wb: Workbook-like object carrying reader and properties cache state.

    Returns:
        A ``DocumentProperties`` instance attached back to ``wb`` so later
        attribute writes mark workbook metadata dirty.
    """
    if wb._properties_cache is not None:  # noqa: SLF001
        return wb._properties_cache  # noqa: SLF001
    from wolfxl.packaging.core import DocumentProperties, _doc_props_from_dict

    if wb._rust_reader is not None:  # noqa: SLF001
        try:
            raw = wb._rust_reader.read_doc_properties()  # noqa: SLF001
        except Exception:
            raw = {}
        props = _doc_props_from_dict(raw)
    else:
        props = DocumentProperties()
    props._attach_workbook(wb)  # noqa: SLF001
    wb._properties_cache = props  # noqa: SLF001
    return props


def set_properties(wb: Any, value: Any) -> None:
    """Replace the workbook document properties object.

    Args:
        wb: Workbook-like object carrying properties cache state.
        value: ``DocumentProperties`` instance to attach and mark dirty.

    Raises:
        TypeError: If ``value`` is not a ``DocumentProperties`` instance.
    """
    from wolfxl.packaging.core import DocumentProperties

    if not isinstance(value, DocumentProperties):
        raise TypeError(
            f"properties must be a DocumentProperties, got {type(value).__name__}"
        )
    value._attach_workbook(wb)  # noqa: SLF001
    wb._properties_cache = value  # noqa: SLF001
    wb._properties_dirty = True  # noqa: SLF001


def get_defined_names(wb: Any) -> Any:
    """Return a lazy-loaded workbook ``DefinedNameDict``.

    Args:
        wb: Workbook-like object carrying reader and defined-name cache state.

    Returns:
        A ``DefinedNameDict`` attached back to ``wb`` so user writes queue for
        save-time flushing.
    """
    if wb._defined_names_cache is not None:  # noqa: SLF001
        return wb._defined_names_cache  # noqa: SLF001
    from wolfxl.workbook import DefinedNameDict
    from wolfxl.workbook.defined_name import DefinedName

    dnd = DefinedNameDict()
    if wb._rust_reader is not None:  # noqa: SLF001
        seen: set[str] = set()
        for sheet_name in wb._sheet_names:  # noqa: SLF001
            try:
                entries = wb._rust_reader.read_named_ranges(sheet_name)  # noqa: SLF001
            except Exception:
                continue
            for entry in entries:
                name = entry["name"]
                if name in seen:
                    continue
                seen.add(name)
                refers_to = entry["refers_to"]
                if refers_to.startswith("="):
                    refers_to = refers_to[1:]
                scope = entry.get("scope", "workbook")
                local_id: int | None = None
                if scope == "sheet":
                    # The reader encodes sheet scope in refers_to; keep the
                    # previous conservative behavior and do not guess the id.
                    local_id = None
                dn = DefinedName(name=name, value=refers_to, localSheetId=local_id)
                dict.__setitem__(dnd, name, dn)
    dnd._wb = wb  # noqa: SLF001
    wb._defined_names_cache = dnd  # noqa: SLF001
    return dnd


def set_security(wb: Any, value: Any) -> None:
    """Set workbook protection metadata.

    Args:
        wb: Workbook-like object carrying security state.
        value: ``WorkbookProtection`` instance, or ``None``.

    Raises:
        TypeError: If ``value`` has the wrong type.
    """
    from wolfxl.workbook.protection import WorkbookProtection

    if value is not None and not isinstance(value, WorkbookProtection):
        raise TypeError(
            "wb.security must be a WorkbookProtection or None, "
            f"got {type(value).__name__}"
        )
    wb._security = value  # noqa: SLF001
    wb._pending_security_update = True  # noqa: SLF001


def set_file_sharing(wb: Any, value: Any) -> None:
    """Set workbook file-sharing metadata.

    Args:
        wb: Workbook-like object carrying file-sharing state.
        value: ``FileSharing`` instance, or ``None``.

    Raises:
        TypeError: If ``value`` has the wrong type.
    """
    from wolfxl.workbook.protection import FileSharing

    if value is not None and not isinstance(value, FileSharing):
        raise TypeError(
            "wb.fileSharing must be a FileSharing or None, "
            f"got {type(value).__name__}"
        )
    wb._file_sharing = value  # noqa: SLF001
    wb._pending_security_update = True  # noqa: SLF001
