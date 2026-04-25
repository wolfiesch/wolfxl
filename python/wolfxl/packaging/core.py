"""openpyxl.packaging.core.DocumentProperties compatibility.

Matches the openpyxl shape used for ``wb.properties.title`` etc.
Dates are ``datetime`` objects on the Python side; the Rust layer
delivers them as ISO 8601 strings. Missing fields are ``None``.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any


@dataclass
class DocumentProperties:
    """Workbook-level metadata (``docProps/core.xml`` + ``docProps/app.xml``).

    Construction with all defaults produces an "empty" properties object
    — every field is ``None``. This is the shape returned for a fresh
    ``Workbook()`` in write mode (no file to read from yet) and for a
    workbook whose metadata file was missing or malformed.

    In-place attribute assignments (``wb.properties.title = "X"``) flag
    the owning workbook's ``_properties_dirty`` when the
    :meth:`_attach_workbook` helper is called — that lets
    :meth:`Workbook.save` distinguish between "untouched" and "user
    mutated these" across both write and modify modes.
    """

    title: str | None = None
    subject: str | None = None
    creator: str | None = None
    keywords: str | None = None
    description: str | None = None
    lastModifiedBy: str | None = None  # noqa: N815 - openpyxl public API
    category: str | None = None
    contentStatus: str | None = None  # noqa: N815
    identifier: str | None = None
    language: str | None = None
    revision: str | None = None
    version: str | None = None
    created: datetime | None = None
    modified: datetime | None = None

    # Dirty-tracking support. Set via _attach_workbook after construction;
    # the dataclass's own __init__ uses __setattr__, so we only enable
    # the tracking hook once _wb is present.
    def __setattr__(self, name: str, value: Any) -> None:
        object.__setattr__(self, name, value)
        wb = self.__dict__.get("_wb")
        if wb is not None and name != "_wb":
            wb._properties_dirty = True  # noqa: SLF001

    def _attach_workbook(self, wb: Any) -> None:
        """Link this properties object to its owning Workbook.

        After this call, every subsequent attribute assignment flips
        ``wb._properties_dirty = True`` — transparent to the user.
        """
        object.__setattr__(self, "_wb", wb)


def _doc_props_from_dict(raw: dict[str, Any] | None) -> DocumentProperties:
    """Build a ``DocumentProperties`` from the Rust reader's dict output.

    The Rust side emits every field as a string (or omits it); we parse
    the two datetime fields here. A malformed ``created`` / ``modified``
    string collapses to ``None`` rather than raising, so a corrupt
    sidecar can't break opening the workbook.
    """
    raw = raw or {}

    def _parse_dt(value: Any) -> datetime | None:
        if value is None:
            return None
        if isinstance(value, datetime):
            return value
        if not isinstance(value, str) or not value:
            return None
        # OOXML uses ISO 8601 with a trailing Z. Python 3.11+'s
        # fromisoformat handles Z; fall back on a manual strip for 3.10.
        try:
            return datetime.fromisoformat(value)
        except ValueError:
            if value.endswith("Z"):
                try:
                    return datetime.fromisoformat(value[:-1])
                except ValueError:
                    return None
            return None

    return DocumentProperties(
        title=raw.get("title"),
        subject=raw.get("subject"),
        creator=raw.get("creator"),
        keywords=raw.get("keywords"),
        description=raw.get("description"),
        lastModifiedBy=raw.get("lastModifiedBy"),
        category=raw.get("category"),
        contentStatus=raw.get("contentStatus"),
        identifier=raw.get("identifier"),
        language=raw.get("language"),
        revision=raw.get("revision"),
        version=raw.get("version"),
        created=_parse_dt(raw.get("created")),
        modified=_parse_dt(raw.get("modified")),
    )


__all__ = ["DocumentProperties", "_doc_props_from_dict"]
