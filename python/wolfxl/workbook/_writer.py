"""Workbook writer compatibility helpers used by openpyxl-shaped callers."""

from __future__ import annotations

from wolfxl.xml.constants import (
    ARC_APP,
    ARC_CORE,
    ARC_CUSTOM,
    ARC_WORKBOOK,
    DOC_NS,
    PKG_REL_NS,
)
from wolfxl.xml.functions import Element, SubElement, tostring


class WorkbookWriter:
    """Small compatibility facade for workbook-level package XML helpers."""

    def __init__(self, wb: object) -> None:
        self.wb = wb

    def write_root_rels(self) -> bytes:
        root = Element("Relationships", {"xmlns": PKG_REL_NS})
        _append_rel(root, 1, f"{DOC_NS}relationships/officeDocument", ARC_WORKBOOK)
        _append_rel(root, 2, f"{PKG_REL_NS}/metadata/core-properties", ARC_CORE)
        _append_rel(root, 3, f"{DOC_NS}relationships/extended-properties", ARC_APP)
        custom_props = getattr(self.wb, "custom_doc_props", None)
        if custom_props is not None and len(custom_props) >= 1:
            _append_rel(root, 4, f"{DOC_NS}relationships/custom-properties", ARC_CUSTOM)
        return tostring(root)


def _append_rel(root: Element, idx: int, rel_type: str, target: str) -> None:
    SubElement(
        root,
        "Relationship",
        {
            "Type": rel_type,
            "Target": target,
            "Id": f"rId{idx}",
        },
    )


__all__ = ["WorkbookWriter"]
