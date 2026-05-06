"""Package manifest compatibility helpers."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any
from wolfxl.xml.constants import (
    ARC_APP,
    ARC_CORE,
    ARC_STYLE,
    ARC_THEME,
    CONTYPES_NS,
    CPROPS_TYPE,
    STYLES_TYPE,
    THEME_TYPE,
)
from wolfxl.xml.functions import Element


@dataclass
class FileExtension:
    Extension: str
    ContentType: str

    def to_tree(self) -> Element:
        return Element(
            "Default",
            {"Extension": self.Extension, "ContentType": self.ContentType},
        )


@dataclass
class Override:
    PartName: str
    ContentType: str

    def to_tree(self) -> Element:
        return Element(
            "Override",
            {"PartName": self.PartName, "ContentType": self.ContentType},
        )


DEFAULT_TYPES = [
    FileExtension("rels", "application/vnd.openxmlformats-package.relationships+xml"),
    FileExtension("xml", "application/xml"),
]

DEFAULT_OVERRIDE = [
    Override("/" + ARC_STYLE, STYLES_TYPE),
    Override("/" + ARC_THEME, THEME_TYPE),
    Override("/" + ARC_CORE, "application/vnd.openxmlformats-package.core-properties+xml"),
    Override("/" + ARC_APP, "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
]


class Manifest:
    """Content-types manifest with the small public surface openpyxl exposes."""

    tagname = "Types"
    path = "[Content_Types].xml"

    def __init__(
        self,
        Default: list[FileExtension] | tuple[FileExtension, ...] = (),
        Override: list[Override] | tuple[Override, ...] = (),
    ) -> None:
        self.Default = list(Default) if Default else list(DEFAULT_TYPES)
        self.Override = list(Override) if Override else list(DEFAULT_OVERRIDE)

    @property
    def filenames(self) -> list[str]:
        return [part.PartName for part in self.Override]

    def __contains__(self, content_type: str) -> bool:
        return any(part.ContentType == content_type for part in self.Override)

    def findall(self, content_type: str) -> list[Override]:
        return [part for part in self.Override if part.ContentType == content_type]

    def find(self, content_type: str) -> Override | None:
        matches = self.findall(content_type)
        return matches[0] if matches else None

    def append(self, obj: Any) -> None:
        self.Override.append(Override(PartName=obj.path, ContentType=obj.mime_type))

    def to_tree(self) -> Element:
        root = Element("Types", {"xmlns": CONTYPES_NS})
        seen_defaults: set[tuple[str, str]] = set()
        for default in self.Default:
            key = (default.Extension, default.ContentType)
            if key in seen_defaults:
                continue
            seen_defaults.add(key)
            root.append(default.to_tree())
        seen_overrides: set[tuple[str, str]] = set()
        for override in self.Override:
            key = (override.PartName, override.ContentType)
            if key in seen_overrides:
                continue
            seen_overrides.add(key)
            root.append(override.to_tree())
        return root


__all__ = [
    "CPROPS_TYPE",
    "DEFAULT_OVERRIDE",
    "DEFAULT_TYPES",
    "FileExtension",
    "Manifest",
    "Override",
]
