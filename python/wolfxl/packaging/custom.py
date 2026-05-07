"""Custom document property containers compatible with openpyxl."""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from typing import Any, Iterator

from wolfxl.xml import LXML
from wolfxl.xml.constants import CPROPS_FMTID, CUSTPROPS_NS, VTYPES_NS
from wolfxl.xml.functions import Element, SubElement


@dataclass
class _TypedProperty:
    """Base class for a named custom document property."""

    name: str
    value: Any


@dataclass
class StringProperty(_TypedProperty):
    """String custom document property."""

    value: str | None


@dataclass
class IntProperty(_TypedProperty):
    """Integer custom document property."""

    value: int


@dataclass
class FloatProperty(_TypedProperty):
    """Floating-point custom document property."""

    value: float


@dataclass
class BoolProperty(_TypedProperty):
    """Boolean custom document property."""

    value: bool


@dataclass
class DateTimeProperty(_TypedProperty):
    """Datetime custom document property."""

    value: dt.datetime


@dataclass
class LinkProperty(_TypedProperty):
    """Linked custom document property."""

    value: str


class CustomPropertyList:
    """List-like container for workbook custom document properties."""

    def __init__(self) -> None:
        self.props: list[_TypedProperty] = []

    @property
    def names(self) -> list[str]:
        """Return custom property names in document order."""
        return [prop.name for prop in self.props]

    def append(self, prop: _TypedProperty) -> None:
        """Append a custom property, rejecting duplicate names."""
        if prop.name in self.names:
            raise ValueError(f"Property with name {prop.name} already exists")
        self.props.append(prop)

    def __len__(self) -> int:
        return len(self.props)

    def __iter__(self) -> Iterator[_TypedProperty]:
        return iter(self.props)

    def __getitem__(self, name: str) -> _TypedProperty:
        for prop in self.props:
            if prop.name == name:
                return prop
        raise KeyError(f"Property with name {name} not found")

    def __delitem__(self, name: str) -> None:
        for index, prop in enumerate(self.props):
            if prop.name == name:
                self.props.pop(index)
                return
        raise KeyError(f"Property with name {name} not found")

    def __repr__(self) -> str:
        return f"{self.__class__.__name__} containing {self.props}"

    def to_tree(self) -> Element:
        """Serialize custom properties to the OOXML custom-properties part."""
        if LXML:
            root = Element("Properties", nsmap={None: CUSTPROPS_NS, "vt": VTYPES_NS})
        else:
            root = Element("Properties", {"xmlns": CUSTPROPS_NS, "xmlns:vt": VTYPES_NS})
        for pid, prop in enumerate(self.props, 2):
            node = SubElement(
                root,
                "property",
                {
                    "fmtid": CPROPS_FMTID,
                    "pid": str(pid),
                    "name": prop.name,
                },
            )
            tag, text = _value_node(prop)
            child = SubElement(node, f"{{{VTYPES_NS}}}{tag}")
            child.text = text
        return root


def _value_node(prop: _TypedProperty) -> tuple[str, str]:
    class_name = prop.__class__.__name__
    if class_name == "IntProperty":
        return "i4", str(prop.value)
    if class_name == "FloatProperty":
        return "r8", str(prop.value)
    if class_name == "BoolProperty":
        return "bool", "true" if prop.value else "false"
    if class_name == "DateTimeProperty":
        value = prop.value
        if isinstance(value, dt.datetime):
            text = value.replace(microsecond=0).isoformat()
            if value.tzinfo is None:
                text += "Z"
        else:
            text = str(value)
        return "filetime", text
    return "lpwstr", "" if prop.value is None else str(prop.value)


__all__ = [
    "BoolProperty",
    "CustomPropertyList",
    "DateTimeProperty",
    "FloatProperty",
    "IntProperty",
    "LinkProperty",
    "StringProperty",
]
