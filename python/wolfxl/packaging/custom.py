"""Custom document property containers compatible with openpyxl."""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from typing import Any, Iterator


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


__all__ = [
    "BoolProperty",
    "CustomPropertyList",
    "DateTimeProperty",
    "FloatProperty",
    "IntProperty",
    "LinkProperty",
    "StringProperty",
]
