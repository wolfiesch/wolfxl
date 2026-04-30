"""openpyxl.packaging compatibility.

Re-exports :class:`DocumentProperties` from ``core`` so users can write
``from wolfxl.packaging.core import DocumentProperties`` or the shorter
``from wolfxl.packaging import DocumentProperties``.
"""

from __future__ import annotations

from wolfxl.packaging.core import DocumentProperties
from wolfxl.packaging.custom import (
    BoolProperty,
    CustomPropertyList,
    DateTimeProperty,
    FloatProperty,
    IntProperty,
    LinkProperty,
    StringProperty,
)

__all__ = [
    "BoolProperty",
    "CustomPropertyList",
    "DateTimeProperty",
    "DocumentProperties",
    "FloatProperty",
    "IntProperty",
    "LinkProperty",
    "StringProperty",
]
