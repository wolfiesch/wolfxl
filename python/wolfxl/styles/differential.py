"""Shim for ``openpyxl.styles.differential``."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass
class DifferentialStyle:
    """Conditional-format differential style value object."""

    font: Any = None
    numFmt: Any = None  # noqa: N815
    fill: Any = None
    alignment: Any = None
    border: Any = None
    protection: Any = None

    def to_rust_dict(self) -> dict[str, Any]:
        return {
            "font": _style_to_dict(self.font),
            "num_fmt": _style_to_dict(self.numFmt),
            "fill": _style_to_dict(self.fill),
            "alignment": _style_to_dict(self.alignment),
            "border": _style_to_dict(self.border),
            "protection": _style_to_dict(self.protection),
        }


def _style_to_dict(value: Any) -> dict[str, Any] | Any:
    if hasattr(value, "to_rust_dict"):
        return value.to_rust_dict()
    return value

__all__ = ["DifferentialStyle"]
