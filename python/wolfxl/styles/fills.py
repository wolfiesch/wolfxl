"""``openpyxl.styles.fills`` — fill value types."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from wolfxl._styles import PatternFill

# Pattern-type vocabulary mirrored from openpyxl for callers that
# introspect against it (e.g. validation pre-processors).
fills = (
    "none",
    "solid",
    "darkGray",
    "mediumGray",
    "lightGray",
    "gray125",
    "gray0625",
    "darkHorizontal",
    "darkVertical",
    "darkDown",
    "darkUp",
    "darkGrid",
    "darkTrellis",
    "lightHorizontal",
    "lightVertical",
    "lightDown",
    "lightUp",
    "lightGrid",
    "lightTrellis",
)


@dataclass
class Fill:
    """Base fill container.

    openpyxl exposes ``Fill`` as the abstract base for pattern and gradient
    fills. WolfXL treats it as a passive value object so direct construction
    used by migration code no longer raises.
    """

    tagname: str | None = None

    def to_rust_dict(self) -> dict[str, Any]:
        return {"tagname": self.tagname}


@dataclass
class GradientFill(Fill):
    """Gradient fill value object mirroring openpyxl's public shape."""

    type: str = "linear"  # noqa: A003
    degree: float = 0.0
    left: float = 0.0
    right: float = 0.0
    top: float = 0.0
    bottom: float = 0.0
    stop: list[Any] = field(default_factory=list)

    def __init__(
        self,
        type: str = "linear",  # noqa: A002
        degree: float = 0.0,
        left: float = 0.0,
        right: float = 0.0,
        top: float = 0.0,
        bottom: float = 0.0,
        stop: list[Any] | tuple[Any, ...] | None = None,
        *,
        fill_type: str | None = None,
    ) -> None:
        self.tagname = "gradientFill"
        self.type = fill_type if fill_type is not None else type
        self.degree = degree
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.stop = list(stop or [])

    @property
    def fill_type(self) -> str:
        return self.type

    def to_rust_dict(self) -> dict[str, Any]:
        return {
            "type": self.type,
            "degree": self.degree,
            "left": self.left,
            "right": self.right,
            "top": self.top,
            "bottom": self.bottom,
            "stop": self.stop,
        }


__all__ = ["Fill", "GradientFill", "PatternFill", "fills"]
