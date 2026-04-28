"""Style dataclasses matching openpyxl's public names.

These are lightweight, frozen value objects. They mirror openpyxl's Font,
PatternFill, Border, Side, Alignment, and Color classes so that code written
for openpyxl can be ported with minimal changes.
"""

from __future__ import annotations

from dataclasses import dataclass, field

# openpyxl's INDEXED_COLORS - 64 legacy palette entries (index 0..63).
# Copied verbatim from openpyxl 3.1.x (MIT-licensed) so that reading a
# workbook that references indexed colors doesn't crash on Color
# construction. Indexes 64 and 65 are the "system foreground" and
# "system background" slots; openpyxl maps them to the same defaults.
COLOR_INDEX: tuple[str, ...] = (
    "00000000", "00FFFFFF", "00FF0000", "0000FF00", "000000FF",
    "00FFFF00", "00FF00FF", "0000FFFF", "00000000", "00FFFFFF",
    "00FF0000", "0000FF00", "000000FF", "00FFFF00", "00FF00FF",
    "0000FFFF", "00800000", "00008000", "00000080", "00808000",
    "00800080", "00008080", "00C0C0C0", "00808080", "009999FF",
    "00993366", "00FFFFCC", "00CCFFFF", "00660066", "00FF8080",
    "000066CC", "00CCCCFF", "00000080", "00FF00FF", "00FFFF00",
    "0000FFFF", "00800080", "00800000", "00008080", "000000FF",
    "0000CCFF", "00CCFFFF", "00CCFFCC", "00FFFF99", "0099CCFF",
    "00FF99CC", "00CC99FF", "00FFCC99", "003366FF", "0033CCCC",
    "0099CC00", "00FFCC00", "00FF9900", "00FF6600", "00666699",
    "00969696", "00003366", "00339966", "00003300", "00333300",
    "00993300", "00993366", "00333399", "00333333",
    "00000000",  # 64 - system foreground
    "00FFFFFF",  # 65 - system background
)


@dataclass(frozen=True, init=False)
class Color:
    """An Excel color.

    A color is identified by exactly one of three strategies: direct RGB,
    theme index (with optional tint), or legacy indexed palette position.
    Which strategy is in effect is recorded in ``type`` and mirrors
    openpyxl's ``Color`` class shape so callers inspecting ``color.type``
    see the same vocabulary.

    The default `Color()` is opaque black RGB, matching openpyxl.
    """

    rgb: str | None = None
    theme: int | None = None
    indexed: int | None = None
    tint: float = 0.0
    type: str = "rgb"

    def __init__(
        self,
        rgb: str | None = None,
        *,
        theme: int | None = None,
        indexed: int | None = None,
        tint: float = 0.0,
        type: str | None = None,  # noqa: A002 — openpyxl uses this name
    ) -> None:
        # Resolve which constructor form was used. If theme/indexed/rgb is
        # set, that wins; otherwise default to opaque black RGB to match
        # openpyxl's no-arg default.
        if theme is not None:
            resolved_type = type or "theme"
            resolved_rgb: str | None = None
        elif indexed is not None:
            resolved_type = type or "indexed"
            resolved_rgb = None
        elif rgb is not None:
            resolved_type = type or "rgb"
            resolved_rgb = rgb
        else:
            resolved_type = type or "rgb"
            resolved_rgb = "00000000"
        object.__setattr__(self, "rgb", resolved_rgb)
        object.__setattr__(self, "theme", theme)
        object.__setattr__(self, "indexed", indexed)
        object.__setattr__(self, "tint", tint)
        object.__setattr__(self, "type", resolved_type)

    def to_hex(self) -> str:
        """Return '#RRGGBB' (strips the alpha channel).

        For theme-only colors (no rgb, no indexed), falls back to black
        since wolfxl doesn't resolve the theme XML at this layer. For
        indexed colors, looks up the legacy palette.
        """
        if self.rgb is not None:
            raw = self.rgb.lstrip("#")
            if len(raw) == 8:
                return f"#{raw[2:]}"
            return f"#{raw}"
        if self.indexed is not None:
            idx = self.indexed
            if 0 <= idx < len(COLOR_INDEX):
                raw = COLOR_INDEX[idx]
                return f"#{raw[2:]}"
            return "#000000"
        # Theme-only fallback; callers that need the resolved RGB should
        # consult the theme XML directly.
        return "#000000"

    @classmethod
    def from_hex(cls, hex_str: str) -> Color:
        """Create from '#RRGGBB' or 'RRGGBB' (assumes FF alpha)."""
        raw = hex_str.lstrip("#")
        if len(raw) == 6:
            return cls(rgb=f"FF{raw.upper()}")
        return cls(rgb=raw.upper())


@dataclass(frozen=True)
class Font:
    """Text font properties."""

    name: str | None = None
    size: float | None = None
    bold: bool = False
    italic: bool = False
    underline: str | None = None  # "single", "double", etc.
    strike: bool = False
    color: Color | str | None = None

    def _color_hex(self) -> str | None:
        """Resolve color to a '#RRGGBB' string or None."""
        if self.color is None:
            return None
        if isinstance(self.color, Color):
            return self.color.to_hex()
        raw = str(self.color).lstrip("#")
        if len(raw) == 8:
            return f"#{raw[2:]}"
        return f"#{raw}"


@dataclass(frozen=True, init=False)
class PatternFill:
    """Cell fill (solid pattern only for now).

    Accepts both ``patternType=`` and ``fill_type=`` for openpyxl compatibility.
    """

    patternType: str | None = None  # noqa: N815 — matches openpyxl name
    fgColor: Color | str | None = None  # noqa: N815

    def __init__(
        self,
        patternType: str | None = None,  # noqa: N803
        fgColor: Color | str | None = None,  # noqa: N803
        *,
        fill_type: str | None = None,
    ) -> None:
        # fill_type is openpyxl's alias for patternType
        object.__setattr__(self, "patternType", fill_type if patternType is None else patternType)
        object.__setattr__(self, "fgColor", fgColor)

    def _fg_hex(self) -> str | None:
        """Resolve fgColor to a '#RRGGBB' string or None."""
        if self.fgColor is None:
            return None
        if isinstance(self.fgColor, Color):
            return self.fgColor.to_hex()
        raw = str(self.fgColor).lstrip("#")
        if len(raw) == 8:
            return f"#{raw[2:]}"
        return f"#{raw}"


@dataclass(frozen=True)
class Side:
    """One edge of a border."""

    style: str | None = None  # "thin", "medium", "thick", etc.
    color: Color | str | None = None

    def _color_hex(self) -> str | None:
        if self.color is None:
            return None
        if isinstance(self.color, Color):
            return self.color.to_hex()
        raw = str(self.color).lstrip("#")
        if len(raw) == 8:
            return f"#{raw[2:]}"
        return f"#{raw}"


@dataclass(frozen=True)
class Border:
    """Cell borders."""

    left: Side = field(default_factory=Side)
    right: Side = field(default_factory=Side)
    top: Side = field(default_factory=Side)
    bottom: Side = field(default_factory=Side)


@dataclass(frozen=True)
class Alignment:
    """Cell alignment."""

    horizontal: str | None = None
    vertical: str | None = None
    wrap_text: bool = False
    text_rotation: int = 0
    indent: int = 0
