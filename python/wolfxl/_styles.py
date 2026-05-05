"""Style dataclasses matching openpyxl's public names.

These are lightweight, frozen value objects. They mirror openpyxl's Font,
PatternFill, Border, Side, Alignment, and Color classes so that code written
for openpyxl can be ported with minimal changes.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import ClassVar
from xml.etree import ElementTree as ET

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


def _bool_attr(value: bool | None) -> str | None:
    if value is None:
        return None
    return "1" if value else "0"


def _argb(value: Color | str | None) -> str | None:
    if value is None:
        return None
    if isinstance(value, Color):
        if value.rgb is not None:
            raw = value.rgb
        elif value.indexed is not None:
            return None
        else:
            return None
    else:
        raw = str(value)
    raw = raw.lstrip("#").upper()
    if len(raw) == 6:
        return f"00{raw}"
    return raw


class _TreeMixin:
    tagname: ClassVar[str]

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002 - openpyxl signature
        namespace: str | None = None,  # noqa: ARG002 - openpyxl signature
    ) -> ET.Element:
        """Serialize this style value to a compact ElementTree node."""
        return ET.Element(tagname or self.tagname)

    @classmethod
    def from_tree(cls, node: ET.Element):  # type: ignore[no-untyped-def]
        """Build this style value from a compact ElementTree node."""
        return cls(**dict(node.attrib))


@dataclass(frozen=True, init=False)
class Color(_TreeMixin):
    """An Excel color.

    A color is identified by exactly one of three strategies: direct RGB,
    theme index (with optional tint), or legacy indexed palette position.
    Which strategy is in effect is recorded in ``type`` and mirrors
    openpyxl's ``Color`` class shape so callers inspecting ``color.type``
    see the same vocabulary.

    The default `Color()` is opaque black RGB, matching openpyxl.
    """

    rgb: str | None = None
    auto: bool | None = None
    theme: int | None = None
    indexed: int | None = None
    tint: float = 0.0
    type: str = "rgb"
    tagname: ClassVar[str] = "color"
    namespace: ClassVar[str | None] = None
    idx_base: ClassVar[int] = 0

    def __init__(
        self,
        rgb: str | None = None,
        *,
        auto: bool | None = None,
        theme: int | None = None,
        indexed: int | None = None,
        tint: float = 0.0,
        type: str | None = None,  # noqa: A002 — openpyxl uses this name
    ) -> None:
        # Resolve which constructor form was used. If theme/indexed/rgb is
        # set, that wins; otherwise default to opaque black RGB to match
        # openpyxl's no-arg default.
        if auto is not None:
            resolved_type = type or "auto"
            resolved_rgb = None
        elif theme is not None:
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
        object.__setattr__(self, "auto", auto)
        object.__setattr__(self, "theme", theme)
        object.__setattr__(self, "indexed", indexed)
        object.__setattr__(self, "tint", tint)
        object.__setattr__(self, "type", resolved_type)

    @property
    def value(self) -> str | int | bool | None:
        """Return the active color value, matching openpyxl's alias."""
        return getattr(self, self.type, self.rgb)

    @property
    def index(self) -> str | int | bool | None:
        """Openpyxl compatibility alias for :attr:`value`."""
        return self.value

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

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002
        namespace: str | None = None,  # noqa: ARG002
    ) -> ET.Element:
        node = ET.Element(tagname or self.tagname)
        if self.type == "rgb" and self.rgb is not None:
            node.set("rgb", _argb(self.rgb) or self.rgb)
        elif self.type == "indexed" and self.indexed is not None:
            node.set("indexed", str(self.indexed))
        elif self.type == "theme" and self.theme is not None:
            node.set("theme", str(self.theme))
        elif self.type == "auto" and self.auto is not None:
            node.set("auto", _bool_attr(self.auto) or "0")
        if self.tint:
            node.set("tint", str(self.tint))
        return node

    @classmethod
    def from_tree(cls, node: ET.Element) -> Color:
        attrs = node.attrib
        if "rgb" in attrs:
            return cls(rgb=attrs["rgb"], tint=float(attrs.get("tint", 0.0)))
        if "indexed" in attrs:
            return cls(indexed=int(attrs["indexed"]), tint=float(attrs.get("tint", 0.0)))
        if "theme" in attrs:
            return cls(theme=int(attrs["theme"]), tint=float(attrs.get("tint", 0.0)))
        if "auto" in attrs:
            return cls(auto=attrs["auto"] not in {"0", "false", "False"})
        return cls()


@dataclass(frozen=True, init=False)
class Font(_TreeMixin):
    """Text font properties."""

    name: str | None = None
    size: float | None = None
    bold: bool = False
    italic: bool = False
    underline: str | None = None  # "single", "double", etc.
    strike: bool = False
    color: Color | str | None = None
    family: float | None = None
    charset: int | None = None
    scheme: str | None = None
    vertAlign: str | None = None  # noqa: N815
    outline: bool = False
    shadow: bool = False
    condense: bool = False
    extend: bool = False
    tagname: ClassVar[str] = "font"
    namespace: ClassVar[str | None] = None
    idx_base: ClassVar[int] = 0
    UNDERLINE_SINGLE: ClassVar[str] = "single"
    UNDERLINE_DOUBLE: ClassVar[str] = "double"
    UNDERLINE_SINGLE_ACCOUNTING: ClassVar[str] = "singleAccounting"
    UNDERLINE_DOUBLE_ACCOUNTING: ClassVar[str] = "doubleAccounting"

    def __init__(
        self,
        name: str | None = None,
        size: float | None = None,
        bold: bool = False,
        italic: bool = False,
        underline: str | bool | None = None,
        strike: bool = False,
        color: Color | str | None = None,
        *,
        sz: float | None = None,
        b: bool | None = None,
        i: bool | None = None,
        u: str | bool | None = None,
        strikethrough: bool | None = None,
        family: float | None = None,
        charset: int | None = None,
        scheme: str | None = None,
        vertAlign: str | None = None,  # noqa: N803
        outline: bool = False,
        shadow: bool = False,
        condense: bool = False,
        extend: bool = False,
    ) -> None:
        resolved_underline = u if u is not None else underline
        if resolved_underline is True:
            resolved_underline = "single"
        object.__setattr__(self, "name", name)
        object.__setattr__(self, "size", sz if sz is not None else size)
        object.__setattr__(self, "bold", bool(b) if b is not None else bool(bold))
        object.__setattr__(self, "italic", bool(i) if i is not None else bool(italic))
        object.__setattr__(self, "underline", resolved_underline)
        object.__setattr__(
            self,
            "strike",
            bool(strikethrough) if strikethrough is not None else bool(strike),
        )
        object.__setattr__(self, "color", color)
        object.__setattr__(self, "family", family)
        object.__setattr__(self, "charset", charset)
        object.__setattr__(self, "scheme", scheme)
        object.__setattr__(self, "vertAlign", vertAlign)
        object.__setattr__(self, "outline", bool(outline))
        object.__setattr__(self, "shadow", bool(shadow))
        object.__setattr__(self, "condense", bool(condense))
        object.__setattr__(self, "extend", bool(extend))

    @property
    def b(self) -> bool:
        return self.bold

    @property
    def i(self) -> bool:
        return self.italic

    @property
    def u(self) -> str | None:
        return self.underline

    @property
    def sz(self) -> float | None:
        return self.size

    @property
    def strikethrough(self) -> bool:
        return self.strike

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

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002
        namespace: str | None = None,  # noqa: ARG002
    ) -> ET.Element:
        node = ET.Element(tagname or self.tagname)
        for attr, tag in ((self.bold, "b"), (self.italic, "i"), (self.strike, "strike")):
            if attr:
                child = ET.SubElement(node, tag)
                child.set("val", "1")
        if self.color is not None:
            node.append(Color(rgb=_argb(self.color)).to_tree("color"))
        if self.size is not None:
            ET.SubElement(node, "sz").set("val", str(self.size))
        if self.name is not None:
            ET.SubElement(node, "name").set("val", self.name)
        if self.underline is not None:
            child = ET.SubElement(node, "u")
            if self.underline not in {True, "single"}:
                child.set("val", str(self.underline))
        if self.vertAlign is not None:
            ET.SubElement(node, "vertAlign").set("val", self.vertAlign)
        if self.scheme is not None:
            ET.SubElement(node, "scheme").set("val", self.scheme)
        return node

    @classmethod
    def from_tree(cls, node: ET.Element) -> Font:
        kwargs: dict[str, object] = {}
        for child in node:
            if child.tag == "b":
                kwargs["bold"] = child.attrib.get("val", "1") != "0"
            elif child.tag == "i":
                kwargs["italic"] = child.attrib.get("val", "1") != "0"
            elif child.tag == "strike":
                kwargs["strike"] = child.attrib.get("val", "1") != "0"
            elif child.tag == "color":
                kwargs["color"] = Color.from_tree(child)
            elif child.tag == "sz":
                kwargs["size"] = float(child.attrib["val"])
            elif child.tag == "name":
                kwargs["name"] = child.attrib.get("val")
            elif child.tag == "u":
                kwargs["underline"] = child.attrib.get("val", "single")
            elif child.tag == "vertAlign":
                kwargs["vertAlign"] = child.attrib.get("val")
            elif child.tag == "scheme":
                kwargs["scheme"] = child.attrib.get("val")
        return cls(**kwargs)


@dataclass(frozen=True, init=False)
class PatternFill(_TreeMixin):
    """Cell fill (solid pattern only for now).

    Accepts both ``patternType=`` and ``fill_type=`` for openpyxl compatibility.
    """

    patternType: str | None = None  # noqa: N815 — matches openpyxl name
    fgColor: Color | str | None = None  # noqa: N815
    bgColor: Color | str | None = None  # noqa: N815
    tagname: ClassVar[str] = "patternFill"
    namespace: ClassVar[str | None] = None
    idx_base: ClassVar[int] = 0

    def __init__(
        self,
        patternType: str | None = None,  # noqa: N803
        fgColor: Color | str | None = None,  # noqa: N803
        bgColor: Color | str | None = None,  # noqa: N803
        *,
        fill_type: str | None = None,
        start_color: Color | str | None = None,
        end_color: Color | str | None = None,
    ) -> None:
        # fill_type/start_color/end_color are openpyxl aliases.
        object.__setattr__(self, "patternType", fill_type if patternType is None else patternType)
        object.__setattr__(self, "fgColor", start_color if fgColor is None else fgColor)
        object.__setattr__(self, "bgColor", end_color if bgColor is None else bgColor)

    @property
    def fill_type(self) -> str | None:
        return self.patternType

    @property
    def start_color(self) -> Color | str | None:
        return self.fgColor

    @property
    def end_color(self) -> Color | str | None:
        return self.bgColor

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

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002
        namespace: str | None = None,  # noqa: ARG002
    ) -> ET.Element:
        outer = ET.Element(tagname or "fill")
        inner = ET.SubElement(outer, self.tagname)
        if self.patternType is not None:
            inner.set("patternType", self.patternType)
        if self.fgColor is not None:
            inner.append(Color(rgb=_argb(self.fgColor)).to_tree("fgColor"))
        if self.bgColor is not None:
            inner.append(Color(rgb=_argb(self.bgColor)).to_tree("bgColor"))
        return outer

    @classmethod
    def from_tree(cls, node: ET.Element) -> PatternFill:
        inner = next(iter(node), node)
        kwargs: dict[str, object] = {"patternType": inner.attrib.get("patternType")}
        for child in inner:
            if child.tag == "fgColor":
                kwargs["fgColor"] = Color.from_tree(child)
            elif child.tag == "bgColor":
                kwargs["bgColor"] = Color.from_tree(child)
        return cls(**kwargs)


@dataclass(frozen=True, init=False)
class Side(_TreeMixin):
    """One edge of a border."""

    style: str | None = None  # "thin", "medium", "thick", etc.
    color: Color | str | None = None
    namespace: ClassVar[str | None] = None
    idx_base: ClassVar[int] = 0
    tagname: ClassVar[str] = "side"

    def __init__(
        self,
        style: str | None = None,
        color: Color | str | None = None,
        *,
        border_style: str | None = None,
    ) -> None:
        object.__setattr__(self, "style", border_style if style is None else style)
        object.__setattr__(self, "color", color)

    @property
    def border_style(self) -> str | None:
        return self.style

    def _color_hex(self) -> str | None:
        if self.color is None:
            return None
        if isinstance(self.color, Color):
            return self.color.to_hex()
        raw = str(self.color).lstrip("#")
        if len(raw) == 8:
            return f"#{raw[2:]}"
        return f"#{raw}"

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002
        namespace: str | None = None,  # noqa: ARG002
    ) -> ET.Element:
        node = ET.Element(tagname or self.tagname)
        if self.style is not None:
            node.set("style", self.style)
        if self.color is not None:
            node.append(Color(rgb=_argb(self.color)).to_tree("color"))
        return node

    @classmethod
    def from_tree(cls, node: ET.Element) -> Side:
        color = None
        for child in node:
            if child.tag == "color":
                color = Color.from_tree(child)
        return cls(style=node.attrib.get("style"), color=color)


@dataclass(frozen=True)
class Border(_TreeMixin):
    """Cell borders."""

    left: Side = field(default_factory=Side)
    right: Side = field(default_factory=Side)
    top: Side = field(default_factory=Side)
    bottom: Side = field(default_factory=Side)
    diagonal: Side = field(default_factory=Side)
    vertical: Side | None = None
    horizontal: Side | None = None
    start: Side | None = None
    end: Side | None = None
    diagonalUp: bool = False  # noqa: N815
    diagonalDown: bool = False  # noqa: N815
    outline: bool = True
    tagname: ClassVar[str] = "border"
    namespace: ClassVar[str | None] = None
    idx_base: ClassVar[int] = 0

    @property
    def diagonal_direction(self) -> str | None:
        if self.diagonalUp and self.diagonalDown:
            return "both"
        if self.diagonalUp:
            return "up"
        if self.diagonalDown:
            return "down"
        return None

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002
        namespace: str | None = None,  # noqa: ARG002
    ) -> ET.Element:
        node = ET.Element(tagname or self.tagname)
        if self.diagonalUp:
            node.set("diagonalUp", "1")
        if self.diagonalDown:
            node.set("diagonalDown", "1")
        for name in ("left", "right", "top", "bottom", "diagonal"):
            side = getattr(self, name)
            if side is not None and side != Side():
                node.append(side.to_tree(name))
        return node

    @classmethod
    def from_tree(cls, node: ET.Element) -> Border:
        kwargs: dict[str, object] = {
            "diagonalUp": node.attrib.get("diagonalUp") == "1",
            "diagonalDown": node.attrib.get("diagonalDown") == "1",
        }
        for child in node:
            if child.tag in {"left", "right", "top", "bottom", "diagonal"}:
                kwargs[child.tag] = Side.from_tree(child)
        return cls(**kwargs)


@dataclass(frozen=True, init=False)
class Alignment(_TreeMixin):
    """Cell alignment."""

    horizontal: str | None = None
    vertical: str | None = None
    wrap_text: bool = False
    text_rotation: int = 0
    indent: int = 0
    relativeIndent: int = 0  # noqa: N815
    justifyLastLine: bool | None = None  # noqa: N815
    readingOrder: float = 0.0  # noqa: N815
    shrink_to_fit: bool | None = None
    tagname: ClassVar[str] = "alignment"
    namespace: ClassVar[str | None] = None
    idx_base: ClassVar[int] = 0

    def __init__(
        self,
        horizontal: str | None = None,
        vertical: str | None = None,
        wrap_text: bool | None = None,
        text_rotation: int | None = None,
        indent: int = 0,
        *,
        wrapText: bool | None = None,  # noqa: N803
        textRotation: int | None = None,  # noqa: N803
        shrink_to_fit: bool | None = None,
        shrinkToFit: bool | None = None,  # noqa: N803
        relativeIndent: int = 0,  # noqa: N803
        justifyLastLine: bool | None = None,  # noqa: N803
        readingOrder: float = 0.0,  # noqa: N803
    ) -> None:
        object.__setattr__(self, "horizontal", horizontal)
        object.__setattr__(self, "vertical", vertical)
        object.__setattr__(self, "wrap_text", bool(wrapText if wrapText is not None else wrap_text))
        object.__setattr__(
            self,
            "text_rotation",
            int(textRotation if textRotation is not None else (text_rotation or 0)),
        )
        object.__setattr__(self, "indent", int(indent))
        object.__setattr__(self, "relativeIndent", int(relativeIndent))
        object.__setattr__(self, "justifyLastLine", justifyLastLine)
        object.__setattr__(self, "readingOrder", readingOrder)
        object.__setattr__(
            self,
            "shrink_to_fit",
            shrinkToFit if shrinkToFit is not None else shrink_to_fit,
        )

    @property
    def wrapText(self) -> bool:
        return self.wrap_text

    @property
    def textRotation(self) -> int:
        return self.text_rotation

    @property
    def shrinkToFit(self) -> bool | None:
        return self.shrink_to_fit

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002
        namespace: str | None = None,  # noqa: ARG002
    ) -> ET.Element:
        node = ET.Element(tagname or self.tagname)
        attrs = {
            "horizontal": self.horizontal,
            "vertical": self.vertical,
            "wrapText": _bool_attr(self.wrap_text),
            "textRotation": str(self.text_rotation) if self.text_rotation else None,
            "indent": str(self.indent) if self.indent else None,
            "relativeIndent": str(self.relativeIndent) if self.relativeIndent else None,
            "justifyLastLine": _bool_attr(self.justifyLastLine),
            "readingOrder": str(self.readingOrder) if self.readingOrder else None,
            "shrinkToFit": _bool_attr(self.shrink_to_fit),
        }
        for key, value in attrs.items():
            if value is not None:
                node.set(key, value)
        return node

    @classmethod
    def from_tree(cls, node: ET.Element) -> Alignment:
        attrs = node.attrib
        return cls(
            horizontal=attrs.get("horizontal"),
            vertical=attrs.get("vertical"),
            wrapText=attrs.get("wrapText") == "1" if "wrapText" in attrs else None,
            textRotation=int(attrs.get("textRotation", 0)),
            indent=int(attrs.get("indent", 0)),
            relativeIndent=int(attrs.get("relativeIndent", 0)),
            justifyLastLine=attrs.get("justifyLastLine") == "1"
            if "justifyLastLine" in attrs
            else None,
            readingOrder=float(attrs.get("readingOrder", 0.0)),
            shrinkToFit=attrs.get("shrinkToFit") == "1" if "shrinkToFit" in attrs else None,
        )
