"""``NamedStyle`` + ``_NamedStyleList`` registry — RFC-064 §2.1."""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass, field
from typing import Any


@dataclass
class NamedStyle:
    """Named style (CT_CellStyle §18.8.7)."""

    name: str = ""
    font: Any = None
    fill: Any = None
    border: Any = None
    alignment: Any = None
    protection: Any = None
    number_format: str = "General"
    builtinId: int | None = None  # noqa: N815
    customBuiltin: bool = False  # noqa: N815
    hidden: bool = False
    xfId: int | None = None  # noqa: N815

    @property
    def is_builtin(self) -> bool:
        return self.builtinId is not None and not self.customBuiltin

    def to_rust_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "font": _style_to_dict(self.font),
            "fill": _style_to_dict(self.fill),
            "border": _style_to_dict(self.border),
            "alignment": _style_to_dict(self.alignment),
            "protection": _style_to_dict(self.protection),
            "number_format": self.number_format,
            "builtin_id": self.builtinId,
            "custom_builtin": self.customBuiltin,
            "hidden": self.hidden,
            "xf_id": self.xfId,
        }


def _style_to_dict(value: Any) -> dict[str, Any] | None:
    if value is None:
        return None
    if hasattr(value, "to_rust_dict"):
        return value.to_rust_dict()
    out: dict[str, Any] = {}
    for key in (
        "name",
        "size",
        "bold",
        "italic",
        "underline",
        "strike",
        "color",
        "patternType",
        "fgColor",
        "horizontal",
        "vertical",
        "wrap_text",
        "text_rotation",
        "indent",
        "locked",
        "hidden",
    ):
        attr = getattr(value, key, None)
        if attr is not None and attr is not False and attr != 0:
            out[key] = getattr(attr, "rgb", attr)
    return out or None


_BUILTIN_SEEDS: tuple[tuple[str, int], ...] = (
    ("Normal", 0),
    ("Comma", 3),
    ("Comma [0]", 6),
    ("Currency", 4),
    ("Currency [0]", 7),
    ("Percent", 5),
    ("Hyperlink", 8),
    ("Followed Hyperlink", 9),
    ("Note", 10),
    ("Warning Text", 11),
    ("Title", 15),
    ("Heading 1", 16),
    ("Heading 2", 17),
    ("Heading 3", 18),
    ("Heading 4", 19),
    ("Input", 20),
    ("Output", 21),
    ("Calculation", 22),
    ("Check Cell", 23),
    ("Linked Cell", 24),
    ("Total", 25),
    ("Good", 26),
    ("Bad", 27),
    ("Neutral", 28),
    ("Accent1", 29),
    ("20% - Accent1", 30),
    ("40% - Accent1", 31),
    ("60% - Accent1", 32),
    ("Accent2", 33),
    ("20% - Accent2", 34),
    ("40% - Accent2", 35),
    ("60% - Accent2", 36),
    ("Accent3", 37),
    ("20% - Accent3", 38),
    ("40% - Accent3", 39),
    ("60% - Accent3", 40),
    ("Accent4", 41),
    ("20% - Accent4", 42),
    ("40% - Accent4", 43),
    ("60% - Accent4", 44),
    ("Accent5", 45),
    ("20% - Accent5", 46),
    ("40% - Accent5", 47),
    ("60% - Accent5", 48),
    ("Accent6", 49),
    ("20% - Accent6", 50),
    ("40% - Accent6", 51),
    ("60% - Accent6", 52),
)


@dataclass
class _NamedStyleList:
    """Workbook-level named-style registry exposed as ``wb.named_styles``."""

    _styles: list[NamedStyle] = field(default_factory=list)
    _by_name: dict[str, NamedStyle] = field(default_factory=dict)
    _seeded: bool = False

    def _seed_builtins(self) -> None:
        if self._seeded:
            return
        self._seeded = True
        for name, builtin_id in _BUILTIN_SEEDS:
            ns = NamedStyle(name=name, builtinId=builtin_id)
            self._styles.append(ns)
            self._by_name[name] = ns

    def append(self, ns: NamedStyle) -> None:
        if not isinstance(ns, NamedStyle):
            raise TypeError(
                f"named_styles.append requires a NamedStyle, got {type(ns).__name__}"
            )
        if not ns.name:
            raise ValueError("NamedStyle.name must be a non-empty string before append")
        self._seed_builtins()
        prior = self._by_name.get(ns.name)
        if prior is not None:
            self._styles[self._styles.index(prior)] = ns
        else:
            self._styles.append(ns)
        self._by_name[ns.name] = ns

    def add(self, ns: NamedStyle) -> None:
        self.append(ns)

    def __getitem__(self, name: str) -> NamedStyle:
        self._seed_builtins()
        try:
            return self._by_name[name]
        except KeyError:
            raise KeyError(f"NamedStyle {name!r} is not registered on this workbook") from None

    def __contains__(self, name: object) -> bool:
        self._seed_builtins()
        return name in self._by_name

    def __iter__(self) -> Iterator[NamedStyle]:
        self._seed_builtins()
        return iter(self._styles)

    def __len__(self) -> int:
        self._seed_builtins()
        return len(self._styles)

    def names(self) -> list[str]:
        self._seed_builtins()
        return [ns.name for ns in self._styles]

    def user_styles(self) -> list[NamedStyle]:
        self._seed_builtins()
        return [ns for ns in self._styles if not ns.is_builtin]


__all__ = ["NamedStyle", "_NamedStyleList"]
