"""openpyxl.formatting compatibility.

Exposes ``ConditionalFormatting`` (one range + its rules) and
``ConditionalFormattingList`` (the ws.conditional_formatting container).
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

from wolfxl.formatting.rule import Rule
from wolfxl.worksheet.cell_range import MultiCellRange

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


# Rule kinds that wolfxl can serialize end-to-end (write mode + modify
# mode). Anything outside this set raises NotImplementedError on
# ``ConditionalFormattingList.add`` until the CF expansion wave lands.
# See ``crates/wolfxl-writer/src/emit/sheet_xml.rs::emit_conditional_formats``
# stub-variant arm for the same list on the writer side.
_SUPPORTED_CF_KINDS: frozenset[str] = frozenset(
    {
        "cellIs",
        "expression",
        "colorScale",
        "dataBar",
        "iconSet",
        "top10",
        "uniqueValues",
        "duplicateValues",
        "containsText",
        "notContainsText",
        "beginsWith",
        "endsWith",
        "containsBlanks",
        "notContainsBlanks",
        "containsErrors",
        "notContainsErrors",
        "timePeriod",
        "aboveAverage",
    }
)


@dataclass
class ConditionalFormatting:
    """All CF rules that apply to a single range.

    openpyxl groups rules by range: one ``ConditionalFormatting`` per
    distinct ``sqref``. ``cfRule`` is the legacy openpyxl alias for
    ``rules`` — both return the same list.
    """

    sqref: str
    rules: list[Rule] = field(default_factory=list)
    pivot: bool | None = None

    @property
    def cells(self) -> MultiCellRange:
        return MultiCellRange(self.sqref)

    @property
    def cfRule(self) -> list[Rule]:  # noqa: N802 - openpyxl alias
        return self.rules


class ConditionalFormattingList:
    """Container for a worksheet's conditional formatting entries.

    Iterates ``ConditionalFormatting`` objects. In write mode, users
    attach new CF rules via ``ws.conditional_formatting.add(range, rule)``
    — that lands in PR5. Reads work in any mode.
    """

    __slots__ = ("_entries", "_ws")

    def __init__(self, ws: Worksheet | None = None) -> None:
        self._entries: list[ConditionalFormatting] = []
        self._ws = ws

    def __iter__(self) -> Iterator[ConditionalFormatting]:
        return iter(self._entries)

    def __len__(self) -> int:
        return len(self._entries)

    def __bool__(self) -> bool:
        return bool(self._entries)

    @property
    def max_priority(self) -> int:
        """Highest explicit rule priority, matching openpyxl's helper."""
        priorities = [
            int(priority)
            for entry in self._entries
            for rule in entry.rules
            if (priority := getattr(rule, "priority", None)) is not None
        ]
        return max(priorities, default=0)

    def _append_entry(self, entry: ConditionalFormatting) -> None:
        """Internal: used by the lazy reader to populate the container."""
        self._entries.append(entry)

    def add(self, range_string: str, rule: Rule) -> None:
        """Attach a new conditional-formatting rule.

        Works in write mode (``Workbook()``) and modify mode
        (``load_workbook(path, modify=True)``). Both modes queue onto
        ``ws._pending_conditional_formats`` here; the workbook's
        ``save()`` routes through to the right backend (RFC-026 wires
        the modify-mode flush).

        Supports openpyxl's public ``Rule`` taxonomy, including text,
        duplicate/unique, blanks/errors, time-period, top10, and
        above-average rules.
        """
        ws = self._ws
        if ws is None:
            raise RuntimeError("ConditionalFormattingList is not attached to a worksheet")
        wb = ws._workbook  # noqa: SLF001
        if wb._rust_writer is None and wb._rust_patcher is None:  # noqa: SLF001
            raise RuntimeError("ConditionalFormattingList is not attached to a workbook")
        if rule.type not in _SUPPORTED_CF_KINDS:
            raise NotImplementedError(
                f"Conditional-formatting rule type {rule.type!r} is not yet supported. "
                f"Supported in this release: {sorted(_SUPPORTED_CF_KINDS)}. "
                "See Plans/rfcs/026-conditional-formatting.md §10 for the expansion wave."
            )
        # Find or create the CF entry for this range in our container.
        for entry in self._entries:
            if entry.sqref == range_string:
                entry.rules.append(rule)
                break
        else:
            self._entries.append(ConditionalFormatting(sqref=range_string, rules=[rule]))
        ws._pending_conditional_formats.append((range_string, rule))  # noqa: SLF001


def _cf_to_patcher_dict(sqref: str, rules: list[Rule]) -> dict[str, Any]:
    """Convert (sqref, openpyxl-shaped Rules) into the patcher's payload.

    Returns ONE ``ConditionalFormattingPatch`` dict per ``sqref`` with a
    ``rules: list[dict]`` key. Mirrors RFC-026 §4.2's
    ``ConditionalFormattingPatch`` + ``CfRulePatch`` shape. Filters out
    ``None`` values before crossing the PyO3 boundary (RFC-025 lesson:
    PyO3 ``.extract::<String>()`` rejects ``None`` — pass either a real
    string or omit the key entirely).

    For ColorScale / DataBar rules, pulls extra keys (``start_type``,
    ``start_color``, etc.) out of ``rule.extra``.
    """
    rule_dicts: list[dict[str, Any]] = []
    for rule in rules:
        rd: dict[str, Any] = {
            "kind": rule.type,
            "stop_if_true": bool(rule.stopIfTrue),
        }
        _add_generic_cf_attrs(rd, rule)
        if rule.type == "cellIs":
            if rule.operator is not None:
                rd["operator"] = rule.operator
            formulas = list(rule.formula or [])
            if formulas:
                rd["formula_a"] = formulas[0]
            if len(formulas) > 1:
                rd["formula_b"] = formulas[1]
            rd["dxf"] = _dxf_from_rule(rule)
        elif rule.type == "expression":
            formulas = list(rule.formula or [])
            if formulas:
                rd["formula"] = formulas[0]
            rd["dxf"] = _dxf_from_rule(rule)
        elif rule.type == "colorScale":
            extra = rule.extra or {}
            stops: list[dict[str, Any]] = []
            for prefix in ("start", "mid", "end"):
                t = extra.get(f"{prefix}_type")
                if t is None:
                    continue
                stop: dict[str, Any] = {"cfvo_type": t}
                v = extra.get(f"{prefix}_value")
                if v is not None:
                    stop["val"] = str(v)
                color = extra.get(f"{prefix}_color")
                if color is not None:
                    stop["color_rgb"] = _normalize_color(color)
                stops.append(stop)
            rd["stops"] = stops
            # ColorScale has no dxf — keep the key absent so patcher
            # treats it as None.
        elif rule.type == "dataBar":
            extra = rule.extra or {}
            if extra.get("start_type") is not None:
                rd["min_cfvo_type"] = extra["start_type"]
            if extra.get("start_value") is not None:
                rd["min_val"] = str(extra["start_value"])
            if extra.get("end_type") is not None:
                rd["max_cfvo_type"] = extra["end_type"]
            if extra.get("end_value") is not None:
                rd["max_val"] = str(extra["end_value"])
            color = extra.get("color")
            if color is not None:
                rd["color_rgb"] = _normalize_color(color)
            if extra.get("show_value") is not None:
                rd["show_value"] = bool(extra["show_value"])
            if extra.get("min_length") is not None:
                rd["min_length"] = int(extra["min_length"])
            if extra.get("max_length") is not None:
                rd["max_length"] = int(extra["max_length"])
        elif rule.type == "iconSet":
            extra = rule.extra or {}
            if extra.get("icon_style") is not None:
                rd["icon_style"] = extra["icon_style"]
            if extra.get("value_type") is not None:
                rd["value_type"] = extra["value_type"]
            if extra.get("values") is not None:
                rd["values"] = list(extra["values"])
            if extra.get("show_value") is not None:
                rd["show_value"] = bool(extra["show_value"])
            if extra.get("percent") is not None:
                rd["percent"] = bool(extra["percent"])
            if extra.get("reverse") is not None:
                rd["reverse"] = bool(extra["reverse"])
        else:
            rd["dxf"] = _dxf_from_rule(rule)
        rule_dicts.append({k: v for k, v in rd.items() if v is not None})

    return {"sqref": sqref, "rules": rule_dicts}


def _add_generic_cf_attrs(rd: dict[str, Any], rule: Rule) -> None:
    formulas = [str(part) for part in rule.formula] if rule.formula else []
    if formulas:
        rd["formulas"] = formulas
    if rule.operator is not None:
        rd["operator"] = rule.operator
    attr_map = {
        "aboveAverage": "above_average",
        "percent": "percent",
        "bottom": "bottom",
        "text": "text",
        "timePeriod": "time_period",
        "rank": "rank",
        "stdDev": "std_dev",
        "equalAverage": "equal_average",
    }
    for public_name, payload_name in attr_map.items():
        value = getattr(rule, public_name, None)
        if value is not None:
            rd[payload_name] = value


def _dxf_from_rule(rule: Rule) -> dict[str, Any] | None:
    """Pull a Rust-shaped dxf dict out of a ``Rule.extra`` blob.

    The wolfxl ``Rule`` dataclass stashes user-supplied formatting under
    the ``extra`` dict (since openpyxl's descriptor system is not
    mirrored here). We accept either an explicit ``"dxf"`` sub-dict
    inside ``extra`` (preferred for callers that build via the patcher
    payload directly) or a flat ``font``/``fill`` shape (preferred for
    openpyxl-style construction). Returns ``None`` if no formatting was
    supplied — the patcher then emits a `<cfRule>` without a `dxfId`.
    """
    extra = rule.extra or {}
    if "dxf" in extra and extra["dxf"] is not None:
        return _normalize_dxf_dict(extra["dxf"])
    # Heuristic: if any of the wolfxl-shaped dxf keys exist directly on
    # extra, treat the whole `extra` as a dxf blob.
    direct_keys = {
        "font_bold",
        "font_italic",
        "font_color_rgb",
        "fill_pattern_type",
        "fill_fg_color_rgb",
        "border_top_style",
        "border_bottom_style",
        "border_left_style",
        "border_right_style",
    }
    if any(k in extra for k in direct_keys):
        return _normalize_dxf_dict({k: extra[k] for k in direct_keys if k in extra})
    return None


def _normalize_dxf_dict(d: Any) -> dict[str, Any]:
    """Strip ``None`` values so PyO3's ``extract::<String>()`` doesn't reject them.

    Color fields normalize to ``"FFRRGGBB"`` so the Rust side can write
    them straight into the OOXML ARGB attribute.
    """
    if not isinstance(d, dict):
        d = _dxf_object_to_dict(d)
    out: dict[str, Any] = {}
    for key, val in d.items():
        if val is None:
            continue
        if key in {"font_color_rgb", "fill_fg_color_rgb"}:
            out[key] = _normalize_color(val)
        else:
            out[key] = val
    return out


def _dxf_object_to_dict(dxf: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    font = getattr(dxf, "font", None)
    if font is not None:
        if getattr(font, "bold", None) is not None:
            out["font_bold"] = bool(font.bold)
        if getattr(font, "italic", None) is not None:
            out["font_italic"] = bool(font.italic)
        color = getattr(font, "color", None)
        if color is not None and getattr(color, "rgb", None):
            out["font_color_rgb"] = color.rgb
    fill = getattr(dxf, "fill", None)
    if fill is not None:
        out["fill_pattern_type"] = (
            getattr(fill, "fill_type", None)
            or getattr(fill, "patternType", None)
            or "solid"
        )
        color = (
            getattr(fill, "fgColor", None)
            or getattr(fill, "start_color", None)
            or getattr(fill, "color", None)
        )
        if color is not None and getattr(color, "rgb", None):
            out["fill_fg_color_rgb"] = color.rgb
    return out


def _normalize_color(color: str) -> str:
    """Normalize ``"#RRGGBB"`` / ``"RRGGBB"`` / ``"FFRRGGBB"`` to OOXML ARGB."""
    s = color.lstrip("#").upper()
    if len(s) == 6:
        return f"FF{s}"
    if len(s) == 8:
        return s
    return f"FF{s}"


__all__ = [
    "ConditionalFormatting",
    "ConditionalFormattingList",
]
