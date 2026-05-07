"""openpyxl.formatting.rule compatibility.

Conditional-formatting rules are real dataclasses in T1. Each specific
rule type (``CellIsRule``, ``FormulaRule``, ``ColorScaleRule``, etc.)
constructs a ``Rule`` with the correct ``type`` tag set.

Excel's rule taxonomy:

- ``cellIs`` — compares cell value to a formula (``CellIsRule``)
- ``expression`` — arbitrary boolean formula (``FormulaRule``)
- ``colorScale`` — 2/3-color gradient (``ColorScaleRule``)
- ``dataBar`` — horizontal bar (``DataBarRule``)
- ``iconSet`` — traffic-light icons (``IconSetRule``)

wolfxl stores the rule metadata; actual style (color scale stops, bar
color) is preserved on modify-mode round-trip via the Rust layer but
not fully exposed to Python construction yet. Write-mode authoring
via ``ws.conditional_formatting.add()`` lands in PR5.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(init=False)
class Rule:
    """A generic conditional-formatting rule.

    Direct construction is rare — users build via the specific
    subclasses below, which default ``type`` to the right tag. The
    generic class exists because Excel has more rule types (above/below
    average, top/bottom, duplicates, text contains, ...) than we've
    wrapped with a dedicated constructor.
    """

    type: str
    priority: int = 1
    operator: str | None = None
    formula: list[str] = field(default_factory=list)
    stopIfTrue: bool = False  # noqa: N815 - openpyxl public API
    dxfId: int | None = None  # noqa: N815 - openpyxl public API
    aboveAverage: bool | None = None  # noqa: N815 - openpyxl public API
    percent: bool | None = None
    bottom: bool | None = None
    text: str | None = None
    timePeriod: str | None = None  # noqa: N815 - openpyxl public API
    rank: int | None = None
    stdDev: int | None = None  # noqa: N815 - openpyxl public API
    equalAverage: bool | None = None  # noqa: N815 - openpyxl public API
    # ``color_scale`` / ``data_bar`` / ``icon_set`` metadata blobs are
    # preserved on round-trip but not decomposed here — T2 territory.
    extra: dict[str, Any] = field(default_factory=dict)

    def __init__(
        self,
        type: str,  # noqa: A002 - openpyxl public API
        dxfId: int | None = None,  # noqa: N803
        priority: int = 1,
        stopIfTrue: bool | None = False,  # noqa: N803
        aboveAverage: bool | None = None,  # noqa: N803
        percent: bool | None = None,
        bottom: bool | None = None,
        operator: str | None = None,
        text: str | None = None,
        timePeriod: str | None = None,  # noqa: N803
        rank: int | None = None,
        stdDev: int | None = None,  # noqa: N803
        equalAverage: bool | None = None,  # noqa: N803
        formula: list[str] | tuple[str, ...] | str | None = None,
        dxf: Any = None,
        extra: dict[str, Any] | None = None,
        **kw: Any,
    ) -> None:
        self.type = type
        self.priority = priority
        self.operator = operator
        if formula is None:
            self.formula = []
        elif isinstance(formula, str):
            self.formula = [formula]
        else:
            self.formula = [str(item) for item in formula]
        self.stopIfTrue = bool(stopIfTrue) if stopIfTrue is not None else False
        self.dxfId = dxfId
        self.aboveAverage = aboveAverage
        self.percent = percent
        self.bottom = bottom
        self.text = text
        self.timePeriod = timePeriod
        self.rank = rank
        self.stdDev = stdDev
        self.equalAverage = equalAverage
        extras = dict(extra or {})
        if dxf is not None:
            extras["dxf"] = dxf
        for key in ("colorScale", "dataBar", "iconSet", "extLst"):
            if kw.get(key) is not None:
                extras[key] = kw[key]
        # Preserve unrecognized keyword payloads instead of rejecting
        # openpyxl-shaped Rule construction that carries extension data.
        for key, value in kw.items():
            if key not in extras and value is not None:
                extras[key] = value
        self.extra = extras

    @property
    def dxf(self) -> "DifferentialStyle | None":
        """openpyxl-shaped ``DifferentialStyle`` view of this rule's fill /
        font / border state.

        openpyxl exposes ``rule.dxf`` as a :class:`DifferentialStyle` whose
        ``font`` / ``fill`` / ``border`` mirror the kwargs the user passed
        in. Wolfxl stashes those kwargs inside :attr:`extra`; this property
        reconstructs the shim on demand. Returns ``None`` when no styling
        was supplied so callers can branch on truthiness.
        """
        extra = self.extra or {}
        if isinstance(extra.get("dxf"), DifferentialStyle):
            return extra["dxf"]
        if not any(extra.get(k) is not None for k in ("font", "fill", "border")):
            return None
        return DifferentialStyle(
            font=extra.get("font"),
            fill=extra.get("fill"),
            border=extra.get("border"),
        )

    @dxf.setter
    def dxf(self, value: "DifferentialStyle | None") -> None:
        if self.extra is None:
            self.extra = {}
        self.extra["dxf"] = value


def _absorb_dxf_kwargs(kw: dict[str, Any]) -> dict[str, Any]:
    """Pull openpyxl-shaped dxf kwargs (``fill=``, ``font=``, ``border=``, ``dxf=``)
    off ``kw`` and stash them inside ``extra`` so they survive the ``Rule``
    dataclass constructor (G14).

    openpyxl's ``CellIsRule(fill=PatternFill(...))`` collapses the kwarg into a
    ``DifferentialStyle`` and the rule grows a ``dxfId`` at write time.
    Wolfxl mirrors the surface here: the kwargs are recorded inside
    ``Rule.extra`` and the write-mode payload helper
    (``_conditional_format_payload``) translates them into the Rust-side cfg
    dict so ``dict_to_conditional_format`` can intern a ``DxfRecord`` and
    stamp the resulting ``dxfId`` on the emitted ``<cfRule>``.
    """
    if not any(k in kw for k in ("fill", "font", "border", "dxf")):
        return kw
    extra = dict(kw.pop("extra", {}) or {})
    for key in ("fill", "font", "border", "dxf"):
        if key in kw:
            extra[key] = kw.pop(key)
    kw["extra"] = extra
    return kw


class CellIsRule(Rule):
    """Conditional format triggered when a cell value matches an operator+operand.

    Example: ``CellIsRule(operator="greaterThan", formula=["50"])``.

    The openpyxl-compatible ``fill=PatternFill(...)`` / ``font=Font(...)`` /
    ``border=Border(...)`` / ``dxf=DifferentialStyle(...)`` kwargs are accepted
    and routed through ``Rule.extra`` so the writer can intern a matching
    ``<dxf>`` record and stamp its index as ``dxfId`` on the emitted
    ``<cfRule>`` (G14).
    """

    def __init__(
        self,
        operator: str | None = None,
        formula: list[str] | None = None,
        stopIfTrue: bool = False,  # noqa: N803
        **kw: Any,
    ) -> None:
        kw = _absorb_dxf_kwargs(kw)
        super().__init__(
            type="cellIs",
            operator=operator,
            formula=list(formula or []),
            stopIfTrue=stopIfTrue,
            **kw,
        )


class FormulaRule(Rule):
    """Conditional format triggered when a boolean formula is TRUE.

    Example: ``FormulaRule(formula=["$A1>100"])``.

    Accepts the same openpyxl ``fill=`` / ``font=`` / ``border=`` / ``dxf=``
    kwargs as ``CellIsRule``; they ride on ``Rule.extra`` and feed the dxf
    intern path on save (G14).
    """

    def __init__(
        self,
        formula: list[str] | None = None,
        stopIfTrue: bool = False,  # noqa: N803
        **kw: Any,
    ) -> None:
        kw = _absorb_dxf_kwargs(kw)
        super().__init__(
            type="expression",
            formula=list(formula or []),
            stopIfTrue=stopIfTrue,
            **kw,
        )


class ColorScaleRule(Rule):
    """2- or 3-stop color scale.

    ``start_type`` / ``mid_type`` / ``end_type`` are openpyxl's
    interpolation anchors (``"min"``, ``"max"``, ``"percentile"``,
    ``"num"``, ``"formula"``). We capture them in ``extra`` for round-
    trip; they don't feed the Rust writer yet (PR5 passes a simplified
    shape through).
    """

    def __init__(
        self,
        start_type: str | None = None,
        start_value: Any = None,
        start_color: str | None = None,
        mid_type: str | None = None,
        mid_value: Any = None,
        mid_color: str | None = None,
        end_type: str | None = None,
        end_value: Any = None,
        end_color: str | None = None,
        **kw: Any,
    ) -> None:
        extra = {
            "start_type": start_type,
            "start_value": start_value,
            "start_color": start_color,
            "mid_type": mid_type,
            "mid_value": mid_value,
            "mid_color": mid_color,
            "end_type": end_type,
            "end_value": end_value,
            "end_color": end_color,
        }
        super().__init__(type="colorScale", extra=extra, **kw)


class DataBarRule(Rule):
    """In-cell horizontal data bar."""

    def __init__(
        self,
        start_type: str | None = None,
        start_value: Any = None,
        end_type: str | None = None,
        end_value: Any = None,
        color: str | None = None,
        showValue: bool | None = None,  # noqa: N803
        minLength: int | None = None,  # noqa: N803
        maxLength: int | None = None,  # noqa: N803
        **kw: Any,
    ) -> None:
        extra = {
            "start_type": start_type,
            "start_value": start_value,
            "end_type": end_type,
            "end_value": end_value,
            "color": color,
            "show_value": True if showValue is None else showValue,
            "min_length": minLength,
            "max_length": maxLength,
        }
        super().__init__(type="dataBar", extra=extra, **kw)


class IconSetRule(Rule):
    """Icon set (3 arrows, 5 traffic lights, etc.)."""

    def __init__(
        self,
        icon_style: str | None = None,
        type: str | None = None,  # noqa: A002 - keyword openpyxl uses
        values: list[Any] | None = None,
        showValue: bool | None = None,  # noqa: N803
        percent: bool | None = None,
        reverse: bool | None = None,
        **kw: Any,
    ) -> None:
        # openpyxl's IconSetRule positional ``type`` ("percent", "percentile",
        # "num", "formula") differs from Rule.type — we stash it inside extra.
        extra = {
            "icon_style": icon_style,
            "value_type": type,
            "values": list(values or []),
            "show_value": True if showValue is None else showValue,
            "percent": percent,
            "reverse": reverse,
        }
        super().__init__(type="iconSet", extra=extra, **kw)


# ---------------------------------------------------------------------------
# Pod 2 (RFC-060 §2.4) — additional openpyxl-shaped names.
#
# openpyxl exposes ColorScale / DataBar / IconSet (no "Rule" suffix) as
# value classes that bundle the start/mid/end stops directly.  The
# corresponding ``ColorScaleRule`` / ``DataBarRule`` / ``IconSetRule``
# above already provide the same fields under the openpyxl-Rule alias;
# the un-suffixed names here are aliases so ``from
# openpyxl.formatting.rule import ColorScale`` swaps mechanically.
# ---------------------------------------------------------------------------

ColorScale = ColorScaleRule
DataBar = DataBarRule
IconSet = IconSetRule


class DifferentialStyle:
    """``<dxf>``-shaped formatting carried by CF rules' ``dxfId`` reference.

    A :class:`DifferentialStyle` is the per-rule formatting payload —
    ``font``, ``fill``, ``border``, ``alignment``, ``numFmt``.  Wolfxl
    stores this metadata inline on the :class:`Rule` ``extra`` blob
    (see :func:`wolfxl.formatting._dxf_from_rule`); this class is a
    thin construction shim that mirrors openpyxl's keyword surface so
    user code that builds a CF rule with explicit dxf state ports
    mechanically.

    Pod 2 (RFC-060 §2.4).
    """

    __slots__ = ("font", "fill", "border", "alignment", "number_format", "numFmt")

    def __init__(
        self,
        font: Any = None,
        fill: Any = None,
        border: Any = None,
        alignment: Any = None,
        number_format: Any = None,
        numFmt: Any = None,  # noqa: N803 — openpyxl alias
    ) -> None:
        self.font = font
        self.fill = fill
        self.border = border
        self.alignment = alignment
        # Accept either ``number_format`` (snake) or ``numFmt`` (camel).
        self.number_format = number_format if number_format is not None else numFmt
        self.numFmt = self.number_format


class RuleType:
    """Marker for parametrized CF rule kinds.

    Constants mirror openpyxl's ``RuleType`` enum so user code that
    references ``RuleType.COLOR_SCALE`` keeps working.

    Pod 2 (RFC-060 §2.4).
    """

    AVERAGE = "aboveAverage"
    COLOR_SCALE = "colorScale"
    DATA_BAR = "dataBar"
    ICON_SET = "iconSet"
    FORMULA = "expression"
    EXPRESSION = "expression"
    DUPLICATE_VALUES = "duplicateValues"
    UNIQUE_VALUES = "uniqueValues"
    CONTAINS_TEXT = "containsText"
    NOT_CONTAINS_TEXT = "notContainsText"
    BEGINS_WITH = "beginsWith"
    ENDS_WITH = "endsWith"
    CONTAINS_BLANKS = "containsBlanks"
    CONTAINS_NO_BLANKS = "notContainsBlanks"
    CONTAINS_ERRORS = "containsErrors"
    CONTAINS_NO_ERRORS = "notContainsErrors"
    TIME_PERIOD = "timePeriod"
    ABOVE_AVERAGE = "aboveAverage"
    TOP10 = "top10"
    CELL_IS = "cellIs"


__all__ = [
    "CellIsRule",
    "ColorScale",
    "ColorScaleRule",
    "DataBar",
    "DataBarRule",
    "DifferentialStyle",
    "FormulaRule",
    "IconSet",
    "IconSetRule",
    "Rule",
    "RuleType",
]
