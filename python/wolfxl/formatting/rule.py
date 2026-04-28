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


@dataclass
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
    # ``color_scale`` / ``data_bar`` / ``icon_set`` metadata blobs are
    # preserved on round-trip but not decomposed here — T2 territory.
    extra: dict[str, Any] = field(default_factory=dict)


class CellIsRule(Rule):
    """Conditional format triggered when a cell value matches an operator+operand.

    Example: ``CellIsRule(operator="greaterThan", formula=["50"])``.
    """

    def __init__(
        self,
        operator: str | None = None,
        formula: list[str] | None = None,
        stopIfTrue: bool = False,  # noqa: N803
        **kw: Any,
    ) -> None:
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
    """

    def __init__(
        self,
        formula: list[str] | None = None,
        stopIfTrue: bool = False,  # noqa: N803
        **kw: Any,
    ) -> None:
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
        showValue: bool = True,  # noqa: N803
        **kw: Any,
    ) -> None:
        extra = {
            "start_type": start_type,
            "start_value": start_value,
            "end_type": end_type,
            "end_value": end_value,
            "color": color,
            "show_value": showValue,
        }
        super().__init__(type="dataBar", extra=extra, **kw)


class IconSetRule(Rule):
    """Icon set (3 arrows, 5 traffic lights, etc.)."""

    def __init__(
        self,
        icon_style: str | None = None,
        type: str | None = None,  # noqa: A002 - keyword openpyxl uses
        values: list[Any] | None = None,
        showValue: bool = True,  # noqa: N803
        **kw: Any,
    ) -> None:
        # openpyxl's IconSetRule positional ``type`` ("percent", "percentile",
        # "num", "formula") differs from Rule.type — we stash it inside extra.
        extra = {
            "icon_style": icon_style,
            "value_type": type,
            "values": list(values or []),
            "show_value": showValue,
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
