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


__all__ = [
    "CellIsRule",
    "ColorScaleRule",
    "DataBarRule",
    "FormulaRule",
    "IconSetRule",
    "Rule",
]
