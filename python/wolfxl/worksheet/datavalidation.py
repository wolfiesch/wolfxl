"""openpyxl.worksheet.datavalidation compatibility.

``DataValidation`` is a real dataclass. ``DataValidationList`` mimics
openpyxl's container — the user iterates ``.dataValidation`` (a list)
and appends new DVs with ``.append(dv)``.

Append works in both write mode (``Workbook()`` → native writer) and
modify mode (``load_workbook(path, modify=True)`` → patcher). The
divergence happens at ``save()`` time, not at ``append()`` — both modes
queue onto ``ws._pending_data_validations`` here; the workbook's
``save()`` routes to the writer or the patcher accordingly.
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, ClassVar
from xml.etree import ElementTree as ET

from wolfxl.worksheet.cell_range import MultiCellRange

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


@dataclass(init=False)
class DataValidation:
    """A data-validation rule attached to one or more cell ranges.

    ``type`` is the validation kind: ``"list"``, ``"whole"``, ``"decimal"``,
    ``"date"``, ``"time"``, ``"textLength"``, ``"custom"``, or ``None``
    for the ``"any"`` fallback (no validation). ``operator`` qualifies
    numeric/date validations (``"between"``, ``"greaterThan"``, etc.).
    ``formula1`` / ``formula2`` hold the operands (openpyxl keeps the
    leading ``=``).

    ``sqref`` is the target range — we keep it as a plain string instead
    of openpyxl's ``MultiCellRange`` to avoid pulling that dependency.
    Space-separated multi-range strings round-trip as-is.
    """

    type: str | None = None
    operator: str | None = None
    formula1: str | None = None
    formula2: str | None = None
    allowBlank: bool = False  # noqa: N815 - openpyxl public API
    showErrorMessage: bool = False  # noqa: N815
    showInputMessage: bool = False  # noqa: N815
    error: str | None = None
    errorTitle: str | None = None  # noqa: N815
    prompt: str | None = None
    promptTitle: str | None = None  # noqa: N815
    sqref: str | MultiCellRange = ""
    showDropDown: bool = False  # noqa: N815
    errorStyle: str | None = None  # noqa: N815
    imeMode: str | None = None  # noqa: N815
    tagname: ClassVar[str] = "dataValidation"
    namespace: ClassVar[str | None] = None
    idx_base: ClassVar[int] = 0

    def __init__(
        self,
        type: str | None = None,  # noqa: A002 - openpyxl public API
        formula1: str | None = None,
        formula2: str | None = None,
        showErrorMessage: bool = False,  # noqa: N803
        showInputMessage: bool = False,  # noqa: N803
        showDropDown: bool = False,  # noqa: N803
        allowBlank: bool = False,  # noqa: N803
        sqref: str | MultiCellRange = "",
        promptTitle: str | None = None,  # noqa: N803
        errorStyle: str | None = None,  # noqa: N803
        error: str | None = None,
        prompt: str | None = None,
        errorTitle: str | None = None,  # noqa: N803
        imeMode: str | None = None,  # noqa: N803
        operator: str | None = None,
        allow_blank: bool | None = None,
    ) -> None:
        self.type = type
        self.operator = operator
        self.formula1 = formula1
        self.formula2 = formula2
        self.allowBlank = allowBlank if allow_blank is None else bool(allow_blank)
        self.showErrorMessage = showErrorMessage
        self.showInputMessage = showInputMessage
        self.error = error
        self.errorTitle = errorTitle
        self.prompt = prompt
        self.promptTitle = promptTitle
        self.sqref = sqref
        self.showDropDown = showDropDown
        self.errorStyle = errorStyle
        self.imeMode = imeMode
        self.__post_init__()

    def __post_init__(self) -> None:
        if not isinstance(self.sqref, MultiCellRange):
            self.sqref = MultiCellRange(str(self.sqref)) if self.sqref else MultiCellRange()

    @property
    def validation_type(self) -> str | None:
        return self.type

    @validation_type.setter
    def validation_type(self, value: str | None) -> None:
        self.type = value

    @property
    def allow_blank(self) -> bool:
        return self.allowBlank

    @allow_blank.setter
    def allow_blank(self, value: bool) -> None:
        self.allowBlank = bool(value)

    @property
    def hide_drop_down(self) -> bool:
        return self.showDropDown

    @hide_drop_down.setter
    def hide_drop_down(self, value: bool) -> None:
        self.showDropDown = bool(value)

    @property
    def ranges(self) -> MultiCellRange:
        return self.sqref

    @ranges.setter
    def ranges(self, value: str | MultiCellRange) -> None:
        self.sqref = value if isinstance(value, MultiCellRange) else MultiCellRange(value)

    @property
    def cells(self) -> MultiCellRange:
        return self.sqref

    def add(self, cell: str) -> None:
        """Attach this validation to a cell or range."""
        self.sqref.add(cell)

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002 - openpyxl signature
        namespace: str | None = None,  # noqa: ARG002 - openpyxl signature
    ) -> ET.Element:
        node = ET.Element(tagname or self.tagname)
        attrs: dict[str, str | None] = {
            "sqref": str(self.sqref),
            "showDropDown": "1" if self.showDropDown else "0",
            "showInputMessage": "1" if self.showInputMessage else "0",
            "showErrorMessage": "1" if self.showErrorMessage else "0",
            "allowBlank": "1" if self.allowBlank else "0",
            "type": self.type,
            "operator": self.operator,
            "errorStyle": self.errorStyle,
            "imeMode": self.imeMode,
            "error": self.error,
            "errorTitle": self.errorTitle,
            "prompt": self.prompt,
            "promptTitle": self.promptTitle,
        }
        for key, value in attrs.items():
            if value is not None:
                node.set(key, value)
        if self.formula1 is not None:
            ET.SubElement(node, "formula1").text = self.formula1
        if self.formula2 is not None:
            ET.SubElement(node, "formula2").text = self.formula2
        return node

    @classmethod
    def from_tree(cls, node: ET.Element) -> DataValidation:
        attrs = node.attrib
        formula1 = formula2 = None
        for child in node:
            if child.tag == "formula1":
                formula1 = child.text
            elif child.tag == "formula2":
                formula2 = child.text
        return cls(
            type=attrs.get("type"),
            operator=attrs.get("operator"),
            formula1=formula1,
            formula2=formula2,
            allowBlank=attrs.get("allowBlank", "0") not in {"0", "false", "False"},
            showErrorMessage=attrs.get("showErrorMessage", "0")
            not in {"0", "false", "False"},
            showInputMessage=attrs.get("showInputMessage", "0")
            not in {"0", "false", "False"},
            error=attrs.get("error"),
            errorTitle=attrs.get("errorTitle"),
            prompt=attrs.get("prompt"),
            promptTitle=attrs.get("promptTitle"),
            sqref=attrs.get("sqref", ""),
            showDropDown=attrs.get("showDropDown", "0") not in {"0", "false", "False"},
            errorStyle=attrs.get("errorStyle"),
            imeMode=attrs.get("imeMode"),
        )


class DataValidationList:
    """Container for a worksheet's data validations.

    openpyxl users write ``ws.data_validations.dataValidation`` to get a
    list of DV objects and ``ws.data_validations.append(dv)`` to attach a
    new one. We mimic that exactly.

    The ``_ws`` back-reference lets ``append()`` queue writes through to
    the Rust writer in write mode (PR5). In read mode, ``_ws`` is set
    but ``append()`` raises with a T1.5 pointer — the current plan
    doesn't handle appending DVs to existing files.
    """

    tagname: ClassVar[str] = "dataValidations"
    namespace: ClassVar[str | None] = None
    idx_base: ClassVar[int] = 0

    __slots__ = ("dataValidation", "_ws", "disablePrompts", "xWindow", "yWindow")

    def __init__(
        self,
        dvs: list[DataValidation] | None = None,
        ws: Worksheet | None = None,
        disablePrompts: bool | None = None,  # noqa: N803
        xWindow: int | None = None,  # noqa: N803
        yWindow: int | None = None,  # noqa: N803
    ) -> None:
        # openpyxl calls this ``dataValidation`` (singular list name).
        self.dataValidation = list(dvs or [])  # noqa: N815
        self._ws = ws
        self.disablePrompts = disablePrompts  # noqa: N815
        self.xWindow = xWindow  # noqa: N815
        self.yWindow = yWindow  # noqa: N815

    def __iter__(self) -> Iterator[DataValidation]:
        return iter(self.dataValidation)

    def __len__(self) -> int:
        return len(self.dataValidation)

    def __contains__(self, item: Any) -> bool:
        return item in self.dataValidation

    @property
    def count(self) -> int:
        return len(self.dataValidation)

    def append(self, dv: DataValidation) -> None:
        """Attach a new data validation.

        Works in write mode and modify mode. Pure read mode (no writer,
        no patcher) raises — there is no path for the change to flow to
        disk. Both writable modes queue here; ``Workbook.save`` routes
        to the right backend.
        """
        self.dataValidation.append(dv)
        ws = self._ws
        if ws is None:
            return
        wb = ws._workbook  # noqa: SLF001
        if wb._rust_writer is None and wb._rust_patcher is None:  # noqa: SLF001
            raise RuntimeError("DataValidationList is not attached to a workbook")
        ws._pending_data_validations.append(dv)  # noqa: SLF001

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002 - openpyxl signature
        namespace: str | None = None,  # noqa: ARG002 - openpyxl signature
    ) -> ET.Element:
        node = ET.Element(tagname or self.tagname)
        attrs = {
            "disablePrompts": self.disablePrompts,
            "xWindow": self.xWindow,
            "yWindow": self.yWindow,
        }
        for key, value in attrs.items():
            if value is not None:
                node.set(key, "1" if value is True else "0" if value is False else str(value))
        children = [dv for dv in self.dataValidation if str(dv.sqref)]
        node.set("count", str(len(children)))
        for dv in children:
            node.append(dv.to_tree())
        return node

    @classmethod
    def from_tree(cls, node: ET.Element) -> DataValidationList:
        attrs = node.attrib
        return cls(
            dvs=[DataValidation.from_tree(child) for child in node if child.tag == "dataValidation"],
            disablePrompts=attrs.get("disablePrompts", "0") not in {"0", "false", "False"}
            if "disablePrompts" in attrs
            else None,
            xWindow=int(attrs["xWindow"]) if "xWindow" in attrs else None,
            yWindow=int(attrs["yWindow"]) if "yWindow" in attrs else None,
        )


def _dv_to_patcher_dict(dv: DataValidation) -> dict[str, Any]:
    """Convert an openpyxl-shaped ``DataValidation`` into the patcher's payload.

    The patcher's PyO3 method (``XlsxPatcher.queue_data_validation``)
    accepts a flat dict whose key names match RFC-025 §4.2's
    ``DataValidationPatch`` fields. Unknown keys are ignored on the
    Rust side, so this helper can over-supply safely; we surface the
    keys the patcher actually consumes.

    ``showDropDown`` and ``errorStyle`` aren't on ``DataValidation``
    today (see field list at lines 35-46). They're sent as defaults so
    the patcher round-trip stays predictable; if either field is added
    to the dataclass later, the helper picks them up automatically via
    ``getattr``.
    """
    # Bools and the required ``sqref`` always go in. Optional strings are
    # only included if non-None — the patcher's PyO3 layer expects either
    # a real string or absence (Option::None on the Rust side); a Python
    # ``None`` value would fail the ``.extract::<String>()`` cast.
    payload: dict[str, Any] = {
        "sqref": str(dv.sqref),
        "validation_type": dv.type or "none",
        "allow_blank": bool(dv.allowBlank),
        "show_dropdown": bool(getattr(dv, "showDropDown", False)),
        "show_input_message": bool(dv.showInputMessage),
        "show_error_message": bool(dv.showErrorMessage),
    }
    optional = {
        "operator": dv.operator,
        "formula1": dv.formula1,
        "formula2": dv.formula2,
        "error_style": getattr(dv, "errorStyle", None),
        "error_title": dv.errorTitle,
        "error": dv.error,
        "prompt_title": dv.promptTitle,
        "prompt": dv.prompt,
    }
    payload.update({k: v for k, v in optional.items() if v is not None})
    return payload


__all__ = ["DataValidation", "DataValidationList", "_dv_to_patcher_dict"]
