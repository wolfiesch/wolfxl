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
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


@dataclass
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
    sqref: str = ""


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

    __slots__ = ("dataValidation", "_ws")

    def __init__(
        self,
        dvs: list[DataValidation] | None = None,
        ws: Worksheet | None = None,
    ) -> None:
        # openpyxl calls this ``dataValidation`` (singular list name).
        self.dataValidation = list(dvs or [])  # noqa: N815
        self._ws = ws

    def __iter__(self) -> Iterator[DataValidation]:
        return iter(self.dataValidation)

    def __len__(self) -> int:
        return len(self.dataValidation)

    def __contains__(self, item: Any) -> bool:
        return item in self.dataValidation

    def append(self, dv: DataValidation) -> None:
        """Attach a new data validation.

        Works in write mode and modify mode. Pure read mode (no writer,
        no patcher) raises — there is no path for the change to flow to
        disk. Both writable modes queue here; ``Workbook.save`` routes
        to the right backend.
        """
        ws = self._ws
        if ws is None:
            raise RuntimeError("DataValidationList is not attached to a worksheet")
        wb = ws._workbook  # noqa: SLF001
        if wb._rust_writer is None and wb._rust_patcher is None:  # noqa: SLF001
            raise RuntimeError("DataValidationList is not attached to a workbook")
        self.dataValidation.append(dv)
        ws._pending_data_validations.append(dv)  # noqa: SLF001


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
    return {
        "sqref": dv.sqref,
        "validation_type": dv.type or "none",
        "operator": dv.operator,
        "formula1": dv.formula1,
        "formula2": dv.formula2,
        "allow_blank": bool(dv.allowBlank),
        "show_dropdown": bool(getattr(dv, "showDropDown", False)),
        "show_input_message": bool(dv.showInputMessage),
        "show_error_message": bool(dv.showErrorMessage),
        "error_style": getattr(dv, "errorStyle", None),
        "error_title": dv.errorTitle,
        "error": dv.error,
        "prompt_title": dv.promptTitle,
        "prompt": dv.prompt,
    }


__all__ = ["DataValidation", "DataValidationList", "_dv_to_patcher_dict"]
