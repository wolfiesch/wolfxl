"""openpyxl.worksheet.datavalidation compatibility.

``DataValidation`` is a real dataclass. ``DataValidationList`` mimics
openpyxl's container — the user iterates ``.dataValidation`` (a list)
and appends new DVs with ``.append(dv)``. Reads land in PR2, writes
(append in write mode) in PR5.
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

        Wired up in T1 PR5. The import is lazy so construction of this
        container in read mode (where ``append`` just raises) doesn't
        force loading the writer helpers.
        """
        ws = self._ws
        if ws is None:
            raise RuntimeError("DataValidationList is not attached to a worksheet")
        wb = ws._workbook  # noqa: SLF001
        if wb._rust_writer is None:  # noqa: SLF001
            raise NotImplementedError(
                "Appending data validations to existing files is a T1.5 follow-up. "
                "Write mode (Workbook() + save) is supported."
            )
        self.dataValidation.append(dv)
        ws._pending_data_validations.append(dv)  # noqa: SLF001


__all__ = ["DataValidation", "DataValidationList"]
