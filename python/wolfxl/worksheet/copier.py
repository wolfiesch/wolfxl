"""``openpyxl.worksheet.copier`` — :class:`WorksheetCopy` value type.

Wolfxl exposes worksheet copying via :meth:`Workbook.copy_worksheet` and
the :class:`CopyOptions` value type.  openpyxl additionally exposes a
``WorksheetCopy(source, target)`` class with a ``.copy_worksheet()``
method; we surface a real wrapper here that delegates to the existing
RFC-035 deep-clone path for source-compatibility.

Sprint Π Pod-β (RFC-063) replaced the construction stub.
"""

from __future__ import annotations

from typing import Any


class WorksheetCopy:
    """Thin wrapper over :meth:`Workbook.copy_worksheet` (RFC-035).

    openpyxl exposes worksheet duplication as a class with a
    ``.copy_worksheet()`` method that runs the actual copy when called.
    Wolfxl users typically reach for ``wb.copy_worksheet(src)`` directly,
    but this class is provided so user code that constructs the openpyxl
    type by hand continues to work.

    Parameters
    ----------
    source:
        The source :class:`Worksheet` to clone.
    target:
        The destination :class:`Worksheet` whose ``title`` is used as
        the new sheet name.  The target need not already belong to the
        workbook — ``copy_worksheet`` only consults ``target.title``.

    Notes
    -----
    Calling :meth:`copy_worksheet` delegates to
    :meth:`Workbook.copy_worksheet` with ``name=target.title``.  If the
    target's title collides with an already-existing sheet, the
    workbook helper raises :class:`ValueError` (matching openpyxl
    behaviour).
    """

    __slots__ = ("source", "target")

    def __init__(self, source: Any, target: Any) -> None:
        self.source = source
        self.target = target

    def copy_worksheet(self) -> Any:
        """Run the deep clone and return the new sheet proxy."""
        wb = getattr(self.source, "_workbook", None)
        if wb is None:
            raise RuntimeError(
                "WorksheetCopy.copy_worksheet: source is not bound to a Workbook"
            )
        target_title = getattr(self.target, "title", None)
        if not isinstance(target_title, str) or not target_title:
            raise ValueError(
                "WorksheetCopy.copy_worksheet: target must expose a non-empty `title`"
            )
        return wb.copy_worksheet(self.source, name=target_title)

    def copy_cells(self) -> Any:
        """openpyxl alias for :meth:`copy_worksheet` (kept for symmetry).

        openpyxl's helper internally calls ``copy_worksheet`` to do the
        cell-level copy; surface the alias so user code that calls the
        method by either name keeps working.
        """
        return self.copy_worksheet()


__all__ = ["WorksheetCopy"]
