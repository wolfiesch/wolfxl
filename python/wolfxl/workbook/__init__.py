"""openpyxl.workbook compatibility.

Exposes :class:`DefinedNameDict` — the openpyxl-shape container for
``wb.defined_names``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from wolfxl.workbook.defined_name import DefinedName

if TYPE_CHECKING:
    from wolfxl._workbook import Workbook


class DefinedNameDict(dict):
    """``dict``-subclass container for workbook defined names.

    Looks like a dict, but enforces that values are
    :class:`DefinedName` objects and that the key matches the DN's
    ``.name``. Having a subclass (rather than a plain dict) also lets
    write-mode callers queue a Rust-side flush on ``__setitem__`` in
    PR6 — the ``_wb`` back-ref is set by ``Workbook.defined_names``.
    """

    def __init__(self, *args: object, **kwargs: object) -> None:
        super().__init__(*args, **kwargs)
        # The workbook attaches itself after construction so the dict's
        # superclass init can stay plain. ``_wb`` is used by the PR6
        # write path; PR3 reads only need the container behavior.
        self._wb: Workbook | None = None

    def __setitem__(self, key: str, value: DefinedName) -> None:
        if not isinstance(value, DefinedName):
            raise TypeError(
                f"value must be DefinedName, got {type(value).__name__}"
            )
        if key != value.name:
            raise ValueError(
                f"key '{key}' does not match DefinedName.name '{value.name}'"
            )
        super().__setitem__(key, value)
        # Queue a flush in both write and modify mode. Write mode ships
        # via ``_rust_writer.add_named_range`` during ``Workbook.save()``;
        # modify mode raises in ``save()`` with a T1.5 pointer — but we
        # still need to know a write was attempted, hence the queue.
        wb = self._wb
        if wb is not None:
            wb._pending_defined_names[key] = value  # noqa: SLF001

    def add(self, value: DefinedName) -> None:
        """openpyxl helper: ``dn.add(DefinedName(...))`` is common."""
        self[value.name] = value


__all__ = ["DefinedNameDict", "DefinedName"]
