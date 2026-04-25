"""openpyxl.workbook.defined_name compatibility.

T1 makes ``DefinedName`` a real dataclass.  Read access comes through
``wb.defined_names`` — that returns a :class:`DefinedNameDict` whose
values are ``DefinedName`` objects instead of bare strings.

Breaking change from T0: callers that did ``wb.defined_names["X"]`` and
expected a string must now do ``wb.defined_names["X"].value`` (or
``.attr_text`` for openpyxl parity). See CHANGELOG for the 1-line
migration.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class DefinedName:
    """A workbook-scoped or sheet-scoped name range.

    ``value`` holds the ``refers_to`` expression (``Sheet1!$A$1:$A$10``
    or an external reference). ``localSheetId`` is ``None`` for
    workbook-scoped names or the 0-based sheet index for sheet-scoped
    ones. ``hidden=True`` marks internal names Excel uses for print
    areas and table ranges.
    """

    name: str
    value: str
    comment: str | None = None
    localSheetId: int | None = None  # noqa: N815 - openpyxl public API
    hidden: bool = False

    @property
    def attr_text(self) -> str:
        """openpyxl alias for ``.value``."""
        return self.value


__all__ = ["DefinedName"]
