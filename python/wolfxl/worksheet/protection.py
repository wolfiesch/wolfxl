"""Sheet protection (RFC-055 §2.6).

Backs ``ws.protection``. Mirrors openpyxl's
``openpyxl.worksheet.protection.SheetProtection`` field surface
including the ``set_password`` / ``check_password`` helpers.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from wolfxl.utils.protection import hash_password as _hash_password
from wolfxl.utils.protection import check_password as _check_password


@dataclass
class SheetProtection:
    """SheetProtection (CT_SheetProtection, ECMA-376 §18.3.1.85).

    Defaults match Excel's behaviour when ``Tools → Protection → Protect Sheet``
    is invoked with no further options: the sheet is locked but every "allow
    these actions" toggle defaults to allowed (because the sheet's underlying
    ``locked`` cell-level flag still gates each operation).
    """

    sheet: bool = False
    objects: bool = False
    scenarios: bool = False
    formatCells: bool = True  # noqa: N815
    formatColumns: bool = True  # noqa: N815
    formatRows: bool = True  # noqa: N815
    insertColumns: bool = True  # noqa: N815
    insertRows: bool = True  # noqa: N815
    insertHyperlinks: bool = True  # noqa: N815
    deleteColumns: bool = True  # noqa: N815
    deleteRows: bool = True  # noqa: N815
    selectLockedCells: bool = False  # noqa: N815
    sort: bool = True
    autoFilter: bool = True  # noqa: N815
    pivotTables: bool = True  # noqa: N815
    selectUnlockedCells: bool = False  # noqa: N815
    password: str | None = None  # already-hashed (4-char uppercase hex)

    # snake_case aliases (the canonical openpyxl attr names use camelCase
    # but several call sites in the wider Python ecosystem use snake_case).
    @property
    def format_cells(self) -> bool:
        """Return whether users may format cells on a protected sheet.

        Returns:
            ``True`` when cell-formatting operations are allowed.
        """
        return self.formatCells

    @format_cells.setter
    def format_cells(self, value: bool) -> None:
        """Set whether users may format cells on a protected sheet.

        Args:
            value: Truthy value to allow cell-formatting operations.
        """
        self.formatCells = bool(value)

    @property
    def format_columns(self) -> bool:
        return self.formatColumns

    @format_columns.setter
    def format_columns(self, value: bool) -> None:
        self.formatColumns = bool(value)

    @property
    def format_rows(self) -> bool:
        return self.formatRows

    @format_rows.setter
    def format_rows(self, value: bool) -> None:
        self.formatRows = bool(value)

    @property
    def insert_columns(self) -> bool:
        return self.insertColumns

    @insert_columns.setter
    def insert_columns(self, value: bool) -> None:
        self.insertColumns = bool(value)

    @property
    def insert_rows(self) -> bool:
        return self.insertRows

    @insert_rows.setter
    def insert_rows(self, value: bool) -> None:
        self.insertRows = bool(value)

    @property
    def insert_hyperlinks(self) -> bool:
        return self.insertHyperlinks

    @insert_hyperlinks.setter
    def insert_hyperlinks(self, value: bool) -> None:
        self.insertHyperlinks = bool(value)

    @property
    def delete_columns(self) -> bool:
        return self.deleteColumns

    @delete_columns.setter
    def delete_columns(self, value: bool) -> None:
        self.deleteColumns = bool(value)

    @property
    def delete_rows(self) -> bool:
        return self.deleteRows

    @delete_rows.setter
    def delete_rows(self, value: bool) -> None:
        self.deleteRows = bool(value)

    @property
    def select_locked_cells(self) -> bool:
        return self.selectLockedCells

    @select_locked_cells.setter
    def select_locked_cells(self, value: bool) -> None:
        self.selectLockedCells = bool(value)

    @property
    def auto_filter(self) -> bool:
        return self.autoFilter

    @auto_filter.setter
    def auto_filter(self, value: bool) -> None:
        self.autoFilter = bool(value)

    @property
    def pivot_tables(self) -> bool:
        return self.pivotTables

    @pivot_tables.setter
    def pivot_tables(self, value: bool) -> None:
        self.pivotTables = bool(value)

    @property
    def select_unlocked_cells(self) -> bool:
        return self.selectUnlockedCells

    @select_unlocked_cells.setter
    def select_unlocked_cells(self, value: bool) -> None:
        self.selectUnlockedCells = bool(value)

    def enable(self) -> None:
        """Turn protection on. ``ws.protection.enable()`` is the single-call
        idiom most users reach for after constructing the SheetProtection.
        """
        self.sheet = True

    def disable(self) -> None:
        self.sheet = False
        self.password = None

    def set_password(self, plaintext: str) -> None:
        """Hash ``plaintext`` and store the 4-hex result in ``password``.

        Empty / None plaintext clears the password. Setting a password
        does NOT automatically enable protection — call ``enable()`` or
        set ``sheet = True`` separately so the OOXML emit knows to
        write the ``<sheetProtection>`` block.
        """
        if not plaintext:
            self.password = None
            return
        self.password = _hash_password(plaintext)

    def check_password(self, plaintext: str) -> bool:
        if self.password is None:
            return False
        return _check_password(plaintext, self.password)

    def to_rust_dict(self) -> dict[str, Any]:
        return {
            "sheet": bool(self.sheet),
            "objects": bool(self.objects),
            "scenarios": bool(self.scenarios),
            "format_cells": bool(self.formatCells),
            "format_columns": bool(self.formatColumns),
            "format_rows": bool(self.formatRows),
            "insert_columns": bool(self.insertColumns),
            "insert_rows": bool(self.insertRows),
            "insert_hyperlinks": bool(self.insertHyperlinks),
            "delete_columns": bool(self.deleteColumns),
            "delete_rows": bool(self.deleteRows),
            "select_locked_cells": bool(self.selectLockedCells),
            "sort": bool(self.sort),
            "auto_filter": bool(self.autoFilter),
            "pivot_tables": bool(self.pivotTables),
            "select_unlocked_cells": bool(self.selectUnlockedCells),
            "password_hash": self.password,
        }

    def is_default(self) -> bool:
        return self == SheetProtection()


__all__ = ["SheetProtection"]
