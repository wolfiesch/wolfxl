"""Workbook — multi-mode openpyxl-compatible wrapper.

Write mode (``Workbook()``): creates a new workbook via RustXlsxWriterBook.
Read mode (``Workbook._from_reader(path)``): opens an existing .xlsx via CalamineStyledBook.
Modify mode (``Workbook._from_patcher(path)``): read via CalamineStyledBook, save via XlsxPatcher.
"""

from __future__ import annotations

import os
from typing import Any

from wolfxl._worksheet import Worksheet


class Workbook:
    """openpyxl-compatible workbook backed by Rust."""

    def __init__(self) -> None:
        """Create a new workbook in write mode with a default 'Sheet'."""
        from wolfxl import _rust

        self._rust_writer: Any = _rust.RustXlsxWriterBook()
        self._rust_reader: Any = None
        self._rust_patcher: Any = None
        self._sheet_names: list[str] = ["Sheet"]
        self._sheets: dict[str, Worksheet] = {}
        self._sheets["Sheet"] = Worksheet(self, "Sheet")
        self._rust_writer.add_sheet("Sheet")

    @classmethod
    def _from_reader(cls, path: str) -> Workbook:
        """Open an existing .xlsx file in read mode."""
        from wolfxl import _rust

        wb = object.__new__(cls)
        wb._rust_writer = None
        wb._rust_patcher = None
        wb._rust_reader = _rust.CalamineStyledBook.open(path)
        names = [str(n) for n in wb._rust_reader.sheet_names()]
        wb._sheet_names = names
        wb._sheets = {}
        for name in names:
            wb._sheets[name] = Worksheet(wb, name)
        return wb

    @classmethod
    def _from_patcher(cls, path: str) -> Workbook:
        """Open an existing .xlsx file in modify mode (read + surgical save)."""
        from wolfxl import _rust

        wb = object.__new__(cls)
        wb._rust_writer = None
        wb._rust_reader = _rust.CalamineStyledBook.open(path)
        wb._rust_patcher = _rust.XlsxPatcher.open(path)
        names = [str(n) for n in wb._rust_reader.sheet_names()]
        wb._sheet_names = names
        wb._sheets = {}
        for name in names:
            wb._sheets[name] = Worksheet(wb, name)
        return wb

    # ------------------------------------------------------------------
    # Sheet access
    # ------------------------------------------------------------------

    @property
    def sheetnames(self) -> list[str]:
        return list(self._sheet_names)

    @property
    def active(self) -> Worksheet | None:
        """Return the first sheet, or None if no sheets exist."""
        if self._sheet_names:
            return self._sheets[self._sheet_names[0]]
        return None

    def __getitem__(self, name: str) -> Worksheet:
        if name not in self._sheets:
            raise KeyError(f"Worksheet '{name}' does not exist")
        return self._sheets[name]

    def __contains__(self, name: str) -> bool:
        return name in self._sheets

    def __iter__(self):  # type: ignore[no-untyped-def]
        return iter(self._sheet_names)

    # ------------------------------------------------------------------
    # Write-mode operations
    # ------------------------------------------------------------------

    def create_sheet(self, title: str) -> Worksheet:
        """Add a new sheet (write mode only)."""
        if self._rust_writer is None:
            raise RuntimeError("create_sheet requires write mode")
        if title in self._sheets:
            raise ValueError(f"Sheet '{title}' already exists")
        self._rust_writer.add_sheet(title)
        self._sheet_names.append(title)
        ws = Worksheet(self, title)
        self._sheets[title] = ws
        return ws

    def save(self, filename: str | os.PathLike[str]) -> None:
        """Flush all pending writes and save to disk."""
        filename = str(filename)
        if self._rust_patcher is not None:
            # Modify mode — flush to patcher, then surgical save.
            for ws in self._sheets.values():
                ws._flush()  # noqa: SLF001
            self._rust_patcher.save(filename)
        elif self._rust_writer is not None:
            # Write mode — flush to writer, then full save.
            for ws in self._sheets.values():
                ws._flush()  # noqa: SLF001
            self._rust_writer.save(filename)
        else:
            raise RuntimeError("save requires write or modify mode")

    # ------------------------------------------------------------------
    # Context manager + cleanup
    # ------------------------------------------------------------------

    def close(self) -> None:
        """Release resources."""
        self._rust_reader = None
        self._rust_writer = None
        self._rust_patcher = None

    def __enter__(self) -> Workbook:
        return self

    def __exit__(self, *args: object) -> None:
        self.close()

    def __repr__(self) -> str:
        if self._rust_patcher is not None:
            mode = "modify"
        elif self._rust_reader is not None:
            mode = "read"
        else:
            mode = "write"
        return f"<Workbook [{mode}] sheets={self._sheet_names}>"
