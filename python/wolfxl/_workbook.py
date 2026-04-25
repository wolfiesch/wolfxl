"""Workbook — multi-mode openpyxl-compatible wrapper.

Write mode (``Workbook()``): creates a new workbook via RustXlsxWriterBook.
Read mode (``Workbook._from_reader(path)``): opens an existing .xlsx via CalamineStyledBook.
Modify mode (``Workbook._from_patcher(path)``): read via CalamineStyledBook, save via XlsxPatcher.
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, Any

from wolfxl._worksheet import Worksheet

if TYPE_CHECKING:
    from wolfxl.calc._protocol import RecalcResult


class Workbook:
    """openpyxl-compatible workbook backed by Rust."""

    def __init__(self) -> None:
        """Create a new workbook in write mode with a default 'Sheet'."""
        from wolfxl import _backend, _rust  # noqa: F401  (_rust kept for typing parity)

        self._rust_writer: Any = _backend.make_writer()
        self._rust_reader: Any = None
        self._rust_patcher: Any = None
        self._data_only = False
        self._evaluator: Any = None
        self._sheet_names: list[str] = ["Sheet"]
        self._sheets: dict[str, Worksheet] = {}
        self._sheets["Sheet"] = Worksheet(self, "Sheet")
        self._rust_writer.add_sheet("Sheet")

    @classmethod
    def _from_reader(cls, path: str, *, data_only: bool = False) -> Workbook:
        """Open an existing .xlsx file in read mode."""
        from wolfxl import _rust

        wb = object.__new__(cls)
        wb._rust_writer = None
        wb._rust_patcher = None
        wb._data_only = data_only
        wb._evaluator = None
        wb._rust_reader = _rust.CalamineStyledBook.open(path)
        names = [str(n) for n in wb._rust_reader.sheet_names()]
        wb._sheet_names = names
        wb._sheets = {}
        for name in names:
            wb._sheets[name] = Worksheet(wb, name)
        return wb

    @classmethod
    def _from_patcher(cls, path: str, *, data_only: bool = False) -> Workbook:
        """Open an existing .xlsx file in modify mode (read + surgical save)."""
        from wolfxl import _rust

        wb = object.__new__(cls)
        wb._rust_writer = None
        wb._data_only = data_only
        wb._evaluator = None
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
    # Named ranges
    # ------------------------------------------------------------------

    @property
    def defined_names(self) -> dict[str, str]:
        """Return all defined names as ``{NAME: refers_to}`` (case-preserved).

        Reads from the Rust backend (read mode) or returns empty (write mode).
        Workbook-scoped names override sheet-scoped on collision.
        """
        if self._rust_reader is None:
            return {}
        result: dict[str, str] = {}
        for sheet_name in self._sheet_names:
            try:
                entries = self._rust_reader.read_named_ranges(sheet_name)
            except Exception:
                continue
            for entry in entries:
                name = entry["name"]
                refers_to = entry["refers_to"]
                # Strip leading '=' if present
                if refers_to.startswith("="):
                    refers_to = refers_to[1:]
                result[name] = refers_to
        return result

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

    def copy_worksheet(self, source: Worksheet) -> Worksheet:
        """Duplicate *source* into a new sheet within this workbook.

        Tracked by RFC-035 (Phase 4 / WolfXL 1.1). See
        ``Plans/rfcs/035-copy-worksheet.md`` for the implementation plan.
        """
        raise NotImplementedError(
            "Workbook.copy_worksheet is scheduled for WolfXL 1.1 (RFC-035). "
            "See Plans/rfcs/035-copy-worksheet.md for the implementation plan. "
            "Workaround: use openpyxl for structural ops, then load the result "
            "with wolfxl.load_workbook() to do the heavy reads."
        )

    def move_sheet(self, sheet: Worksheet | str, offset: int = 0) -> None:
        """Move *sheet* by *offset* positions within the sheet order.

        Tracked by RFC-036 (Phase 4 / WolfXL 1.1). See
        ``Plans/rfcs/036-move-sheet.md`` for the implementation plan.
        """
        raise NotImplementedError(
            "Workbook.move_sheet is scheduled for WolfXL 1.1 (RFC-036). "
            "See Plans/rfcs/036-move-sheet.md for the implementation plan. "
            "Workaround: use openpyxl for structural ops, then load the result "
            "with wolfxl.load_workbook() to do the heavy reads."
        )

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
    # Formula evaluation (requires wolfxl.calc)
    # ------------------------------------------------------------------

    def calculate(self) -> dict[str, Any]:
        """Evaluate all formulas in the workbook.

        Returns a dict of cell_ref -> computed value for all formula cells.
        Requires the ``wolfxl.calc`` module (install via ``pip install wolfxl[calc]``).

        The internal evaluator is cached so that a subsequent
        :meth:`recalculate` call can reuse it without rescanning.
        """
        from wolfxl.calc._evaluator import WorkbookEvaluator

        ev = WorkbookEvaluator()
        ev.load(self)
        result = ev.calculate()
        self._evaluator = ev  # cache for recalculate()
        return result

    def cached_formula_values(self) -> dict[str, Any]:
        """Return Excel-saved cached formula results for every sheet.

        Keys are workbook-qualified cell references like ``"Sheet1!B2"``.
        This is a fast read-only path for ingestion workloads that need
        Excel's last-calculated formula values without evaluating formulas in
        Python. Cells whose formulas have no cached value are omitted.
        """
        if self._rust_reader is None:
            return {}
        values: dict[str, Any] = {}
        for sheet_name in self._sheet_names:
            values.update(self._sheets[sheet_name].cached_formula_values(qualified=True))
        return values

    def recalculate(
        self,
        perturbations: dict[str, float | int],
        tolerance: float = 1e-10,
    ) -> RecalcResult:
        """Perturb input cells and recompute affected formulas.

        Returns a ``RecalcResult`` describing which cells changed.
        Requires the ``wolfxl.calc`` module.

        If :meth:`calculate` was called first, the cached evaluator is
        reused (avoiding a full rescan + recalculate).
        """
        ev = self._evaluator
        if ev is None:
            from wolfxl.calc._evaluator import WorkbookEvaluator

            ev = WorkbookEvaluator()
            ev.load(self)
            ev.calculate()
            self._evaluator = ev
        return ev.recalculate(perturbations, tolerance)

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
