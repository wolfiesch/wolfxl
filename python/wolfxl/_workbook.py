"""Workbook — multi-mode openpyxl-compatible wrapper.

Write mode (``Workbook()``): creates a new workbook via NativeWorkbook.
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
        # T1 PR3 — workbook-level metadata + defined names.
        self._properties_cache: Any | None = None
        self._properties_dirty: bool = False
        self._defined_names_cache: Any | None = None
        self._pending_defined_names: dict[str, Any] = {}

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
        wb._properties_cache = None
        wb._properties_dirty = False
        wb._defined_names_cache = None
        wb._pending_defined_names = {}
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
        wb._properties_cache = None
        wb._properties_dirty = False
        wb._defined_names_cache = None
        wb._pending_defined_names = {}
        return wb

    # ------------------------------------------------------------------
    # Sheet access
    # ------------------------------------------------------------------

    @property
    def sheetnames(self) -> list[str]:
        return list(self._sheet_names)

    @property
    def worksheets(self) -> list[Worksheet]:
        """List of Worksheet objects in sheet order — openpyxl alias."""
        return [self._sheets[name] for name in self._sheet_names]

    @property
    def active(self) -> Worksheet | None:
        """Return the first sheet, or None if no sheets exist."""
        if self._sheet_names:
            return self._sheets[self._sheet_names[0]]
        return None

    @property
    def read_only(self) -> bool:
        """True if this workbook was opened read-only (no writer, no patcher)."""
        return self._rust_reader is not None and self._rust_patcher is None

    @property
    def chartsheets(self) -> list[Any]:
        """Chart sheets - always empty in T0 (wolfxl treats charts as preserved-only)."""
        return []

    @property
    def named_styles(self) -> list[Any]:
        """Named styles - always empty in T0 (construction lands in T2)."""
        return []

    def __getitem__(self, name: str) -> Worksheet:
        if name not in self._sheets:
            raise KeyError(f"Worksheet '{name}' does not exist")
        return self._sheets[name]

    def __contains__(self, name: str) -> bool:
        return name in self._sheets

    def __iter__(self):  # type: ignore[no-untyped-def]
        return iter(self._sheet_names)

    def get_sheet_by_name(self, name: str) -> Worksheet:
        """Look up a sheet by name. Deprecated in openpyxl but still widely used."""
        return self[name]

    def index(self, worksheet: Worksheet) -> int:
        """Return the 0-based index of ``worksheet`` in sheet order."""
        return self._sheet_names.index(worksheet.title)

    def remove(self, worksheet: Worksheet) -> None:
        """Remove a worksheet from the workbook (write mode only).

        In read mode, the on-disk sheet is untouched — raise instead so
        callers don't assume a destructive edit succeeded. Modify mode does
        not yet support sheet removal (the patcher has no ``remove_sheet``
        API surface), so it also raises.
        """
        if self._rust_writer is None:
            raise RuntimeError("remove requires write mode")
        if worksheet.title not in self._sheets:
            raise ValueError(f"Worksheet '{worksheet.title}' is not in this workbook")
        title = worksheet.title
        self._sheet_names.remove(title)
        self._sheets.pop(title)
        # If the Rust writer exposes remove_sheet, call it so the saved file
        # doesn't include the now-dropped sheet. If the writer lacks the
        # method, the Python bookkeeping still produces the right output
        # because ``save()`` iterates our ``_sheets`` dict.
        remove_fn = getattr(self._rust_writer, "remove_sheet", None)
        if remove_fn is not None:
            remove_fn(title)

    def remove_sheet(self, worksheet: Worksheet) -> None:
        """openpyxl alias for :meth:`remove` (deprecated there, kept for parity)."""
        self.remove(worksheet)

    # ------------------------------------------------------------------
    # Workbook-level metadata (T1 PR3)
    # ------------------------------------------------------------------

    @property
    def properties(self) -> Any:
        """Return the workbook's :class:`DocumentProperties` (lazy-loaded).

        In read/modify mode, parses ``docProps/core.xml`` once via the
        Rust reader and caches the result. In write mode, starts as an
        empty (all-fields-None) ``DocumentProperties`` whose attribute
        assignments flip ``self._properties_dirty`` so :meth:`save` knows
        to flush them.
        """
        if self._properties_cache is not None:
            return self._properties_cache
        from wolfxl.packaging.core import DocumentProperties, _doc_props_from_dict

        if self._rust_reader is not None:
            try:
                raw = self._rust_reader.read_doc_properties()
            except Exception:
                raw = {}
            props = _doc_props_from_dict(raw)
        else:
            props = DocumentProperties()
        # Attach the back-reference so subsequent ``props.title = "X"``
        # marks the workbook dirty without further user action.
        props._attach_workbook(self)  # noqa: SLF001
        self._properties_cache = props
        return props

    @properties.setter
    def properties(self, value: Any) -> None:
        """Replace the entire properties object wholesale.

        Used by callers that prefer to construct a fresh
        ``DocumentProperties`` rather than mutate fields one at a time.
        Sets the dirty flag unconditionally — replacing the object is by
        definition a write intent.
        """
        from wolfxl.packaging.core import DocumentProperties

        if not isinstance(value, DocumentProperties):
            raise TypeError(
                f"properties must be a DocumentProperties, got {type(value).__name__}"
            )
        value._attach_workbook(self)  # noqa: SLF001
        self._properties_cache = value
        self._properties_dirty = True

    # ------------------------------------------------------------------
    # Named ranges
    # ------------------------------------------------------------------

    @property
    def defined_names(self) -> Any:
        """Return the workbook's :class:`DefinedNameDict`.

        Lazy-loaded on first access. The container is a ``dict``
        subclass whose values are :class:`DefinedName` objects.
        Workbook-scoped names override sheet-scoped on collision.
        Mutations (``wb.defined_names["X"] = DefinedName(...)``) queue
        through to the Rust writer in write mode.
        """
        if self._defined_names_cache is not None:
            return self._defined_names_cache
        from wolfxl.workbook import DefinedNameDict
        from wolfxl.workbook.defined_name import DefinedName

        dnd = DefinedNameDict()
        if self._rust_reader is not None:
            seen: set[str] = set()
            for sheet_name in self._sheet_names:
                try:
                    entries = self._rust_reader.read_named_ranges(sheet_name)
                except Exception:
                    continue
                for entry in entries:
                    name = entry["name"]
                    if name in seen:
                        continue
                    seen.add(name)
                    refers_to = entry["refers_to"]
                    if refers_to.startswith("="):
                        refers_to = refers_to[1:]
                    scope = entry.get("scope", "workbook")
                    local_id: int | None = None
                    if scope == "sheet":
                        # The sheet-scope encoding in the Rust reader puts
                        # the sheet name in the ``refers_to`` prefix; we
                        # don't try to recover the original index.
                        local_id = None
                    dn = DefinedName(name=name, value=refers_to, localSheetId=local_id)
                    # Bypass __setitem__'s queue side-effect — this is a
                    # pure read, not a user write.
                    dict.__setitem__(dnd, name, dn)
        # Attach the workbook back-ref so subsequent user writes queue.
        dnd._wb = self  # noqa: SLF001
        self._defined_names_cache = dnd
        return dnd

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
            # Modify mode — workbook-level metadata writes don't have a
            # patcher path yet (T1.5 follow-up). Surface the limitation
            # before mutating the file rather than silently dropping the
            # user's edits.
            if self._properties_dirty:
                raise NotImplementedError(
                    "Mutating wb.properties on an existing file is a T1.5 follow-up. "
                    "Workaround: open via Workbook() and re-author, or stop modifying "
                    "wb.properties before save()."
                )
            if self._pending_defined_names:
                raise NotImplementedError(
                    "Mutating wb.defined_names on an existing file is a T1.5 follow-up. "
                    "Workaround: open via Workbook() and re-author, or stop modifying "
                    "wb.defined_names before save()."
                )
            for ws in self._sheets.values():
                ws._flush()  # noqa: SLF001
            self._rust_patcher.save(filename)
        elif self._rust_writer is not None:
            # Write mode — flush workbook-level writes, then sheets.
            self._flush_workbook_writes()
            for ws in self._sheets.values():
                ws._flush()  # noqa: SLF001
            self._rust_writer.save(filename)
        else:
            raise RuntimeError("save requires write or modify mode")

    def _flush_workbook_writes(self) -> None:
        """Push workbook-level metadata + defined names into the Rust writer."""
        writer = self._rust_writer
        if writer is None:
            return

        if self._properties_dirty and self._properties_cache is not None:
            props = self._properties_cache
            payload = {
                "title": props.title,
                "subject": props.subject,
                "creator": props.creator,
                "keywords": props.keywords,
                "description": props.description,
                "lastModifiedBy": props.lastModifiedBy,
                "category": props.category,
                "contentStatus": props.contentStatus,
                "identifier": props.identifier,
                "language": props.language,
                "revision": props.revision,
                "version": props.version,
                "created": props.created.isoformat() if props.created else None,
                "modified": props.modified.isoformat() if props.modified else None,
            }
            writer.set_properties(payload)
            self._properties_dirty = False

        if self._pending_defined_names:
            # The native writer's add_named_range expects a sheet hint —
            # for workbook-scoped names we pick the first sheet; the Rust
            # side reads ``localSheetId`` from the dict to override.
            primary_sheet = self._sheet_names[0] if self._sheet_names else "Sheet"
            for _, dn in self._pending_defined_names.items():
                writer.add_named_range(primary_sheet, {
                    "name": dn.name,
                    "refers_to": dn.value,
                    "comment": dn.comment,
                    "local_sheet_id": dn.localSheetId,
                    "hidden": dn.hidden,
                })
            self._pending_defined_names.clear()

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
