"""Workbook — multi-mode openpyxl-compatible wrapper.

Write mode (``Workbook()``): creates a new workbook via NativeWorkbook.
Read mode (``Workbook._from_reader(path)``): opens an existing .xlsx via CalamineStyledBook.
Modify mode (``Workbook._from_patcher(path)``): read via CalamineStyledBook, save via XlsxPatcher.
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, Any

from wolfxl._workbook_state import (
    CopyOptions,
    build_xlsx_wb,
    build_xlsb_xls_wb,
    xlsb_xls_via_tempfile,
)
from wolfxl import _workbook_features
from wolfxl import _workbook_calc
from wolfxl import _workbook_metadata
from wolfxl import _workbook_lifecycle
from wolfxl import _workbook_patcher_flush
from wolfxl import _workbook_save
from wolfxl import _workbook_sheets
from wolfxl import _workbook_writer_flush
from wolfxl._worksheet import Worksheet


if TYPE_CHECKING:
    from wolfxl.calc._protocol import RecalcResult


class Workbook:
    """Openpyxl-compatible workbook backed by Rust.

    A workbook operates in one of three modes:

    * write mode, created with ``Workbook()``;
    * read mode, created with ``load_workbook(path)``;
    * modify mode, created with ``load_workbook(path, modify=True)``.

    Public methods mirror openpyxl where practical while routing heavy I/O
    through WolfXL's native reader, writer, and patcher backends.
    """

    def __init__(self) -> None:
        """Create a new workbook in write mode with a default 'Sheet'."""
        from wolfxl import _backend, _rust  # noqa: F401  (_rust kept for typing parity)

        self._rust_writer: Any = _backend.make_writer()
        self._rust_reader: Any = None
        self._rust_patcher: Any = None
        self._data_only = False
        # Sprint Ι Pod-α — flipped to True via load_workbook(rich_text=True).
        self._rich_text: bool = False
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
        # RFC-058 — workbook-level security (workbookProtection + fileSharing).
        # ``_security`` and ``_file_sharing`` hold user-supplied
        # WorkbookProtection / FileSharing instances. ``_pending_security_update``
        # is a sentinel: True once a setter touched either slot, drained at
        # save() time so the writer / patcher emit the corresponding XML
        # blocks. None ⇒ no security configured (default).
        self._security: Any | None = None
        self._file_sharing: Any | None = None
        self._pending_security_update: bool = False
        # RFC-030 / RFC-031 — append-order list of structural shift ops.
        # Tuple shape: ``(sheet_title, axis: "row"|"col", idx, n_signed)``.
        self._pending_axis_shifts: list[tuple[str, str, int, int]] = []
        # RFC-034 — append-order list of range-move ops.
        # Tuple shape: ``(sheet_title, src_min_col, src_min_row,
        # src_max_col, src_max_row, d_row, d_col, translate)``.
        self._pending_range_moves: list[
            tuple[str, int, int, int, int, int, int, bool]
        ] = []
        # RFC-035 — append-order list of sheet-copy ops.
        # Tuple shape: ``(src_title, dst_title, deep_copy_images)``.
        # The deep_copy_images flag is snapshot at copy_worksheet()
        # call time so a later toggle of wb.copy_options doesn't
        # retroactively affect already-queued copies.
        self._pending_sheet_copies: list[tuple[str, str, bool]] = []
        # Sprint Μ Pod-γ (RFC-046 §6) — pending modify-mode chart adds.
        # Per-sheet list of ``(chart_xml: bytes, anchor_a1: str,
        # width_emu: int, height_emu: int)`` tuples. Pod-β's
        # ``Worksheet.add_chart`` populates this in modify mode (the
        # writer-mode path stays on Pod-α's NativeWorkbook bindings).
        # Drained by ``_flush_pending_charts_to_patcher`` in save().
        self._pending_chart_adds: dict[
            str, list[tuple[bytes, str, int, int]]
        ] = {}
        # Sprint Ν Pod-γ (RFC-047 / RFC-048) — pending pivot caches +
        # pivot table adds. Caches are workbook-scope (one cache → N
        # tables); tables live on the owner Worksheet's
        # ``_pending_pivot_tables``. Drained by
        # ``_flush_pending_pivots_to_patcher`` at save() time AFTER
        # charts (Phase 2.5l) so the matching patcher Phase 2.5m runs
        # against an already-stable rels graph.
        self._pending_pivot_caches: list[Any] = []
        # 0-based cache id allocator. Bumps when add_pivot_cache() is
        # called so the first cache is `cache_id=0` (matches OOXML
        # convention of 0-based cacheId in <pivotCache>).
        self._next_pivot_cache_id: int = 0
        # RFC-061 Sub-feature 3.1 — slicer caches (workbook-scoped).
        self._pending_slicer_caches: list[Any] = []
        self._next_slicer_cache_id: int = 0
        # Sprint Θ Pod-C2 — workbook-level copy options.
        self.copy_options: CopyOptions = CopyOptions()
        # Sprint Ι Pod-β — streaming read flag (write mode never streams).
        self._read_only: bool = False
        self._source_path: str | None = None
        # Sprint Κ Pod-β — file format the workbook came from.  Write
        # mode is xlsx by definition; the read/modify constructors set
        # this to "xlsx" / "xlsb" / "xls" as appropriate.
        self._format: str = "xlsx"

    @classmethod
    def _from_reader(
        cls,
        path: str,
        *,
        data_only: bool = False,
        permissive: bool = False,
        read_only: bool = False,
    ) -> Workbook:
        """Open an existing .xlsx file in read mode.

        ``permissive`` plumbs through to the Rust reader and triggers a
        rels-graph fallback when ``<sheets>`` is empty/self-closing.

        ``read_only`` activates the SAX streaming fast path on
        ``iter_rows`` (Sprint Ι Pod-β). The CalamineStyledBook reader
        is still constructed for style/format lookups (used by the
        non-streaming Cell properties), but the streaming reader
        bypasses calamine's eager materialization for the large-sheet
        scan path. See :func:`wolfxl.load_workbook` for details.
        """
        from wolfxl import _rust

        return build_xlsx_wb(
            cls,
            rust_reader=_rust.CalamineStyledBook.open(path, permissive),
            rust_patcher=None,
            data_only=data_only,
            read_only=read_only,
            source_path=path,
        )

    @classmethod
    def _from_encrypted(
        cls,
        path: str | None = None,
        *,
        data: bytes | bytearray | memoryview | None = None,
        password: str | bytes,
        data_only: bool = False,
        permissive: bool = False,
        modify: bool = False,
        read_only: bool = False,
    ) -> Workbook:
        """Open an OOXML-encrypted .xlsx via msoffcrypto-tool (Sprint Ι Pod-γ).

        Decrypts the source (path or in-memory blob) into an in-memory
        buffer, then dispatches through the bytes-aware reader path
        (Sprint Κ Pod-β unified entry point). On a non-encrypted file
        the password is silently ignored and the normal path is used
        (matches openpyxl).

        Wrong / missing passwords raise ``ValueError`` with a clear
        message; ``ImportError`` (with install hint) surfaces when
        ``msoffcrypto-tool`` isn't installed.

        Modify mode + password works because the decrypted bytes are
        rematerialised through ``_from_bytes``; on save the result is
        plaintext (write-side encryption is documented T3
        out-of-scope).

        Exactly one of ``path`` / ``data`` must be supplied — the
        ``load_workbook`` dispatcher (Sprint Κ Pod-β) threads whichever
        the caller passed in.
        """
        if (path is None) == (data is None):
            raise TypeError(
                "_from_encrypted requires exactly one of path / data"
            )

        if path is not None:
            with open(path, "rb") as fp:
                is_plain_xlsx = fp.read(4).startswith(b"PK")
        else:
            is_plain_xlsx = bytes(data).startswith(b"PK")  # type: ignore[arg-type]

        if is_plain_xlsx:
            # openpyxl-style: silently ignore password on a non-encrypted
            # xlsx without requiring the optional encryption dependency.
            if path is not None:
                if modify:
                    return cls._from_patcher(
                        path, data_only=data_only, permissive=permissive
                    )
                return cls._from_reader(
                    path,
                    data_only=data_only,
                    permissive=permissive,
                    read_only=read_only,
                )
            return cls._from_bytes(
                bytes(data),  # type: ignore[arg-type]
                data_only=data_only,
                permissive=permissive,
                modify=modify,
                read_only=read_only,
            )

        # Lazy import — users without the optional dep pay no cost.
        try:
            import msoffcrypto  # type: ignore[import-not-found]
        except ImportError as exc:
            raise ImportError(
                "password reads require msoffcrypto-tool; install with: "
                "pip install wolfxl[encrypted]"
            ) from exc

        import io

        pw_str: str
        if isinstance(password, bytes):
            pw_str = password.decode("utf-8")
        else:
            pw_str = password

        # Funnel both inputs into a BytesIO so msoffcrypto sees the same
        # shape regardless of caller path.
        if path is not None:
            src_fp = open(path, "rb")  # noqa: SIM115 — closed in finally
        else:
            src_fp = io.BytesIO(bytes(data))  # type: ignore[arg-type]

        try:
            office = msoffcrypto.OfficeFile(src_fp)
            try:
                is_encrypted = office.is_encrypted()
            except Exception:
                # Some msoffcrypto versions raise on non-OOXML inputs.
                is_encrypted = False

            if not is_encrypted:
                # openpyxl-style: silently ignore password on a
                # non-encrypted file.
                if path is not None:
                    if modify:
                        return cls._from_patcher(
                            path, data_only=data_only, permissive=permissive
                        )
                    return cls._from_reader(
                        path,
                        data_only=data_only,
                        permissive=permissive,
                        read_only=read_only,
                    )
                # Bytes input that isn't encrypted: route through the
                # bytes shim so we don't lose the data.
                return cls._from_bytes(
                    bytes(data),  # type: ignore[arg-type]
                    data_only=data_only,
                    permissive=permissive,
                    modify=modify,
                    read_only=read_only,
                )

            try:
                office.load_key(password=pw_str)
            except Exception as exc:
                raise ValueError(
                    f"failed to load decryption key: {exc}"
                ) from exc

            buf = io.BytesIO()
            try:
                office.decrypt(buf)
            except Exception as exc:
                # msoffcrypto raises InvalidKeyError on wrong password;
                # normalise to ValueError so callers don't have to
                # import msoffcrypto just to except.
                raise ValueError(
                    f"failed to decrypt workbook (wrong password?): {exc}"
                ) from exc
            decrypted_bytes = buf.getvalue()
        finally:
            src_fp.close()

        return cls._from_bytes(
            decrypted_bytes,
            data_only=data_only,
            permissive=permissive,
            modify=modify,
            read_only=read_only,
        )

    @classmethod
    def _from_bytes(
        cls,
        data: bytes | bytearray | memoryview,
        *,
        data_only: bool = False,
        permissive: bool = False,
        modify: bool = False,
        read_only: bool = False,
    ) -> Workbook:
        """Open an .xlsx blob from memory (Sprint Ι Pod-γ, Sprint Κ Pod-β).

        When the underlying Rust reader exposes ``open_from_bytes``
        (Pod-α), the blob is handed to the reader directly with no
        intermediate tempfile.  Otherwise the bytes are materialised to
        a tempfile and the path-based reader / patcher is used; the
        tempfile is tracked on the workbook so :meth:`close` can clean
        it up.  Either way the public surface (``Workbook._format``,
        ``Workbook._source_path``, etc.) is identical.

        ``read_only`` plumbs through to the streaming SAX path (Sprint
        Ι Pod-β); ``modify=True`` always uses a tempfile because the
        XlsxPatcher is path-only by design (it reopens the source zip
        on save).
        """
        from wolfxl import _rust

        data_bytes = bytes(data)

        # Modify mode requires the patcher, which is path-only.  Same
        # for the no-bytes-direct fallback when the Rust reader hasn't
        # been taught about bytes inputs yet (Pod-α dependency).
        bytes_open = getattr(_rust.CalamineStyledBook, "open_from_bytes", None)
        needs_tempfile = modify or bytes_open is None

        if needs_tempfile:
            import tempfile

            # delete=False so the file persists across the
            # NamedTemporaryFile context manager. We track the path on
            # the workbook and remove it from ``close()``.
            with tempfile.NamedTemporaryFile(
                prefix="wolfxl-", suffix=".xlsx", delete=False
            ) as tmp:
                tmp.write(data_bytes)
                tmp_path = tmp.name

            if modify:
                wb = cls._from_patcher(
                    tmp_path, data_only=data_only, permissive=permissive
                )
            else:
                wb = cls._from_reader(
                    tmp_path,
                    data_only=data_only,
                    permissive=permissive,
                    read_only=read_only,
                )
            wb._tempfile_path = tmp_path
            return wb

        # Bytes-direct reader path (Pod-α onwards): no tempfile needed.
        return build_xlsx_wb(
            cls,
            rust_reader=bytes_open(data_bytes, permissive),
            rust_patcher=None,
            data_only=data_only,
            read_only=read_only,
            source_path=None,
        )

    @classmethod
    def _from_patcher(
        cls,
        path: str,
        *,
        data_only: bool = False,
        permissive: bool = False,
    ) -> Workbook:
        """Open an existing .xlsx file in modify mode (read + surgical save).

        ``permissive`` plumbs through to both the reader and the
        patcher; see :func:`wolfxl.load_workbook` for the user-facing
        contract.
        """
        from wolfxl import _rust

        return build_xlsx_wb(
            cls,
            rust_reader=_rust.CalamineStyledBook.open(path, permissive),
            rust_patcher=_rust.XlsxPatcher.open(path, permissive),
            data_only=data_only,
            read_only=False,
            source_path=path,
        )

    @classmethod
    def _from_xlsb(
        cls,
        *,
        path: str | None,
        data: bytes | None,
        data_only: bool = False,
        permissive: bool = False,
    ) -> Workbook:
        """Open an .xlsb workbook via Pod-α's ``CalamineXlsbBook``.

        xlsb is a binary OOXML container; we surface values + cached
        formula results only (no per-cell styles, no rich text, no
        comments — that's the same shape calamine's stock xlsb reader
        exposes).  Callers that need style metadata should
        load + transcribe to xlsx first.
        """
        from wolfxl import _rust

        rust_cls = getattr(_rust, "CalamineXlsbBook", None)
        if rust_cls is None:
            raise NotImplementedError(
                ".xlsb reads require the CalamineXlsbBook backend "
                "(Sprint Κ Pod-α). Rebuild the wolfxl extension after "
                "Pod-α merges, or use openpyxl/xlrd as an interim."
            )

        if data is not None:
            bytes_open = getattr(rust_cls, "open_from_bytes", None)
            if bytes_open is None:
                # Fall back to a tempfile so we can still hand the
                # backend a path while Pod-α plumbs the bytes overload.
                rust_book, tmp_path = xlsb_xls_via_tempfile(
                    rust_cls, data, suffix=".xlsb", permissive=permissive
                )
                _wb = build_xlsb_xls_wb(
                    cls,
                    rust_book=rust_book,
                    fmt="xlsb",
                    data_only=data_only,
                    source_path=None,
                )
                _wb._tempfile_path = tmp_path
                return _wb
            try:
                rust_book = bytes_open(data, permissive)
            except TypeError:
                # xlsb/xls backends don't expose a permissive flag.
                rust_book = bytes_open(data)
        else:
            opener = getattr(rust_cls, "open", None)
            if opener is None:
                raise NotImplementedError(
                    "CalamineXlsbBook.open is not yet exposed by the "
                    "Rust extension; rebuild after Sprint Κ Pod-α."
                )
            try:
                rust_book = opener(path, permissive)
            except TypeError:
                # Pod-α may not yet thread `permissive` through.
                rust_book = opener(path)

        return build_xlsb_xls_wb(
            cls,
            rust_book=rust_book,
            fmt="xlsb",
            data_only=data_only,
            source_path=path,
        )

    @classmethod
    def _from_xls(
        cls,
        *,
        path: str | None,
        data: bytes | None,
        data_only: bool = False,
        permissive: bool = False,
    ) -> Workbook:
        """Open a legacy .xls workbook via Pod-α's ``CalamineXlsBook``.

        Same shape as :meth:`_from_xlsb` — values + cached formula
        results only.
        """
        from wolfxl import _rust

        rust_cls = getattr(_rust, "CalamineXlsBook", None)
        if rust_cls is None:
            raise NotImplementedError(
                ".xls reads require the CalamineXlsBook backend "
                "(Sprint Κ Pod-α). Rebuild the wolfxl extension after "
                "Pod-α merges, or use xlrd as an interim."
            )

        if data is not None:
            bytes_open = getattr(rust_cls, "open_from_bytes", None)
            if bytes_open is None:
                rust_book, tmp_path = xlsb_xls_via_tempfile(
                    rust_cls, data, suffix=".xls", permissive=permissive
                )
                _wb = build_xlsb_xls_wb(
                    cls,
                    rust_book=rust_book,
                    fmt="xls",
                    data_only=data_only,
                    source_path=None,
                )
                _wb._tempfile_path = tmp_path
                return _wb
            try:
                rust_book = bytes_open(data, permissive)
            except TypeError:
                # xlsb/xls backends don't expose a permissive flag.
                rust_book = bytes_open(data)
        else:
            opener = getattr(rust_cls, "open", None)
            if opener is None:
                raise NotImplementedError(
                    "CalamineXlsBook.open is not yet exposed by the "
                    "Rust extension; rebuild after Sprint Κ Pod-α."
                )
            try:
                rust_book = opener(path, permissive)
            except TypeError:
                rust_book = opener(path)

        return build_xlsb_xls_wb(
            cls,
            rust_book=rust_book,
            fmt="xls",
            data_only=data_only,
            source_path=path,
        )

    # ------------------------------------------------------------------
    # Sheet access
    # ------------------------------------------------------------------

    @property
    def sheetnames(self) -> list[str]:
        """Return worksheet titles in tab order."""
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
        """True if this workbook was opened with ``read_only=True``.

        Sprint Ι Pod-β changes the semantics of this flag: it now
        reflects the *explicit* ``read_only=True`` opt-in passed to
        :func:`wolfxl.load_workbook`, not the historic
        "no-writer-no-patcher" inference. The streaming
        ``iter_rows`` fast path keys off the explicit flag (or the
        > 50k row auto-trigger). Workbooks opened in plain read mode
        (``read_only=False``, the default) retain the historic
        in-memory behaviour for ``__getitem__`` / mutation paths.
        """
        return bool(getattr(self, "_read_only", False))

    @property
    def chartsheets(self) -> list[Any]:
        """Chart sheets - always empty in T0 (wolfxl treats charts as preserved-only)."""
        return []

    @property
    def named_styles(self) -> list[Any]:
        """Named styles - always empty in T0 (construction lands in T2)."""
        return []

    def __getitem__(self, name: str) -> Worksheet:
        """Return a worksheet by title."""
        if name not in self._sheets:
            raise KeyError(f"Worksheet '{name}' does not exist")
        return self._sheets[name]

    def __contains__(self, name: str) -> bool:
        """Return whether the workbook contains a sheet named ``name``."""
        return name in self._sheets

    def __iter__(self):  # type: ignore[no-untyped-def]
        """Iterate worksheet titles in tab order."""
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
        _workbook_sheets.remove_sheet(self, worksheet)

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
        return _workbook_metadata.get_properties(self)

    @properties.setter
    def properties(self, value: Any) -> None:
        """Replace the entire properties object wholesale.

        Used by callers that prefer to construct a fresh
        ``DocumentProperties`` rather than mutate fields one at a time.
        Sets the dirty flag unconditionally — replacing the object is by
        definition a write intent.
        """
        _workbook_metadata.set_properties(self, value)

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
        return _workbook_metadata.get_defined_names(self)

    # ------------------------------------------------------------------
    # RFC-058 — workbook-level security
    # ------------------------------------------------------------------

    @property
    def security(self) -> Any:
        """Return the workbook's :class:`WorkbookProtection` block, if any.

        ``None`` when no protection is configured (the default). Assign a
        :class:`wolfxl.workbook.protection.WorkbookProtection` instance to
        enable structure / window / revision locks. Mutating an already-
        attached instance also queues the update — call
        ``wb.security = wb.security`` after mutating to force the flush
        if needed (the property write is what flips the dirty flag).
        """
        return self._security

    @security.setter
    def security(self, value: Any) -> None:
        """Set workbook protection metadata.

        Args:
            value: A ``WorkbookProtection`` instance, or ``None`` to clear the
                in-memory protection block before the next save.
        """
        _workbook_metadata.set_security(self, value)

    @property
    def fileSharing(self) -> Any:  # noqa: N802 — openpyxl-shape camelCase
        """Return the workbook's :class:`FileSharing` block, if any.

        ``None`` when no file-sharing block is configured (the default).
        Attribute name matches openpyxl's ``wb.fileSharing`` exactly so
        existing code continues to work.
        """
        return self._file_sharing

    @fileSharing.setter
    def fileSharing(self, value: Any) -> None:  # noqa: N802
        """Set workbook file-sharing metadata.

        Args:
            value: A ``FileSharing`` instance, or ``None`` to clear the
                in-memory file-sharing block before the next save.
        """
        _workbook_metadata.set_file_sharing(self, value)

    # ------------------------------------------------------------------
    # Write-mode operations
    # ------------------------------------------------------------------

    def create_sheet(self, title: str) -> Worksheet:
        """Create and append a worksheet.

        Args:
            title: Unique worksheet title.

        Returns:
            The newly created :class:`Worksheet`.

        Raises:
            RuntimeError: If the workbook is not in write mode.
            ValueError: If ``title`` already exists.
        """
        return _workbook_sheets.create_sheet(self, title)

    def copy_worksheet(
        self, source: Worksheet, *, name: str | None = None
    ) -> Worksheet:
        """Duplicate *source* into a new sheet within this workbook (RFC-035).

        Supported in BOTH modify mode and write mode (Sprint Θ Pod-C1).
        Read-only mode raises ``RuntimeError``.

        The new sheet appends at the end of the tab list. The default
        title is ``f"{source.title} Copy"``; on collision an incrementing
        suffix (`Copy 2`, `Copy 3`, …) is appended until unique. An
        explicit ``name`` keyword argument overrides the default and
        must not collide with any existing sheet name.

        Modify mode: the returned ``Worksheet`` is a fresh proxy bound
        to the cloned title. The actual ZIP-level clone runs at
        ``save()`` time via Phase 2.7 of the patcher.

        Write mode: the source's pending writes are materialized
        immediately and replayed onto a freshly-created destination
        sheet (cell values, formats, row heights, column widths, merged
        ranges, freeze pane). Native-writer-tracked features added by
        the API after `copy_worksheet` returns flow through normally.
        """
        return _workbook_sheets.copy_worksheet(self, source, name=name)

    def _copy_worksheet_write_mode(
        self, source: Worksheet, new_title: str
    ) -> Worksheet:
        """Clone an in-memory worksheet into a fresh sheet (write mode).

        Materialises any pending append/bulk-write buffers on the source
        so every cell lives in ``source._cells`` first, then walks that
        map and replays each cell's value + format/border onto the
        destination's lazily-allocated ``Cell`` objects. Sheet-scope
        attributes (row heights, column widths, merged ranges, freeze
        pane) are copied verbatim. The destination sheet is registered
        with the native writer via ``create_sheet`` so that downstream
        save/flush passes see it like any other sheet.
        """
        return _workbook_sheets.copy_worksheet_write_mode(self, source, new_title)

    def move_sheet(self, sheet: Worksheet | str, offset: int = 0) -> None:
        """Move *sheet* by *offset* positions within the workbook tab list.

        Mirrors openpyxl's ``Workbook.move_sheet`` (RFC-036). The new
        position is ``current_index + offset``, clamped to ``[0, n-1]``
        where ``n`` is the current sheet count. The in-memory tab list
        (``self._sheet_names``) is updated immediately so subsequent
        reads of ``wb.sheetnames`` / ``wb.worksheets`` see the post-move
        order, regardless of whether the workbook is in write or modify
        mode.

        In modify mode, the move is queued on the patcher (along with
        any previous moves in this save() session); on save the patcher
        rewrites ``xl/workbook.xml``'s ``<sheets>`` order and re-points
        every sheet-scoped ``<definedName localSheetId>`` accordingly
        (RFC-036 §5).

        Args:
            sheet: A ``Worksheet`` instance or sheet name string.
            offset: Integer count of positions to shift.

        Raises:
            TypeError: ``sheet`` is neither a ``Worksheet`` nor a
                ``str``; or ``offset`` is not an integer (``bool``
                is rejected explicitly).
            KeyError: the resolved sheet name is not in this workbook.
        """
        _workbook_sheets.move_sheet(self, sheet, offset)

    def save(
        self,
        filename: str | os.PathLike[str],
        *,
        password: str | bytes | None = None,
    ) -> None:
        """Flush all pending writes and save to disk.

        Args:
            filename: Destination path. In modify mode this may be the original
                source path; WolfXL uses the patcher's atomic in-place save path
                for that case.
            password: Optional encryption password for the final ``.xlsx``
                payload. Install ``wolfxl[encrypted]`` to enable encryption.

        Raises:
            ValueError: If ``password`` is empty.
            RuntimeError: If the workbook mode cannot save the requested
                pending changes.
        """
        _workbook_save.save_workbook(self, filename, password=password)

    def _save_encrypted(
        self,
        filename: str,
        password: str | bytes,
    ) -> None:
        """Save plaintext to a tempfile then re-route through encryption.

        Sprint Λ Pod-α: write-side encryption stays Python-side (same
        as Sprint Ι Pod-γ's read path); the Rust writer/patcher is
        unchanged. We materialise the unencrypted xlsx via the normal
        save path, slurp it back as bytes, hand it to
        :func:`wolfxl._encryption.encrypt_xlsx_to_path` for the
        in-place encryption + atomic rename onto ``filename``.

        The plaintext tempfile is always cleaned up — including on
        error paths — so we never leak unencrypted user data on disk.
        """
        _workbook_save.save_encrypted(self, filename, password)

    def _flush_pending_hyperlinks_to_patcher(self) -> None:
        """Drain ``_pending_hyperlinks`` on every sheet into the patcher (RFC-022).

        Modify-mode counterpart to the writer-side flush at
        ``_worksheet.py:1911``. Each ``Hyperlink`` is converted to the
        patcher's flat-dict shape and routed to ``queue_hyperlink``;
        the ``None`` sentinel routes to ``queue_hyperlink_delete``
        (INDEX decision #5 — never use ``pop()``).

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit.
        """
        _workbook_patcher_flush.flush_pending_hyperlinks_to_patcher(self)

    def _flush_pending_tables_to_patcher(self) -> None:
        """Drain ``_pending_tables`` on every sheet into the patcher (RFC-024).

        Modify-mode counterpart to the writer flush at
        ``_worksheet.py:1946``. Each ``Table`` is converted to the
        patcher's flat-dict shape and routed to ``queue_table``. The
        patcher allocates a workbook-unique table ``id`` at save time
        (any explicit ``id`` on the Python ``Table`` object is
        ignored), serializes ``xl/tables/tableN.xml``, splices a
        ``<tableParts>`` block into the sheet XML, mutates the sheet
        rels, and adds a ``[Content_Types].xml`` Override.

        Per-sheet drain happens in workbook tab order; within a sheet,
        append order wins (which matches openpyxl's first-add → first-
        slot semantics).

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit.
        """
        _workbook_patcher_flush.flush_pending_tables_to_patcher(self)

    def _flush_pending_images_to_patcher(self) -> None:
        """Sprint Λ Pod-β (RFC-045) — drain pending images into the patcher.

        Modify-mode counterpart to the writer-side flush in
        ``Worksheet._flush_compat_properties``. Each queued
        :class:`wolfxl.drawing.image.Image` is converted to the flat
        dict shape and routed to ``XlsxPatcher.queue_image_add``.

        Sheets that already have a drawing rel will surface
        ``NotImplementedError`` from the patcher at save time —
        appending to an existing drawing is a v1.5 follow-up.
        """
        _workbook_patcher_flush.flush_pending_images_to_patcher(self)

    def _flush_pending_charts_to_patcher(self) -> None:
        """Sprint Μ-prime Pod-γ′ (RFC-046 §6, §10.12) — drain pending
        chart adds in modify mode.

        Two queues are drained here:

        1. Workbook-level ``_pending_chart_adds`` (Pod-γ): a dict keyed
           by sheet title with values of ``(chart_xml: bytes, anchor_a1,
           width_emu, height_emu)`` tuples populated via
           :meth:`add_chart_modify_mode`. These are routed straight to
           ``XlsxPatcher.queue_chart_add``. This is the bytes-level
           escape hatch (v1.6.0) — preserved for callers that want to
           pass pre-serialised chart XML directly.
        2. Per-sheet ``Worksheet._pending_charts`` (Pod-β): a list of
           high-level :class:`~wolfxl.chart._chart.ChartBase` instances
           queued via :meth:`Worksheet.add_chart`. In **write** mode
           these are drained inside
           ``_worksheet._flush_compat_properties`` via
           ``writer.add_chart_native``. In **modify** mode (v1.6.1+,
           Sprint Μ-prime) we now bridge each chart through Pod-α′'s
           ``serialize_chart_dict`` PyO3 export, producing chart XML
           bytes that are routed through the same
           ``patcher.queue_chart_add`` path as the bytes escape hatch.

        Sequenced AFTER images / axis shifts but BEFORE the final
        ``patcher.save()`` so chart cell-range formulas can compose
        with cell rewrites in the same save (the patcher's Phase 2.5l
        runs before Phase 3 cell patches).
        """
        _workbook_patcher_flush.flush_pending_charts_to_patcher(self)

    def add_pivot_cache(self, cache: Any) -> Any:
        """Register a pivot cache with this workbook.

        The cache should already be materialized, for example with
        ``wolfxl.pivot.PivotCache.from_worksheet``. A registered cache can be
        referenced by one or more pivot tables.

        Args:
            cache: Pivot cache object to register.

        Returns:
            The same cache object, with its workbook-scoped cache id set.

        Raises:
            RuntimeError: If the workbook is not open in modify mode.
            ValueError: If the cache has already been registered.
        """
        return _workbook_features.add_pivot_cache(self, cache)

    def _flush_pending_slicers_to_patcher(self) -> None:
        """Sprint Ο Pod 3.5 (RFC-061 §3.1) — drain queued slicers
        into the patcher's Phase 2.5p queue.

        For each worksheet, iterate over ``ws._pending_slicers`` and
        bridge the (cache, slicer) pair to the Rust patcher via
        ``queue_slicer_add(sheet_title, cache_dict, slicer_dict)``.
        Sequenced AFTER pivots (Phase 2.5m) + sheet-setup (Phase 2.5n)
        and BEFORE autofilters (Phase 2.5o).
        """
        _workbook_patcher_flush.flush_pending_slicers_to_patcher(self)

    def add_slicer_cache(self, cache: Any) -> Any:
        """Register a slicer cache with this workbook.

        Slicer caches are workbook-scoped, and one cache can be referenced by
        multiple slicer presentations. The source pivot cache must already be
        registered with :meth:`add_pivot_cache`.

        Args:
            cache: Slicer cache object to register.

        Returns:
            The same cache object, with its workbook-scoped slicer cache id
            set.

        Raises:
            RuntimeError: If the workbook is not open in modify mode.
            ValueError: If the cache has already been registered or
                the source pivot cache is not registered.
        """
        return _workbook_features.add_slicer_cache(self, cache)

    def _flush_pending_sheet_setup_to_patcher(self) -> None:
        """Sprint Ο Pod 1A.5 (RFC-055) — drain each sheet's queued
        sheet-setup mutations into the patcher's Phase 2.5n queue.

        Sheets whose Worksheet has any of ``_page_setup``,
        ``_page_margins``, ``_header_footer``, ``_sheet_view``,
        ``_protection``, ``_print_title_rows``, ``_print_title_cols``
        non-default get their ``to_rust_setup_dict()`` queued. The
        Rust patcher Phase 2.5n then re-emits the 5 sheet-scope
        XML blocks and splices them into the sheet via
        wolfxl_merger::merge_blocks.

        ``print_titles`` (workbook-scope ``_xlnm.Print_Titles``
        definedName) does NOT route through Phase 2.5n on the
        patcher side; it composes through the existing RFC-021
        defined-names queue. The dict still includes a
        ``print_titles`` slot for the writer-mode path.
        """
        _workbook_patcher_flush.flush_pending_sheet_setup_to_patcher(self)

    def _flush_pending_page_breaks_to_patcher(self) -> None:
        """Sprint Π Pod Π-α (RFC-062) — drain each sheet's queued
        page-breaks + sheet-format-pr mutations into the patcher's
        Phase 2.5r queue.

        Sheets whose Worksheet has any of ``_row_breaks``,
        ``_col_breaks``, or ``_sheet_format`` non-default get their
        merged §10 dict queued. The Rust patcher Phase 2.5r then
        re-emits the 3 sheet-scope XML blocks
        (``<rowBreaks>`` / ``<colBreaks>`` / ``<sheetFormatPr>``) and
        splices them into the sheet via wolfxl_merger::merge_blocks.

        Sheets whose all three slots are at construction defaults
        (e.g. zero breaks AND default sheet-format) are skipped to
        keep the no-op save path byte-identical.
        """
        _workbook_patcher_flush.flush_pending_page_breaks_to_patcher(self)

    def _flush_pending_autofilters_to_patcher(self) -> None:
        """Sprint Ο Pod 1B (RFC-056) — drain each sheet's
        ``ws.auto_filter`` into the patcher's Phase 2.5o queue.

        Only sheets where the user actually configured filter columns
        OR a sort state OR (legacy) just a ref are queued. The Rust
        patcher Phase 2.5o then re-emits the ``<autoFilter>`` block
        and computes the ``<row hidden="1">`` markers.
        """
        _workbook_patcher_flush.flush_pending_autofilters_to_patcher(self)

    def _flush_pending_pivots_to_patcher(self) -> None:
        """Sprint Ν Pod-γ (RFC-047 / RFC-048) — drain pending pivot
        caches and tables in modify mode.

        Two queues are drained here:

        1. Workbook-level ``_pending_pivot_caches`` — a list of
           :class:`~wolfxl.pivot.PivotCache` instances queued via
           :meth:`add_pivot_cache`. Each cache is bridged through
           ``serialize_pivot_cache_dict`` (definition XML) +
           ``serialize_pivot_records_dict`` (records XML), then
           routed to ``patcher.queue_pivot_cache_add``.
        2. Per-sheet ``Worksheet._pending_pivot_tables`` — a list of
           :class:`~wolfxl.pivot.PivotTable` instances queued via
           :meth:`Worksheet.add_pivot_table`. Each table is bridged
           through ``serialize_pivot_table_dict`` and routed to
           ``patcher.queue_pivot_table_add(sheet, xml, cache_id)``.

        Sequenced AFTER charts (Phase 2.5l) and BEFORE the final
        ``patcher.save()`` so the patcher's Phase 2.5m runs against
        an already-stable rels graph.
        """
        _workbook_patcher_flush.flush_pending_pivots_to_patcher(self)

    def add_chart_modify_mode(
        self,
        sheet_title: str,
        chart_xml: bytes,
        anchor_a1: str,
        width_emu: int = 4_572_000,
        height_emu: int = 2_743_200,
    ) -> None:
        """Queue a pre-serialized chart XML part for a worksheet.

        This lower-level API is for callers that already have chart XML bytes
        and want to attach them while modifying an existing workbook. Most
        callers should prefer :meth:`Worksheet.add_chart` with a chart object.

        Args:
            sheet_title: Title of the target worksheet.
            chart_xml: Serialized ``<chartSpace>`` XML bytes.
            anchor_a1: A1-style anchor cell, such as ``"D2"``.
            width_emu: Chart width in EMUs.
            height_emu: Chart height in EMUs.

        Raises:
            ValueError: If ``sheet_title`` or ``anchor_a1`` is invalid.
            RuntimeError: If the workbook is not open in modify mode.
        """
        _workbook_features.add_chart_modify_mode(
            self,
            sheet_title,
            chart_xml,
            anchor_a1,
            width_emu,
            height_emu,
        )

    def _flush_pending_comments_to_patcher(self) -> None:
        """Drain ``_pending_comments`` on every sheet into the patcher (RFC-023).

        Modify-mode counterpart to the writer-side flush at
        ``_worksheet.py:1934``. Each ``Comment`` is converted to the
        patcher's flat-dict shape and routed to ``queue_comment``;
        the ``None`` sentinel routes to ``queue_comment_delete``.
        """
        _workbook_patcher_flush.flush_pending_comments_to_patcher(self)

    def _flush_pending_data_validations_to_patcher(self) -> None:
        """Drain ``_pending_data_validations`` on every sheet into the patcher.

        Modify-mode counterpart to the writer flush at
        ``_worksheet.py:1960`` — same drain semantics, different
        backend. Each DV is converted to the patcher's flat-dict
        payload via ``_dv_to_patcher_dict``. Per-sheet drain happens
        in ``ws.title`` order; within a sheet, append order wins so
        the final ``<dataValidations>`` block reflects the order the
        user appended them.

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit.
        """
        _workbook_patcher_flush.flush_pending_data_validations_to_patcher(self)

    def _flush_pending_conditional_formats_to_patcher(self) -> None:
        """Drain ``_pending_conditional_formats`` on every sheet into the patcher.

        Modify-mode counterpart to the writer flush at
        ``_worksheet.py:1974`` — same drain semantics, different backend.
        Rules sharing a sqref are coalesced into a single
        ``ConditionalFormattingPatch`` (one wrapper per range) so
        priority ordering within a wrapper reflects insertion order.
        Multiple ``add()`` calls with different sqrefs produce multiple
        patches; the patcher then emits them in encounter order while
        threading the workbook-wide ``running_dxf_count`` through
        Phase-2.5b on the Rust side.

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit.
        """
        _workbook_patcher_flush.flush_pending_conditional_formats_to_patcher(self)

    def _flush_pending_axis_shifts_to_patcher(self) -> None:
        """Drain ``_pending_axis_shifts`` into the patcher (RFC-030 / RFC-031).

        Each tuple ``(sheet_title, axis, idx, n)`` is forwarded to
        ``_rust_patcher.queue_axis_shift(sheet, axis, idx, n)``. The
        patcher's Phase 2.5i drains the queue in append order during
        ``save()``.

        Empty queue is the no-op identity path — patcher is not
        called, no FFI hop, no file mutation.

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit.
        """
        _workbook_patcher_flush.flush_pending_axis_shifts_to_patcher(self)

    def _flush_pending_range_moves_to_patcher(self) -> None:
        """Drain ``_pending_range_moves`` into the patcher (RFC-034).

        Each tuple ``(sheet_title, src_min_col, src_min_row,
        src_max_col, src_max_row, d_row, d_col, translate)`` is
        forwarded to ``_rust_patcher.queue_range_move(...)``. The
        patcher's Phase 2.5j drains the queue in append order during
        ``save()``.

        Empty queue is the no-op identity path — patcher is not
        called, no FFI hop, no file mutation.

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit.
        """
        _workbook_patcher_flush.flush_pending_range_moves_to_patcher(self)

    def _flush_pending_sheet_copies_to_patcher(self) -> None:
        """Drain ``_pending_sheet_copies`` into the patcher (RFC-035).

        Each ``(src_title, dst_title)`` pair forwards to
        ``_rust_patcher.queue_sheet_copy(src, dst)``. The patcher's
        Phase 2.7 drains the queue in append order during ``save()``,
        BEFORE every per-sheet phase so the cloned sheets are visible
        to downstream drains.

        Empty queue is the no-op identity path — patcher is not
        called, no FFI hop, no file mutation. Cleared after queueing
        so a subsequent ``save()`` on the same workbook doesn't
        double-emit.
        """
        _workbook_patcher_flush.flush_pending_sheet_copies_to_patcher(self)

    def _flush_defined_names_to_patcher(self) -> None:
        """Drain ``_pending_defined_names`` into the patcher (RFC-021).

        Modify-mode counterpart to ``_flush_workbook_writes``'s
        defined-name branch. Each ``DefinedName`` is converted to the
        patcher's flat-dict shape and routed to
        ``_rust_patcher.queue_defined_name``. ``None``-valued optional
        fields are filtered out before crossing the PyO3 boundary so
        the Rust extractors see a clean "missing key" rather than a
        Python ``None`` (matches the convention in
        ``_flush_properties_to_patcher``).

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit. Empty queue is a no-op (the Rust
        side's no-op guard is the second line of defence — workbook.xml
        is left untouched if no upserts arrive).
        """
        _workbook_patcher_flush.flush_defined_names_to_patcher(self)

    def _flush_security_to_patcher(self) -> None:
        """Drain ``_security`` / ``_file_sharing`` into the patcher (RFC-058).

        Builds the §10 flat dict and forwards it to
        ``_rust_patcher.queue_workbook_security``. The Rust side merges
        the payload into ``xl/workbook.xml`` during Phase 2.5q.

        Empty (no setter ever ran) ⇒ no-op; the patcher leaves
        workbook.xml byte-identical with the source.
        """
        _workbook_patcher_flush.flush_security_to_patcher(self)

    def _build_security_dict(self) -> dict[str, Any]:
        """Return the RFC-058 §10 flat dict for the workbook's security blocks.

        Either branch may be ``None`` (the user only set one of the two
        slots). Always returns a dict — never ``None`` — so callers can
        unconditionally forward to the Rust side.
        """
        return _workbook_patcher_flush.build_security_dict(self)

    def _flush_pending_sheet_moves_to_patcher(self, name: str, offset: int) -> None:
        """Queue a single sheet-reorder on the patcher (RFC-036).

        Called eagerly from ``move_sheet`` rather than batched at
        ``save()`` time: each ``move_sheet`` call queues exactly one
        entry, and the patcher composes them in queue order against
        its own running tab list (which is initialised from the
        source ZIP's ``xl/workbook.xml`` and updated in place by
        Phase 2.5h on save).

        The empty-queue invariant lives on the Rust side: an unused
        ``move_sheet`` call (i.e. modify-mode workbook never touched)
        means ``queued_sheet_moves`` is empty, which in turn keeps
        ``xl/workbook.xml`` byte-identical with the source.
        """
        _workbook_patcher_flush.queue_sheet_move_to_patcher(self, name, offset)

    def _flush_properties_to_patcher(self) -> None:
        """Drain dirty document properties into the patcher (RFC-020).

        Modify-mode counterpart to ``_flush_workbook_writes``'s
        property branch. Builds a flat dict keyed with the patcher's
        snake_case schema (``last_modified_by``, ``content_status``,
        ``created_iso``, ``modified_iso``) and filters ``None`` before
        crossing the PyO3 boundary so ``extract_str`` sees a clean
        "missing key" rather than a Python ``None``.

        Resets ``_properties_dirty`` so a subsequent ``save()`` on the
        same workbook doesn't double-emit. ``modified_iso`` is left
        unset on this side — the Rust patcher stamps it via
        ``current_timestamp_iso8601`` (or ``WOLFXL_TEST_EPOCH=0`` →
        ``1970-01-01T00:00:00Z`` for byte-identical save tests). If the
        user explicitly set ``props.modified``, that value wins.
        """
        _workbook_patcher_flush.flush_properties_to_patcher(self)

    def _flush_workbook_writes(self) -> None:
        """Push workbook-level metadata + defined names into the Rust writer."""
        _workbook_writer_flush.flush_workbook_writes(self)

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
        return _workbook_calc.calculate_workbook(self)

    def cached_formula_values(self) -> dict[str, Any]:
        """Return Excel-saved cached formula results for every sheet.

        Keys are workbook-qualified cell references like ``"Sheet1!B2"``.
        This is a fast read-only path for ingestion workloads that need
        Excel's last-calculated formula values without evaluating formulas in
        Python. Cells whose formulas have no cached value are omitted.
        """
        return _workbook_calc.cached_formula_values(self)

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
        return _workbook_calc.recalculate_workbook(self, perturbations, tolerance)

    # ------------------------------------------------------------------
    # Context manager + cleanup
    # ------------------------------------------------------------------

    def close(self) -> None:
        """Release native handles and delete any temporary decrypted input."""
        _workbook_lifecycle.close_workbook(self)

    def __enter__(self) -> Workbook:
        """Return this workbook for ``with`` statement use."""
        return _workbook_lifecycle.enter_workbook(self)

    def __exit__(self, *args: object) -> None:
        """Close this workbook at the end of a ``with`` block."""
        _workbook_lifecycle.exit_workbook(self, *args)

    def __repr__(self) -> str:
        """Return a compact debug representation for this workbook.

        Returns:
            A string containing the workbook mode and sheet names.
        """
        return _workbook_lifecycle.repr_workbook(self)
