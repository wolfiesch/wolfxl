"""Workbook — multi-mode openpyxl-compatible wrapper.

Write mode (``Workbook()``): creates a new workbook via NativeWorkbook.
Read mode (``Workbook._from_reader(path)``): opens an existing .xlsx via CalamineStyledBook.
Modify mode (``Workbook._from_patcher(path)``): read via CalamineStyledBook, save via XlsxPatcher.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from wolfxl._worksheet import Worksheet


@dataclass
class CopyOptions:
    """Per-workbook flags controlling :meth:`Workbook.copy_worksheet`.

    Attributes:
        deep_copy_images: When ``True``, drawings reachable from a
            cloned sheet have their referenced ``xl/media/imageN.<ext>``
            targets DEEP-CLONED into freshly numbered media parts.
            When ``False`` (default), the cloned drawing rels point at
            the same image bytes as the source — Excel's historical
            RFC-035 §5.3 alias behaviour. Modify-mode only;
            write-mode ignores this flag (write-mode clones run via
            in-memory replay, not the modify-mode planner).
    """

    deep_copy_images: bool = False

if TYPE_CHECKING:
    from wolfxl.calc._protocol import RecalcResult


def _xlsb_xls_via_tempfile(
    rust_cls: Any,
    data: bytes | bytearray | memoryview,
    *,
    suffix: str,
    permissive: bool,
) -> tuple[Any, str]:
    """Materialise ``data`` to a tempfile and call ``rust_cls.open(path)``.

    Used as a fallback for ``CalamineXlsbBook`` / ``CalamineXlsBook``
    when Pod-α's ``open_from_bytes`` overload isn't yet exposed.
    Returns ``(rust_book, tempfile_path)`` so the caller can stash the
    path on the workbook for cleanup at ``close()`` time.
    """
    import tempfile

    with tempfile.NamedTemporaryFile(
        prefix="wolfxl-", suffix=suffix, delete=False
    ) as tmp:
        tmp.write(bytes(data))
        tmp_path = tmp.name

    opener = rust_cls.open
    try:
        rust_book = opener(tmp_path, permissive)
    except TypeError:
        rust_book = opener(tmp_path)
    return rust_book, tmp_path


def _build_xlsb_xls_wb(
    cls: type,
    *,
    rust_book: Any,
    fmt: str,
    data_only: bool,
    source_path: str | None,
) -> Any:
    """Wire up the read-mode workbook fields shared by xlsb / xls.

    Skips the workbook-property and defined-name caches because the
    binary backends don't expose them; everything else mirrors
    :meth:`Workbook._from_reader` so existing call sites
    (``wb.sheetnames``, ``wb.active``, ``ws['A1'].value``) keep working.
    """
    wb = object.__new__(cls)
    wb._rust_writer = None
    wb._rust_patcher = None
    wb._rust_reader = rust_book
    wb._data_only = data_only
    wb._rich_text = False
    wb._evaluator = None
    wb._read_only = False
    wb._source_path = source_path
    wb._format = fmt
    names = [str(n) for n in rust_book.sheet_names()]
    wb._sheet_names = names
    wb._sheets = {name: Worksheet(wb, name) for name in names}
    # Keep the rest of the boilerplate empty — these caches and queues
    # are only meaningful for xlsx (modify mode + write mode).
    wb._properties_cache = None
    wb._properties_dirty = False
    wb._defined_names_cache = None
    wb._pending_defined_names = {}
    wb._pending_axis_shifts = []
    wb._pending_range_moves = []
    wb._pending_sheet_copies = []
    wb.copy_options = CopyOptions()
    return wb


class Workbook:
    """openpyxl-compatible workbook backed by Rust."""

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

        wb = object.__new__(cls)
        wb._rust_writer = None
        wb._rust_patcher = None
        wb._data_only = data_only
        wb._rich_text = False
        wb._evaluator = None
        wb._rust_reader = _rust.CalamineStyledBook.open(path, permissive)
        wb._read_only = read_only
        wb._source_path = path
        wb._format = "xlsx"
        names = [str(n) for n in wb._rust_reader.sheet_names()]
        wb._sheet_names = names
        wb._sheets = {}
        for name in names:
            wb._sheets[name] = Worksheet(wb, name)
        wb._properties_cache = None
        wb._properties_dirty = False
        wb._defined_names_cache = None
        wb._pending_defined_names = {}
        wb._pending_axis_shifts = []
        wb._pending_range_moves = []
        wb._pending_sheet_copies = []
        wb.copy_options = CopyOptions()
        return wb

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
                        path, data_only=data_only, permissive=permissive
                    )
                # Bytes input that isn't encrypted: route through the
                # bytes shim so we don't lose the data.
                return cls._from_bytes(
                    bytes(data),  # type: ignore[arg-type]
                    data_only=data_only,
                    permissive=permissive,
                    modify=modify,
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
        wb = object.__new__(cls)
        wb._rust_writer = None
        wb._rust_patcher = None
        wb._data_only = data_only
        wb._rich_text = False
        wb._evaluator = None
        wb._rust_reader = bytes_open(data_bytes, permissive)
        wb._read_only = read_only
        wb._source_path = None
        wb._format = "xlsx"
        names = [str(n) for n in wb._rust_reader.sheet_names()]
        wb._sheet_names = names
        wb._sheets = {}
        for name in names:
            wb._sheets[name] = Worksheet(wb, name)
        wb._properties_cache = None
        wb._properties_dirty = False
        wb._defined_names_cache = None
        wb._pending_defined_names = {}
        wb._pending_axis_shifts = []
        wb._pending_range_moves = []
        wb._pending_sheet_copies = []
        wb.copy_options = CopyOptions()
        return wb

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

        wb = object.__new__(cls)
        wb._rust_writer = None
        wb._data_only = data_only
        wb._rich_text = False
        wb._evaluator = None
        wb._rust_reader = _rust.CalamineStyledBook.open(path, permissive)
        wb._rust_patcher = _rust.XlsxPatcher.open(path, permissive)
        wb._read_only = False
        wb._source_path = path
        wb._format = "xlsx"
        names = [str(n) for n in wb._rust_reader.sheet_names()]
        wb._sheet_names = names
        wb._sheets = {}
        for name in names:
            wb._sheets[name] = Worksheet(wb, name)
        wb._properties_cache = None
        wb._properties_dirty = False
        wb._defined_names_cache = None
        wb._pending_defined_names = {}
        wb._pending_axis_shifts = []
        wb._pending_range_moves = []
        wb._pending_sheet_copies = []
        wb.copy_options = CopyOptions()
        return wb

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
                rust_book, tmp_path = _xlsb_xls_via_tempfile(
                    rust_cls, data, suffix=".xlsb", permissive=permissive
                )
                _wb = _build_xlsb_xls_wb(
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

        return _build_xlsb_xls_wb(
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
                rust_book, tmp_path = _xlsb_xls_via_tempfile(
                    rust_cls, data, suffix=".xls", permissive=permissive
                )
                _wb = _build_xlsb_xls_wb(
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

        return _build_xlsb_xls_wb(
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
        if not isinstance(source, Worksheet):
            raise TypeError(
                f"copy_worksheet: source must be a Worksheet, got {type(source).__name__}"
            )
        if source._workbook is not self:  # noqa: SLF001
            raise ValueError(
                "copy_worksheet: source must belong to this workbook"
            )
        if self._rust_patcher is None and self._rust_writer is None:
            raise RuntimeError(
                "copy_worksheet requires write or modify mode"
            )

        # Compute the new title. Explicit `name` wins; otherwise dedup
        # against the running tab list.
        if name is not None:
            if not isinstance(name, str) or not name:
                raise ValueError("copy_worksheet: name must be a non-empty string")
            if name in self._sheets:
                raise ValueError(f"copy_worksheet: sheet '{name}' already exists")
            new_title = name
        else:
            base = f"{source.title} Copy"
            new_title = base
            suffix = 2
            while new_title in self._sheets:
                new_title = f"{base} {suffix}"
                suffix += 1

        if self._rust_patcher is not None:
            # Modify-mode path — queue + tab-list update; ZIP-level clone
            # happens during save() via Phase 2.7. Snapshot the
            # deep_copy_images flag at queue time so a later toggle
            # of wb.copy_options doesn't retroactively affect this
            # already-queued copy.
            self._pending_sheet_copies.append(
                (source.title, new_title, bool(self.copy_options.deep_copy_images))
            )
            self._sheet_names.append(new_title)
            ws = Worksheet(self, new_title)
            self._sheets[new_title] = ws
            return ws

        # Write-mode path (Sprint Θ Pod-C1).
        return self._copy_worksheet_write_mode(source, new_title)

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
        from wolfxl._cell import _UNSET

        # 1. Materialise the source's pending buffers so every value is
        #    in `_cells` and discoverable. These helpers are idempotent.
        if source._append_buffer:  # noqa: SLF001
            source._materialize_append_buffer()  # noqa: SLF001
        if source._bulk_writes:  # noqa: SLF001
            source._materialize_bulk_writes()  # noqa: SLF001

        # 2. Add a fresh destination sheet. Use the public-ish helper so
        #    the Rust writer gets the new sheet registered first.
        dst = self.create_sheet(new_title)

        # 3. Walk the source's cell map. We iterate `_cells` (not
        #    `_dirty`) because cells materialised from append/bulk
        #    buffers go through `cell(...)` which writes via the
        #    public setter that flips `_value_dirty` — so they're in
        #    `_dirty` too — but a future caller might construct cells
        #    via direct attribute writes. Using `_cells` is the
        #    superset and makes the snapshot deterministic.
        for (row, col), src_cell in source._cells.items():  # noqa: SLF001
            value = src_cell._value  # noqa: SLF001
            has_value = value is not _UNSET and src_cell._value_dirty  # noqa: SLF001
            font = src_cell._font  # noqa: SLF001
            fill = src_cell._fill  # noqa: SLF001
            border = src_cell._border  # noqa: SLF001
            alignment = src_cell._alignment  # noqa: SLF001
            number_format = src_cell._number_format  # noqa: SLF001
            has_format = src_cell._format_dirty  # noqa: SLF001

            if not has_value and not has_format:
                # Cell exists only because it was probed for read; do
                # not propagate (would inflate destination dimensions).
                continue

            # Use cell() which builds a Cell, so the value/format
            # setters mark dirty correctly for downstream `_flush`.
            dst_cell = dst.cell(row=row, column=col)
            if has_value:
                dst_cell.value = value
            if font is not _UNSET:
                dst_cell.font = font  # type: ignore[assignment]
            if fill is not _UNSET:
                dst_cell.fill = fill  # type: ignore[assignment]
            if border is not _UNSET:
                dst_cell.border = border  # type: ignore[assignment]
            if alignment is not _UNSET:
                dst_cell.alignment = alignment  # type: ignore[assignment]
            if number_format is not _UNSET:
                dst_cell.number_format = number_format  # type: ignore[assignment]

        # 4. Sheet-scope properties.
        for r, h in source._row_heights.items():  # noqa: SLF001
            dst._row_heights[r] = h  # noqa: SLF001
        for letter, w in source._col_widths.items():  # noqa: SLF001
            dst._col_widths[letter] = w  # noqa: SLF001
        # Merges: round-trip through merge_cells so the Rust writer
        # also gets the merge — `_merged_ranges` is just the Python
        # mirror set; the writer needs an explicit call to record it.
        for rng in source._merged_ranges:  # noqa: SLF001
            dst.merge_cells(rng)
        if source._freeze_panes is not None:  # noqa: SLF001
            dst._freeze_panes = source._freeze_panes  # noqa: SLF001
        if source._print_area is not None:  # noqa: SLF001
            dst._print_area = source._print_area  # noqa: SLF001

        return dst

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
        # Type-check sheet.
        if isinstance(sheet, Worksheet):
            name = sheet.title
        elif isinstance(sheet, str):
            name = sheet
        else:
            raise TypeError(
                f"move_sheet: 'sheet' must be a Worksheet or str, got {type(sheet).__name__}"
            )

        # Reject bool explicitly (isinstance(True, int) is True in Python,
        # which would silently treat True as 1 / False as 0).
        if isinstance(offset, bool) or not isinstance(offset, int):
            raise TypeError(
                f"move_sheet: 'offset' must be an int, got {type(offset).__name__}"
            )

        # Validate sheet name.
        if name not in self._sheet_names:
            raise KeyError(name)

        n = len(self._sheet_names)
        idx = self._sheet_names.index(name)
        new_pos = idx + offset
        # Clamp to [0, n-1], matching the patcher-side rule. Python's
        # list.insert clamps too, but we do it explicitly so the queued
        # offset matches the position the patcher will compute.
        if new_pos < 0:
            new_pos = 0
        if new_pos > n - 1:
            new_pos = n - 1

        # Update the in-memory tab list. Even when no actual position
        # change happens (offset=0 or clamped no-op), we still walk the
        # patcher-queue path so a downstream caller observing the queue
        # matches the user's intent.
        del self._sheet_names[idx]
        self._sheet_names.insert(new_pos, name)

        # Queue the move on the patcher in modify mode. The Rust side
        # re-resolves the offset against its own running tab list, so
        # we pass the user's original offset (not the clamped one) for
        # symmetry with the openpyxl signature.
        if self._rust_patcher is not None:
            self._flush_pending_sheet_moves_to_patcher(name, offset)

    def save(
        self,
        filename: str | os.PathLike[str],
        *,
        password: str | bytes | None = None,
    ) -> None:
        """Flush all pending writes and save to disk.

        When ``password`` is supplied, the freshly written plaintext
        xlsx is re-encoded as an OOXML-encrypted blob (Agile / AES-256,
        the modern Excel default) via :mod:`wolfxl._encryption` before
        being placed at ``filename``. Both the write-mode (``Workbook()``)
        and modify-mode (``open(..., modify=True)``) save paths are
        wrapped — encryption is applied to the final byte stream
        regardless of which Rust backend produced it.

        ``password`` accepts ``str`` or ``bytes`` (UTF-8 decoded). An
        empty string / empty bytes raises :class:`ValueError`. The
        ``msoffcrypto-tool`` dep is loaded lazily; install via
        ``pip install wolfxl[encrypted]``. See ``docs/encryption.md``
        for the supported-algorithm matrix.
        """
        filename = str(filename)
        if password is not None:
            # Validate password early so we don't write a plaintext
            # tempfile that we'd then have to throw away.
            from wolfxl._encryption import _coerce_password

            _coerce_password(password)  # raises ValueError on empty
            self._save_encrypted(filename, password)
            return
        if self._rust_patcher is not None:
            # Modify mode — workbook-level metadata writes don't have a
            # patcher path yet (T1.5 follow-up). Surface the limitation
            # before mutating the file rather than silently dropping the
            # user's edits.
            # RFC-020: properties round-trip (Phase 2.5d in the patcher).
            # Workbook-level, so it flushes before the per-sheet drains.
            if self._properties_dirty:
                self._flush_properties_to_patcher()
            if self._pending_defined_names:
                self._flush_defined_names_to_patcher()
            for ws in self._sheets.values():
                ws._flush()  # noqa: SLF001
            # RFC-035: sheet copies must flush BEFORE every per-sheet
            # phase so cloned sheets are visible to downstream drains
            # (cell patches, hyperlinks, tables, comments, axis shifts,
            # range moves) as if they had always been part of the
            # source workbook.
            self._flush_pending_sheet_copies_to_patcher()
            # RFC-022: hyperlinks share the sheet rels graph with future
            # rels-touching writers (RFC-024 tables, RFC-023 comments).
            # Flush them first so DV/CF (which don't touch rels) run
            # afterward against an already-stable rels graph.
            self._flush_pending_hyperlinks_to_patcher()
            # RFC-024: tables also touch the rels graph + add new ZIP
            # parts + content-type Overrides. Flush after hyperlinks
            # so the rels graph already carries any external-hyperlink
            # rIds when build_tables iterates rels.iter() to assemble
            # the merged <tableParts> block.
            self._flush_pending_tables_to_patcher()
            # RFC-023: comments + VML drawings.
            self._flush_pending_comments_to_patcher()
            # RFC-025: flush worksheet-level setters that the patcher
            # accepts. The cell-level _flush above handles values +
            # formats; data validations are a separate patcher API
            # because they live in a sibling block, not in <sheetData>.
            self._flush_pending_data_validations_to_patcher()
            # RFC-026: conditional formatting also lives in a sibling
            # block (slot 17). Cross-sheet dxfId allocation happens
            # inside the patcher's Phase-2.5b on the Rust side.
            self._flush_pending_conditional_formats_to_patcher()
            # RFC-030 / RFC-031: structural axis shifts (insert/delete
            # rows/cols). Drained LAST so it sees the per-cell + per-block
            # rewrites from the earlier flush calls and shifts them too.
            self._flush_pending_axis_shifts_to_patcher()
            # RFC-034: range moves. Drained AFTER axis shifts so a
            # sequence like `insert_rows(2, 3)` then
            # `move_range("C3:E10", rows=5)` is applied in source order
            # against the post-shift coordinate space.
            self._flush_pending_range_moves_to_patcher()
            # Sprint Λ Pod-β (RFC-045): drain pending images.
            self._flush_pending_images_to_patcher()
            # Sprint Μ Pod-β (RFC-046): drain pending charts.
            self._flush_pending_charts_to_patcher()
            self._rust_patcher.save(filename)
        elif self._rust_writer is not None:
            # Write mode — flush workbook-level writes, then sheets.
            self._flush_workbook_writes()
            for ws in self._sheets.values():
                ws._flush()  # noqa: SLF001
            self._rust_writer.save(filename)
        else:
            raise RuntimeError("save requires write or modify mode")

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
        import tempfile

        from wolfxl._encryption import encrypt_xlsx_to_path

        tmp_fd, tmp_name = tempfile.mkstemp(
            prefix=".wolfxl-plain-",
            suffix=".xlsx",
        )
        os.close(tmp_fd)
        try:
            # Re-enter save() in plaintext mode by calling the original
            # path. Doing the call without ``password=`` keeps the
            # branch logic simple and ensures both writer and patcher
            # paths are exercised identically.
            self.save(tmp_name)
            with open(tmp_name, "rb") as fp:
                plaintext_bytes = fp.read()
            encrypt_xlsx_to_path(plaintext_bytes, password, filename)
        finally:
            try:
                os.unlink(tmp_name)
            except OSError:
                pass

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
        patcher = self._rust_patcher
        if patcher is None:
            return
        for ws in self._sheets.values():
            pending = ws._pending_hyperlinks  # noqa: SLF001
            if not pending:
                continue
            for coord, hl in pending.items():
                if hl is None:
                    patcher.queue_hyperlink_delete(ws.title, coord)
                    continue
                payload: dict[str, Any] = {}
                if hl.target is not None:
                    payload["target"] = hl.target
                if hl.location is not None:
                    payload["location"] = hl.location
                if hl.tooltip is not None:
                    payload["tooltip"] = hl.tooltip
                if hl.display is not None:
                    payload["display"] = hl.display
                patcher.queue_hyperlink(ws.title, coord, payload)
            pending.clear()

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
        patcher = self._rust_patcher
        if patcher is None:
            return
        for ws in self._sheets.values():
            pending = ws._pending_tables  # noqa: SLF001
            if not pending:
                continue
            for t in pending:
                payload: dict[str, Any] = {
                    "name": t.name,
                    "ref": t.ref,
                    "columns": [c.name for c in t.tableColumns] if t.tableColumns else [],
                    "header_row_count": int(t.headerRowCount or 0),
                    "totals_row_shown": bool(t.totalsRowCount and t.totalsRowCount > 0),
                    "autofilter": True,
                }
                if t.displayName and t.displayName != t.name:
                    payload["display_name"] = t.displayName
                if t.tableStyleInfo is not None and t.tableStyleInfo.name:
                    payload["style"] = {
                        "name": t.tableStyleInfo.name,
                        "show_first_column": bool(t.tableStyleInfo.showFirstColumn),
                        "show_last_column": bool(t.tableStyleInfo.showLastColumn),
                        "show_row_stripes": bool(t.tableStyleInfo.showRowStripes),
                        "show_column_stripes": bool(t.tableStyleInfo.showColumnStripes),
                    }
                patcher.queue_table(ws.title, payload)
            pending.clear()

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
        from wolfxl._images import image_to_writer_payload

        patcher = self._rust_patcher
        if patcher is None:
            return
        for ws in self._sheets.values():
            pending = ws._pending_images  # noqa: SLF001
            if not pending:
                continue
            for img in pending:
                payload = image_to_writer_payload(img)
                patcher.queue_image_add(ws.title, payload)
            pending.clear()

    def _flush_pending_charts_to_patcher(self) -> None:
        """Sprint Μ Pod-β (RFC-046) — drain pending charts into the patcher.

        Modify-mode counterpart to the writer-side flush in
        ``Worksheet._flush_compat_properties``. Each queued ``ChartBase``
        is serialised via :meth:`ChartBase.to_rust_dict` and routed to
        ``XlsxPatcher.queue_chart_add`` (Pod-γ owns the patcher binding).

        If the patcher doesn't expose ``queue_chart_add`` (because Pod-γ
        hasn't merged yet), we warn rather than raise so existing chart-free
        modify-mode flows don't regress.
        """
        patcher = self._rust_patcher
        if patcher is None:
            return
        if not hasattr(patcher, "queue_chart_add"):
            import warnings

            for ws in self._sheets.values():
                if ws._pending_charts:  # noqa: SLF001
                    warnings.warn(
                        "wolfxl.chart: modify-mode chart flush requires "
                        "Pod-γ's queue_chart_add patcher binding (not yet "
                        f"available). Dropping {len(ws._pending_charts)} "  # noqa: SLF001
                        f"chart(s) on sheet {ws.title!r}.",
                        RuntimeWarning,
                        stacklevel=2,
                    )
                    ws._pending_charts.clear()  # noqa: SLF001
            return
        for ws in self._sheets.values():
            pending = ws._pending_charts  # noqa: SLF001
            if not pending:
                continue
            for chart in pending:
                payload = chart.to_rust_dict()
                patcher.queue_chart_add(ws.title, payload, chart._anchor)  # noqa: SLF001
            pending.clear()

    def _flush_pending_comments_to_patcher(self) -> None:
        """Drain ``_pending_comments`` on every sheet into the patcher (RFC-023).

        Modify-mode counterpart to the writer-side flush at
        ``_worksheet.py:1934``. Each ``Comment`` is converted to the
        patcher's flat-dict shape and routed to ``queue_comment``;
        the ``None`` sentinel routes to ``queue_comment_delete``.
        """
        patcher = self._rust_patcher
        if patcher is None:
            return
        for ws in self._sheets.values():
            pending = ws._pending_comments  # noqa: SLF001
            if not pending:
                continue
            for coord, c in pending.items():
                if c is None:
                    patcher.queue_comment_delete(ws.title, coord)
                    continue
                payload: dict[str, Any] = {
                    "text": c.text,
                    "author": c.author or "wolfxl",
                }
                if getattr(c, "width", None) is not None:
                    payload["width_pt"] = float(c.width)
                if getattr(c, "height", None) is not None:
                    payload["height_pt"] = float(c.height)
                patcher.queue_comment(ws.title, coord, payload)
            pending.clear()

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
        from wolfxl.worksheet.datavalidation import _dv_to_patcher_dict

        patcher = self._rust_patcher
        if patcher is None:
            return
        for ws in self._sheets.values():
            pending = ws._pending_data_validations  # noqa: SLF001
            if not pending:
                continue
            for dv in pending:
                patcher.queue_data_validation(ws.title, _dv_to_patcher_dict(dv))
            pending.clear()

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
        from wolfxl.formatting import _cf_to_patcher_dict

        patcher = self._rust_patcher
        if patcher is None:
            return
        for ws in self._sheets.values():
            pending = ws._pending_conditional_formats  # noqa: SLF001
            if not pending:
                continue
            by_sqref: dict[str, list[Any]] = {}
            order: list[str] = []
            for sqref, rule in pending:
                if sqref not in by_sqref:
                    by_sqref[sqref] = []
                    order.append(sqref)
                by_sqref[sqref].append(rule)
            for sqref in order:
                patcher.queue_conditional_formatting(
                    ws.title, _cf_to_patcher_dict(sqref, by_sqref[sqref])
                )
            pending.clear()

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
        patcher = self._rust_patcher
        if patcher is None or not self._pending_axis_shifts:
            return
        for sheet_title, axis, idx, n in self._pending_axis_shifts:
            patcher.queue_axis_shift(sheet_title, axis, idx, n)
        self._pending_axis_shifts.clear()

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
        patcher = self._rust_patcher
        if patcher is None or not self._pending_range_moves:
            return
        for (
            sheet_title,
            src_min_col,
            src_min_row,
            src_max_col,
            src_max_row,
            d_row,
            d_col,
            translate,
        ) in self._pending_range_moves:
            patcher.queue_range_move(
                sheet_title,
                src_min_col,
                src_min_row,
                src_max_col,
                src_max_row,
                d_row,
                d_col,
                translate,
            )
        self._pending_range_moves.clear()

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
        patcher = self._rust_patcher
        if patcher is None or not self._pending_sheet_copies:
            return
        for src_title, dst_title, deep_copy_images in self._pending_sheet_copies:
            patcher.queue_sheet_copy(src_title, dst_title, deep_copy_images)
        self._pending_sheet_copies.clear()

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
        patcher = self._rust_patcher
        if patcher is None or not self._pending_defined_names:
            return
        for _, dn in self._pending_defined_names.items():
            payload: dict[str, Any] = {
                "name": dn.name,
                "formula": dn.value,
            }
            if dn.localSheetId is not None:
                payload["local_sheet_id"] = dn.localSheetId
            if dn.hidden:
                # Only forward when truthy — the Rust side treats
                # missing-key and `None` as "preserve / omit".
                payload["hidden"] = True
            if dn.comment is not None:
                payload["comment"] = dn.comment
            patcher.queue_defined_name(payload)
        self._pending_defined_names.clear()

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
        patcher = self._rust_patcher
        if patcher is None:
            return
        patcher.queue_sheet_move(name, offset)

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
        patcher = self._rust_patcher
        if patcher is None:
            return
        props = self._properties_cache
        if props is None:
            self._properties_dirty = False
            return
        # Per-field "user explicitly set this" set, populated by
        # ``DocumentProperties.__setattr__`` after ``_attach_workbook``.
        # Used below to decide whether to forward ``modified``: by
        # default a dirty save re-stamps it to save-time (Rust side),
        # which is what users expect. The cache hydrates ``modified``
        # from the source on first ``wb.properties`` read — we'd
        # otherwise echo the source's old timestamp forever.
        user_set: set[str] = getattr(props, "_user_set", set())
        modified_iso: str | None = None
        if "modified" in user_set and props.modified is not None:
            modified_iso = props.modified.isoformat()
        payload: dict[str, Any] = {
            "title": props.title,
            "subject": props.subject,
            "creator": props.creator,
            "keywords": props.keywords,
            "description": props.description,
            "last_modified_by": props.lastModifiedBy,
            "category": props.category,
            "content_status": props.contentStatus,
            "created_iso": props.created.isoformat() if props.created else None,
            "modified_iso": modified_iso,
            "sheet_names": list(self._sheet_names),
        }
        payload = {k: v for k, v in payload.items() if v is not None}
        patcher.queue_properties(payload)
        self._properties_dirty = False

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
            # The native writer's add_named_range expects a sheet hint
            # plus an explicit ``scope`` token — workbook-scoped names
            # use the first sheet (the value is ignored when scope ==
            # "workbook"), sheet-scoped names resolve to the sheet at
            # ``localSheetId``.
            primary_sheet = self._sheet_names[0] if self._sheet_names else "Sheet"
            for _, dn in self._pending_defined_names.items():
                if dn.localSheetId is not None:
                    if 0 <= dn.localSheetId < len(self._sheet_names):
                        sheet_hint = self._sheet_names[dn.localSheetId]
                    else:
                        sheet_hint = primary_sheet
                    scope = "sheet"
                else:
                    sheet_hint = primary_sheet
                    scope = "workbook"
                writer.add_named_range(sheet_hint, {
                    "name": dn.name,
                    "refers_to": dn.value,
                    "scope": scope,
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
        # Sprint Ι Pod-γ: clean up the decryption tempfile, if any.
        tmp_path = getattr(self, "_tempfile_path", None)
        if tmp_path is not None:
            import os

            try:
                os.unlink(tmp_path)
            except OSError:
                pass
            self._tempfile_path = None

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
