"""Shared workbook state and constructor helpers."""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any

from wolfxl._worksheet import Worksheet


@dataclass
class CopyOptions:
    """Per-workbook flags controlling :meth:`Workbook.copy_worksheet`.

    Attributes:
        deep_copy_images: When ``True``, drawings reachable from a cloned sheet
            have their referenced ``xl/media/imageN.<ext>`` targets cloned into
            freshly numbered media parts. When ``False`` (default), the cloned
            drawing relationships point at the same image bytes as the source.
    """

    deep_copy_images: bool = False


def same_existing_path(left: str, right: str | None) -> bool:
    """Return whether two paths identify the same existing filesystem entry."""
    if right is None:
        return False
    try:
        return os.path.samefile(left, right)
    except OSError:
        return os.path.abspath(left) == os.path.abspath(right)


def xlsb_xls_via_tempfile(
    rust_cls: Any,
    data: bytes | bytearray | memoryview,
    *,
    suffix: str,
    permissive: bool,
) -> tuple[Any, str]:
    """Materialize bytes to a tempfile and open them with a binary backend.

    Args:
        rust_cls: Rust-backed binary workbook class.
        data: Workbook bytes supplied to ``load_workbook``.
        suffix: File extension used for the temporary file.
        permissive: Whether to pass permissive parsing through to the backend.

    Returns:
        A ``(rust_book, tempfile_path)`` pair. The caller owns cleanup.
    """
    import tempfile

    with tempfile.NamedTemporaryFile(prefix="wolfxl-", suffix=suffix, delete=False) as tmp:
        tmp.write(bytes(data))
        tmp_path = tmp.name

    opener = rust_cls.open
    try:
        rust_book = opener(tmp_path, permissive)
    except TypeError:
        rust_book = opener(tmp_path)
    return rust_book, tmp_path


def build_xlsx_wb(
    cls: type,
    *,
    rust_reader: Any,
    rust_patcher: Any | None,
    data_only: bool,
    read_only: bool,
    source_path: str | None,
) -> Any:
    """Wire up read/modify-mode workbook fields shared by xlsx inputs.

    Args:
        cls: Workbook class to instantiate without calling ``__init__``.
        rust_reader: Open ``CalamineStyledBook``-compatible reader.
        rust_patcher: Optional open ``XlsxPatcher`` for modify mode.
        data_only: Whether formula cells should expose cached values.
        read_only: Whether streaming read mode was explicitly requested.
        source_path: Source path, or ``None`` for bytes-backed readers.

    Returns:
        A workbook instance with sheet proxies and pending queues initialized.
    """
    wb = object.__new__(cls)
    wb._rust_writer = None
    wb._rust_patcher = rust_patcher
    wb._rust_reader = rust_reader
    wb._data_only = data_only
    wb._iso_dates = False
    wb.template = False
    wb.encoding = "utf-8"
    wb._rich_text = False
    wb._evaluator = None
    wb._read_only = read_only
    wb._source_path = source_path
    wb._format = "xlsx"
    _initialize_sheet_proxies(wb, rust_reader)
    initialize_pending_state(wb)
    return wb


def build_xlsb_xls_wb(
    cls: type,
    *,
    rust_book: Any,
    fmt: str,
    data_only: bool,
    source_path: str | None,
) -> Any:
    """Wire up read-mode workbook fields shared by xlsb and xls inputs."""
    wb = object.__new__(cls)
    wb._rust_writer = None
    wb._rust_patcher = None
    wb._rust_reader = rust_book
    wb._data_only = data_only
    wb._iso_dates = False
    wb.template = False
    wb.encoding = "utf-8"
    wb._rich_text = False
    wb._evaluator = None
    wb._read_only = False
    wb._source_path = source_path
    wb._format = fmt
    _initialize_sheet_proxies(wb, rust_book)
    initialize_pending_state(wb)
    return wb


def _initialize_sheet_proxies(wb: Any, rust_book: Any) -> None:
    """Attach worksheet proxies from a Rust reader's tab list."""
    names = [str(n) for n in rust_book.sheet_names()]
    wb._sheet_names = names
    wb._sheets = {name: Worksheet(wb, name) for name in names}


def initialize_pending_state(wb: Any) -> None:
    """Initialize workbook caches and pending mutation queues."""
    wb._properties_cache = None
    wb._custom_doc_props_cache = None
    wb._properties_dirty = False
    wb._defined_names_cache = None
    wb._named_styles_registry = None
    wb._style_names_cache = None
    wb._pending_defined_names = {}
    wb._security = None
    wb._file_sharing = None
    wb._security_loaded = False
    wb._pending_security_update = False
    wb._workbook_properties_cache = None
    wb._calc_properties_cache = None
    wb._views_cache = None
    wb._pending_axis_shifts = []
    wb._pending_range_moves = []
    wb._pending_sheet_copies = []
    wb._pending_chart_adds = {}
    wb._pending_pivot_caches = []
    wb._next_pivot_cache_id = 0
    wb._pending_slicer_caches = []
    wb._next_slicer_cache_id = 0
    wb.copy_options = CopyOptions()
