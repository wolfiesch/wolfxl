"""WolfXL — fast, openpyxl-compatible Excel I/O backed by Rust.

Usage::

    from wolfxl import load_workbook, Workbook, Font, PatternFill

    # Read
    wb = load_workbook("data.xlsx")
    ws = wb["Sheet1"]
    print(ws["A1"].value, ws["A1"].font.bold)

    # Write
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Hello"
    ws["A1"].font = Font(bold=True)
    wb.save("out.xlsx")
"""

from __future__ import annotations

import os
from typing import IO

from wolfxl._cell import Cell
from wolfxl._rust import __version__, classify_format

# File-format detector (xlsx / xlsb / xls / ods / unknown), distinct from the
# existing ``classify_format`` cell-format classifier. Re-exported here so
# callers can use a stable ``wolfxl.classify_file_format(...)`` import without
# depending on the private ``wolfxl._rust`` module.
try:
    from wolfxl._rust import classify_file_format  # type: ignore[attr-defined]
except ImportError:  # pragma: no cover - wheels should expose this
    classify_file_format = None  # type: ignore[assignment]
from wolfxl._styles import Alignment, Border, Color, Font, PatternFill, Side
from wolfxl._workbook import CopyOptions, Workbook
from wolfxl._worksheet import Worksheet
from wolfxl.utils.exceptions import InvalidFileException

__all__ = [
    "__version__",
    "Alignment",
    "Border",
    "Cell",
    "Color",
    "CopyOptions",
    "Font",
    "PatternFill",
    "Side",
    "Workbook",
    "Worksheet",
    "classify_file_format",
    "classify_format",
    "load_workbook",
]


def load_workbook(
    filename: (
        str
        | os.PathLike[str]
        | bytes
        | bytearray
        | memoryview
        | IO[bytes]
    ),
    read_only: bool = False,
    data_only: bool = False,
    keep_links: bool = True,
    modify: bool = False,
    permissive: bool = False,
    rich_text: bool = False,
    password: str | bytes | None = None,
) -> Workbook:
    """Open a workbook for reading, streaming, or modify-mode saves.

    Args:
        filename: Path, bytes-like object, or binary file-like object
            containing an ``.xlsx``, ``.xlsb``, or ``.xls`` workbook.
        read_only: Enable the streaming row reader for ``.xlsx`` files.
            Streaming cells are immutable.
        data_only: Return cached formula results when present.
        keep_links: Compatibility shim accepted for openpyxl-shaped call sites.
        modify: Enable read-modify-write mode for ``.xlsx`` files. Modified
            cells and supported metadata are saved while preserving unchanged
            workbook parts where possible.
        permissive: Fall back to worksheet relationships when workbook sheet
            metadata is malformed but recoverable.
        rich_text: Return structured rich-text values for cells that carry
            shared-string runs.
        password: Decrypt OOXML-encrypted ``.xlsx`` inputs with the optional
            ``wolfxl[encrypted]`` dependency.

    Returns:
        A :class:`Workbook` in read, streaming, or modify mode.

    Raises:
        InvalidFileException: If the input format cannot be identified.
        NotImplementedError: If the requested mode is unsupported for the
            detected format, such as modify mode for ``.xlsb``.
        ValueError: If an encrypted workbook needs a password or the supplied
            password cannot decrypt it.
    """
    from wolfxl._loader import classify_input
    from wolfxl._workbook_sources import open_workbook_source

    fmt, data, path = classify_input(filename)

    # OOXML-encrypted .xlsx files arrive as OLE CFB envelopes that the
    # sniffer reports as ``"encrypted"``.  Once decrypted via
    # msoffcrypto they're plain xlsx, so we route them through the
    # encrypted constructor unconditionally; if the caller forgot to
    # pass ``password=``, surface a clear error before _from_encrypted
    # raises something more cryptic.
    if fmt == "encrypted":
        if password is None:
            raise ValueError(
                "this workbook is OOXML-encrypted; pass password= to "
                "load_workbook() (install with pip install wolfxl[encrypted])"
            )
        # _from_encrypted will produce a plain xlsx workbook.
        fmt = "xlsx"

    # Format-specific guards — surface clear errors *before* we try to
    # materialise a backend that doesn't exist for the requested mode.
    if fmt in ("xlsb", "xls"):
        if modify:
            raise NotImplementedError(
                f".{fmt} files are read-only in wolfxl; load + transcribe to "
                ".xlsx then save: load via load_workbook(path), reconstruct "
                "as a fresh Workbook(), wb.save('out.xlsx')"
            )
        if read_only:
            raise NotImplementedError(
                f"streaming read_only mode is xlsx-only; .{fmt} files load "
                "whole-sheet"
            )
        if password is not None:
            raise NotImplementedError(
                f"password reads are xlsx-only (msoffcrypto-tool); .{fmt} "
                "encryption is out of scope"
            )

    if fmt == "ods":
        raise NotImplementedError(
            ".ods files are not supported by wolfxl; use openpyxl/odfpy "
            "for OpenDocument reads"
        )

    if fmt == "unknown":
        source_kind = "path" if path else "bytes"
        raise InvalidFileException(
            f"could not determine file format from {source_kind}; "
            "expected xlsx/xlsb/xls"
        )

    wb = open_workbook_source(
        Workbook,
        fmt=fmt,
        path=path,
        data=data,
        password=password,
        data_only=data_only,
        permissive=permissive,
        modify=modify,
        read_only=read_only,
    )

    wb._rich_text = rich_text  # noqa: SLF001
    return wb
