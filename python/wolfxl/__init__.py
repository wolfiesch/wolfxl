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

# Sprint Κ Pod-α: file-format detector (xlsx / xlsb / xls / ods / unknown).
# Distinct from the long-standing ``classify_format`` SynthGL archetype
# classifier above. Re-exported here so callers can use a stable
# ``wolfxl.classify_file_format(...)`` import without needing to drop into
# the private ``wolfxl._rust`` module.
try:
    from wolfxl._rust import classify_file_format  # type: ignore[attr-defined]
except ImportError:  # pragma: no cover — Pod-α should always expose this
    classify_file_format = None  # type: ignore[assignment]
from wolfxl._styles import Alignment, Border, Color, Font, PatternFill, Side
from wolfxl._workbook import CopyOptions, Workbook
from wolfxl._worksheet import Worksheet

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
    """Open an .xlsx file for reading or modification.

    Parameters
    ----------
    modify : bool
        If True, enable read-modify-write mode.  Values and formats can be
        changed and saved back to disk via ``wb.save(path)``.  Uses the WolfXL
        engine (surgical ZIP patching) instead of a full DOM rewrite.
    permissive : bool
        If True, fall back to the workbook rels graph when
        ``xl/workbook.xml``'s ``<sheets>`` block is empty or
        self-closing. Each worksheet relationship target is registered
        under a synthesized title (``Sheet1``, ``Sheet2``, ...). This
        unblocks workflows that need to operate on technically-malformed
        (but Excel-tolerant) workbooks — e.g. a self-closing
        ``<sheets/>`` whose rels still reference
        ``xl/worksheets/sheet1.xml``. Default is ``False`` so well-formed
        inputs round-trip unchanged. Added in Sprint Θ Pod-A; tracked in
        tests/parity/KNOWN_GAPS.md (RFC-035 cross-RFC composition bug
        #4).
    password : str | bytes | None
        Decryption password for OOXML-encrypted workbooks. When provided,
        wolfxl lazy-imports ``msoffcrypto-tool`` (install via
        ``pip install wolfxl[encrypted]``), decrypts the file into an
        in-memory buffer, then dispatches through the standard reader
        (or patcher, when ``modify=True``). On a non-encrypted file the
        password is silently ignored, matching openpyxl's behaviour.
        Wrong / missing passwords surface as ``ValueError``. Write-side
        encryption is **not** supported; saving a workbook opened with
        ``password=`` produces a plaintext output.

    rich_text : bool
        Sprint Ι Pod-α: if True, ``Cell.value`` returns a
        :class:`wolfxl.cell.rich_text.CellRichText` for cells whose
        backing string carries `<r>` runs.  Default is ``False`` so
        existing call sites keep returning flattened ``str`` values
        (matches openpyxl 3.x's default).  Use ``Cell.rich_text``
        directly to read structured runs without flipping this flag.

    ``read_only=True`` (Sprint Ι Pod-β) activates the SAX streaming
    read path: ``Worksheet.iter_rows`` becomes a true generator that
    walks ``xl/worksheets/sheetN.xml`` one row at a time without
    materializing the whole sheet. Cells in this mode are immutable —
    assignment raises ``RuntimeError`` immediately. The flag also
    auto-engages transparently for sheets with > 50000 rows so callers
    don't have to opt in just to scale to large workbooks.

    ``data_only=True`` returns cached formula results when they exist.
    ``keep_links`` remains a no-op compatibility shim.

    Sprint Κ Pod-β: ``filename`` may now also be raw ``bytes`` /
    ``bytearray`` / ``memoryview`` or any file-like object whose
    ``.read()`` returns bytes (e.g. :class:`io.BytesIO`). The format is
    sniffed from the leading magic bytes; .xlsb and .xls inputs route
    through dedicated calamine backends (read-only — see the
    :class:`Workbook._format` attribute and the ``modify``/``read_only``
    guards below).
    """
    from wolfxl._loader import classify_input

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
        raise ValueError(
            f"could not determine file format from {source_kind}; "
            "expected xlsx/xlsb/xls"
        )

    # Dispatch.
    if fmt == "xlsx":
        if password is not None:
            wb = Workbook._from_encrypted(  # noqa: SLF001
                path=path,
                data=data,
                password=password,
                data_only=data_only,
                permissive=permissive,
                modify=modify,
            )
        elif data is not None:
            # Bytes / BytesIO input: dispatch through the bytes shim.
            wb = Workbook._from_bytes(
                data,
                data_only=data_only,
                permissive=permissive,
                modify=modify,
                read_only=read_only,
            )
        elif modify:
            wb = Workbook._from_patcher(  # noqa: SLF001
                path, data_only=data_only, permissive=permissive
            )
        else:
            wb = Workbook._from_reader(  # noqa: SLF001
                path,
                data_only=data_only,
                permissive=permissive,
                read_only=read_only,
            )
    elif fmt == "xlsb":
        wb = Workbook._from_xlsb(  # noqa: SLF001
            path=path, data=data, data_only=data_only, permissive=permissive,
        )
    elif fmt == "xls":
        wb = Workbook._from_xls(  # noqa: SLF001
            path=path, data=data, data_only=data_only, permissive=permissive,
        )
    else:  # pragma: no cover — defensive; classify_input only emits the above.
        raise ValueError(f"unsupported file format: {fmt!r}")

    wb._rich_text = rich_text  # noqa: SLF001
    return wb
