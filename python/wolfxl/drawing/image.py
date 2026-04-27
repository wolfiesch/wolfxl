"""openpyxl-shaped ``Image`` class — Sprint Λ Pod-β (RFC-045).

This module replaces the historical ``_make_stub`` placeholder with a
real, working class. ``Image(filename_or_bytesio)`` parses image bytes,
sniffs the format from byte-magic headers, and exposes ``.width``,
``.height``, ``.format``, ``.anchor``, ``.path`` attrs that match
openpyxl 3.1.x.

The four supported formats (PNG, JPEG, GIF, BMP) are detected with
pure-Python sniffers — no Pillow dep required. If Pillow is installed
the module falls back to it for JPEG dimensions which require walking
SOF markers; without Pillow we still parse the four other shapes.

The instance owns the raw bytes (``_data``) and the format string.
``Worksheet.add_image`` consumes both: write mode hands them to the
native writer's ``add_image`` pymethod; modify mode routes them through
the ``XlsxPatcher.queue_image_add`` plumbing.
"""

from __future__ import annotations

import io
import os
import struct
from typing import Any


# ---------------------------------------------------------------------------
# Format detection
# ---------------------------------------------------------------------------

def _sniff_format(data: bytes) -> str:
    """Return ``"png"``/``"jpeg"``/``"gif"``/``"bmp"`` from leading bytes.

    Raises ``ValueError`` for anything we don't recognise. Matches
    openpyxl's accepted extensions on the write path (``.png``/``.jpeg``/
    ``.gif``/``.bmp``).
    """
    if len(data) < 8:
        raise ValueError(f"image data too short ({len(data)} bytes) to sniff format")
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if data[:3] == b"\xff\xd8\xff":
        return "jpeg"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if data[:2] == b"BM":
        return "bmp"
    raise ValueError(
        f"unrecognised image format (first 8 bytes: {data[:8]!r}); "
        "wolfxl supports PNG / JPEG / GIF / BMP"
    )


def _png_dimensions(data: bytes) -> tuple[int, int]:
    """Read PNG IHDR (always at byte 16-24, big-endian width+height)."""
    if len(data) < 24:
        raise ValueError("PNG too short to read IHDR")
    width = struct.unpack(">I", data[16:20])[0]
    height = struct.unpack(">I", data[20:24])[0]
    return width, height


def _gif_dimensions(data: bytes) -> tuple[int, int]:
    """Read GIF logical screen descriptor (bytes 6-10, little-endian)."""
    if len(data) < 10:
        raise ValueError("GIF too short to read screen descriptor")
    width = struct.unpack("<H", data[6:8])[0]
    height = struct.unpack("<H", data[8:10])[0]
    return width, height


def _bmp_dimensions(data: bytes) -> tuple[int, int]:
    """Read BMP DIB header (bytes 18-26, little-endian, height may be signed)."""
    if len(data) < 26:
        raise ValueError("BMP too short to read DIB header")
    width = struct.unpack("<i", data[18:22])[0]
    height = struct.unpack("<i", data[22:26])[0]
    # BMP allows negative height (top-down storage); display dims are abs.
    return abs(width), abs(height)


def _jpeg_dimensions(data: bytes) -> tuple[int, int]:
    """Walk JPEG SOF0/SOF1/SOF2/SOF3 markers; no Pillow dep.

    SOF markers are 0xFFC0..0xFFC3 (and 0xFFC5..0xFFCF excluding the
    DHT/JPG/DAC reserved ones).  Each SOF has the shape:

        FF CN  LL LL  PP  YY YY  XX XX  ...

    where YY is height and XX is width. We start after the SOI (0xFFD8)
    and skip every other marker by reading its 2-byte length.
    """
    if len(data) < 4 or data[:2] != b"\xff\xd8":
        raise ValueError("not a JPEG (missing SOI)")
    i = 2
    while i + 4 <= len(data):
        if data[i] != 0xFF:
            raise ValueError(f"JPEG: unexpected byte at offset {i}: {data[i]:#x}")
        # Skip pad bytes (FF FF ...).
        while i < len(data) and data[i] == 0xFF:
            i += 1
        if i >= len(data):
            break
        marker = data[i]
        i += 1
        # Standalone markers (no length): SOI, EOI, RSTn, TEM.
        if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7 or marker == 0x01:
            continue
        if i + 2 > len(data):
            break
        length = struct.unpack(">H", data[i : i + 2])[0]
        # SOF0..SOF3, SOF5..SOF7, SOF9..SOF11, SOF13..SOF15.
        if (
            0xC0 <= marker <= 0xCF
            and marker not in (0xC4, 0xC8, 0xCC)
        ):
            # Layout: LL LL PP YY YY XX XX ...
            if i + 7 > len(data):
                break
            height = struct.unpack(">H", data[i + 3 : i + 5])[0]
            width = struct.unpack(">H", data[i + 5 : i + 7])[0]
            return width, height
        i += length
    raise ValueError("JPEG: no SOF marker found; cannot determine dimensions")


def _dimensions(fmt: str, data: bytes) -> tuple[int, int]:
    """Dispatch to the format-specific dimension reader."""
    if fmt == "png":
        return _png_dimensions(data)
    if fmt == "jpeg":
        return _jpeg_dimensions(data)
    if fmt == "gif":
        return _gif_dimensions(data)
    if fmt == "bmp":
        return _bmp_dimensions(data)
    raise ValueError(f"no dimension reader for format: {fmt!r}")


# ---------------------------------------------------------------------------
# Image
# ---------------------------------------------------------------------------

class Image:
    """An image to be embedded in a worksheet.

    Mirrors openpyxl's ``openpyxl.drawing.image.Image`` constructor
    signature: accepts a path-like or a binary file-like (``BytesIO``,
    open file, anything with ``.read()``).

    Parameters
    ----------
    img : str | os.PathLike | io.IOBase | bytes
        - ``str``/``PathLike``: opened and read; ``self.path`` is set.
        - file-like: ``.read()`` is called once; ``self.path`` is the
          object's ``.name`` if any, otherwise ``None``.
        - ``bytes``/``bytearray``: stored verbatim; ``self.path`` is
          ``None``.

    Attributes
    ----------
    path : str | None
        Source path on disk, or ``None`` for in-memory inputs.
    width : int
        Pixel width parsed from the image header.
    height : int
        Pixel height.
    format : str
        ``"png"``/``"jpeg"``/``"gif"``/``"bmp"``.
    anchor : str | _AnchorBase | None
        A1 cell ref (e.g. ``"B5"``) by default; settable to a
        ``TwoCellAnchor`` or ``AbsoluteAnchor`` instance from
        :mod:`wolfxl.drawing.spreadsheet_drawing`. ``None`` until
        ``Worksheet.add_image`` runs.
    """

    def __init__(self, img: Any) -> None:
        # Resolve the input to (raw_bytes, optional_path).
        path: str | None = None
        data: bytes
        if isinstance(img, (bytes, bytearray, memoryview)):
            data = bytes(img)
        elif hasattr(img, "read"):
            # File-like input.
            raw = img.read()
            if isinstance(raw, str):
                raise TypeError(
                    "Image(file-like): .read() must return bytes, got str"
                )
            data = bytes(raw)
            # Best-effort path recovery (BytesIO has no .name).
            path = getattr(img, "name", None)
            if path is not None and not isinstance(path, str):
                path = None
        else:
            # Treat as path-like.
            path = os.fspath(img)
            with open(path, "rb") as fh:
                data = fh.read()

        if not data:
            raise ValueError("Image input is empty")

        fmt = _sniff_format(data)
        try:
            width, height = _dimensions(fmt, data)
        except Exception as exc:
            raise ValueError(f"could not parse {fmt} dimensions: {exc}") from exc

        self.path: str | None = path
        self._data: bytes = data
        self.format: str = fmt
        self.width: int = int(width)
        self.height: int = int(height)
        self.anchor: Any = None  # set by Worksheet.add_image

    @property
    def ref(self) -> bytes:
        """Raw image bytes (openpyxl uses ``.ref`` internally)."""
        return self._data

    def __repr__(self) -> str:
        path_repr = self.path or "<bytes>"
        return (
            f"<wolfxl.drawing.image.Image {path_repr!r} "
            f"{self.format} {self.width}x{self.height}>"
        )


__all__ = ["Image"]
