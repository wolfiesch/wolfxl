"""Workbook source-opening helpers."""

from __future__ import annotations

from typing import Any

from wolfxl._workbook_state import (
    build_xlsb_xls_wb,
    build_xlsx_wb,
    xlsb_xls_via_tempfile,
)


def from_reader(
    cls: type,
    path: str,
    *,
    data_only: bool = False,
    permissive: bool = False,
    read_only: bool = False,
) -> Any:
    """Open an existing .xlsx file in read mode."""
    from wolfxl import _rust

    return build_xlsx_wb(
        cls,
        rust_reader=_rust.CalamineStyledBook.open(path, permissive),
        rust_patcher=None,
        data_only=data_only,
        read_only=read_only,
        source_path=path,
    )


def _open_plain_xlsx_source(
    cls: type,
    *,
    path: str | None,
    data: bytes | bytearray | memoryview | None,
    data_only: bool,
    permissive: bool,
    modify: bool,
    read_only: bool,
) -> Any:
    """Open an unencrypted XLSX path or byte buffer with the requested mode."""
    if path is not None:
        if modify:
            return from_patcher(cls, path, data_only=data_only, permissive=permissive)
        return from_reader(
            cls,
            path,
            data_only=data_only,
            permissive=permissive,
            read_only=read_only,
        )
    return from_bytes(
        cls,
        bytes(data),  # type: ignore[arg-type]
        data_only=data_only,
        permissive=permissive,
        modify=modify,
        read_only=read_only,
    )


def from_encrypted(
    cls: type,
    path: str | None = None,
    *,
    data: bytes | bytearray | memoryview | None = None,
    password: str | bytes,
    data_only: bool = False,
    permissive: bool = False,
    modify: bool = False,
    read_only: bool = False,
) -> Any:
    """Open an OOXML-encrypted .xlsx via msoffcrypto-tool."""
    if (path is None) == (data is None):
        raise TypeError("_from_encrypted requires exactly one of path / data")

    if path is not None:
        with open(path, "rb") as fp:
            is_plain_xlsx = fp.read(4).startswith(b"PK")
    else:
        is_plain_xlsx = bytes(data).startswith(b"PK")  # type: ignore[arg-type]

    if is_plain_xlsx:
        return _open_plain_xlsx_source(
            cls,
            path=path,
            data=data,
            data_only=data_only,
            permissive=permissive,
            modify=modify,
            read_only=read_only,
        )

    try:
        import msoffcrypto  # type: ignore[import-not-found]
    except ImportError as exc:
        raise ImportError(
            "password reads require msoffcrypto-tool; install with: "
            "pip install wolfxl[encrypted]"
        ) from exc

    import io

    pw_str = password.decode("utf-8") if isinstance(password, bytes) else password
    if path is not None:
        src_fp = open(path, "rb")  # noqa: SIM115 - closed in finally
    else:
        src_fp = io.BytesIO(bytes(data))  # type: ignore[arg-type]

    try:
        office = msoffcrypto.OfficeFile(src_fp)
        try:
            is_encrypted = office.is_encrypted()
        except Exception:
            is_encrypted = False

        if not is_encrypted:
            return _open_plain_xlsx_source(
                cls,
                path=path,
                data=data,
                data_only=data_only,
                permissive=permissive,
                modify=modify,
                read_only=read_only,
            )

        try:
            office.load_key(password=pw_str)
        except Exception as exc:
            raise ValueError(f"failed to load decryption key: {exc}") from exc

        buf = io.BytesIO()
        try:
            office.decrypt(buf)
        except Exception as exc:
            raise ValueError(
                f"failed to decrypt workbook (wrong password?): {exc}"
            ) from exc
        decrypted_bytes = buf.getvalue()
    finally:
        src_fp.close()

    return from_bytes(
        cls,
        decrypted_bytes,
        data_only=data_only,
        permissive=permissive,
        modify=modify,
        read_only=read_only,
    )


def from_bytes(
    cls: type,
    data: bytes | bytearray | memoryview,
    *,
    data_only: bool = False,
    permissive: bool = False,
    modify: bool = False,
    read_only: bool = False,
) -> Any:
    """Open an .xlsx blob from memory."""
    from wolfxl import _rust

    data_bytes = bytes(data)
    bytes_open = getattr(_rust.CalamineStyledBook, "open_from_bytes", None)
    needs_tempfile = modify or bytes_open is None

    if needs_tempfile:
        import tempfile

        with tempfile.NamedTemporaryFile(
            prefix="wolfxl-", suffix=".xlsx", delete=False
        ) as tmp:
            tmp.write(data_bytes)
            tmp_path = tmp.name

        if modify:
            wb = from_patcher(
                cls, tmp_path, data_only=data_only, permissive=permissive
            )
        else:
            wb = from_reader(
                cls,
                tmp_path,
                data_only=data_only,
                permissive=permissive,
                read_only=read_only,
            )
        wb._tempfile_path = tmp_path
        return wb

    return build_xlsx_wb(
        cls,
        rust_reader=bytes_open(data_bytes, permissive),
        rust_patcher=None,
        data_only=data_only,
        read_only=read_only,
        source_path=None,
    )


def from_patcher(
    cls: type,
    path: str,
    *,
    data_only: bool = False,
    permissive: bool = False,
) -> Any:
    """Open an existing .xlsx file in modify mode."""
    from wolfxl import _rust

    return build_xlsx_wb(
        cls,
        rust_reader=_rust.CalamineStyledBook.open(path, permissive),
        rust_patcher=_rust.XlsxPatcher.open(path, permissive),
        data_only=data_only,
        read_only=False,
        source_path=path,
    )


def from_xlsb(
    cls: type,
    *,
    path: str | None,
    data: bytes | None,
    data_only: bool = False,
    permissive: bool = False,
) -> Any:
    """Open an .xlsb workbook via ``CalamineXlsbBook``."""
    from wolfxl import _rust

    rust_cls = getattr(_rust, "CalamineXlsbBook", None)
    if rust_cls is None:
        raise NotImplementedError(
            ".xlsb reads require the CalamineXlsbBook backend "
            "from the wolfxl Rust extension."
        )

    if data is not None:
        rust_book, tmp_path = _open_binary_bytes(
            rust_cls, data, suffix=".xlsb", permissive=permissive
        )
        wb = build_xlsb_xls_wb(
            cls,
            rust_book=rust_book,
            fmt="xlsb",
            data_only=data_only,
            source_path=None,
        )
        if tmp_path is not None:
            wb._tempfile_path = tmp_path
        return wb

    rust_book = _open_binary_path(rust_cls, path, permissive=permissive)
    return build_xlsb_xls_wb(
        cls,
        rust_book=rust_book,
        fmt="xlsb",
        data_only=data_only,
        source_path=path,
    )


def from_xls(
    cls: type,
    *,
    path: str | None,
    data: bytes | None,
    data_only: bool = False,
    permissive: bool = False,
) -> Any:
    """Open a legacy .xls workbook via ``CalamineXlsBook``."""
    from wolfxl import _rust

    rust_cls = getattr(_rust, "CalamineXlsBook", None)
    if rust_cls is None:
        raise NotImplementedError(
            ".xls reads require the CalamineXlsBook backend "
            "from the wolfxl Rust extension."
        )

    if data is not None:
        rust_book, tmp_path = _open_binary_bytes(
            rust_cls, data, suffix=".xls", permissive=permissive
        )
        wb = build_xlsb_xls_wb(
            cls,
            rust_book=rust_book,
            fmt="xls",
            data_only=data_only,
            source_path=None,
        )
        if tmp_path is not None:
            wb._tempfile_path = tmp_path
        return wb

    rust_book = _open_binary_path(rust_cls, path, permissive=permissive)
    return build_xlsb_xls_wb(
        cls,
        rust_book=rust_book,
        fmt="xls",
        data_only=data_only,
        source_path=path,
    )


def _open_binary_bytes(
    rust_cls: Any,
    data: bytes,
    *,
    suffix: str,
    permissive: bool,
) -> tuple[Any, str | None]:
    bytes_open = getattr(rust_cls, "open_from_bytes", None)
    if bytes_open is None:
        rust_book, tmp_path = xlsb_xls_via_tempfile(
            rust_cls, data, suffix=suffix, permissive=permissive
        )
        return rust_book, tmp_path
    try:
        return bytes_open(data, permissive), None
    except TypeError:
        return bytes_open(data), None


def _open_binary_path(rust_cls: Any, path: str | None, *, permissive: bool) -> Any:
    opener = getattr(rust_cls, "open", None)
    if opener is None:
        raise NotImplementedError(
            f"{rust_cls.__name__}.open is not yet exposed by the Rust extension."
        )
    try:
        return opener(path, permissive)
    except TypeError:
        return opener(path)
