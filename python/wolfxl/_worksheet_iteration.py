"""Worksheet row and column iteration helpers."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from wolfxl._utils import rowcol_to_a1

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def iter_rows(
    ws: Worksheet,
    min_row: int | None = None,
    max_row: int | None = None,
    min_col: int | None = None,
    max_col: int | None = None,
    values_only: bool = False,
) -> Iterator[tuple[Any, ...]]:
    """Iterate worksheet rows with streaming and bulk-read fast paths."""
    workbook = ws._workbook  # noqa: SLF001
    if workbook._rust_reader is not None and getattr(workbook, "_source_path", None):  # noqa: SLF001
        from wolfxl._streaming import should_auto_stream, stream_iter_rows

        stream_now = bool(getattr(workbook, "_read_only", False)) or should_auto_stream(ws)
        if stream_now:
            yield from stream_iter_rows(
                ws,
                min_row,
                max_row,
                min_col,
                max_col,
                values_only=values_only,
            )
            return

    if values_only and workbook._rust_reader is not None:  # noqa: SLF001
        yield from iter_rows_bulk(ws, min_row, max_row, min_col, max_col)
        return

    row_min = min_row or 1
    row_max = max_row or ws._max_row()  # noqa: SLF001
    col_min = min_col or 1
    col_max = max_col or ws._max_col()  # noqa: SLF001

    for row in range(row_min, row_max + 1):
        if values_only:
            yield tuple(
                ws._get_or_create_cell(row, col).value  # noqa: SLF001
                for col in range(col_min, col_max + 1)
            )
        else:
            yield tuple(
                ws._get_or_create_cell(row, col)  # noqa: SLF001
                for col in range(col_min, col_max + 1)
            )


def iter_cols(
    ws: Worksheet,
    min_col: int | None = None,
    max_col: int | None = None,
    min_row: int | None = None,
    max_row: int | None = None,
    values_only: bool = False,
) -> Iterator[tuple[Any, ...]]:
    """Iterate worksheet columns with a bulk-read fast path when possible."""
    if values_only and ws._workbook._rust_reader is not None:  # noqa: SLF001
        yield from iter_cols_bulk(ws, min_col, max_col, min_row, max_row)
        return

    row_min = min_row or 1
    row_max = max_row or ws._max_row()  # noqa: SLF001
    col_min = min_col or 1
    col_max = max_col or ws._max_col()  # noqa: SLF001

    for col in range(col_min, col_max + 1):
        if values_only:
            yield tuple(
                ws._get_or_create_cell(row, col).value  # noqa: SLF001
                for row in range(row_min, row_max + 1)
            )
        else:
            yield tuple(
                ws._get_or_create_cell(row, col)  # noqa: SLF001
                for row in range(row_min, row_max + 1)
            )


def iter_cols_bulk(
    ws: Worksheet,
    min_col: int | None,
    max_col: int | None,
    min_row: int | None,
    max_row: int | None,
) -> Iterator[tuple[Any, ...]]:
    """Bulk-read column values through one Rust FFI call, then transpose."""
    from wolfxl._cell import _payload_to_python

    reader = ws._workbook._rust_reader  # noqa: SLF001
    sheet = ws._title  # noqa: SLF001
    data_only = getattr(ws._workbook, "_data_only", False)  # noqa: SLF001

    row_min = min_row or 1
    row_max = max_row or ws._max_row()  # noqa: SLF001
    col_min = min_col or 1
    col_max = max_col or ws._max_col()  # noqa: SLF001
    range_str = f"{rowcol_to_a1(row_min, col_min)}:{rowcol_to_a1(row_max, col_max)}"

    use_plain = hasattr(reader, "read_sheet_values_plain")
    if use_plain:
        rows = reader.read_sheet_values_plain(sheet, range_str, data_only)
    else:
        rows = reader.read_sheet_values(sheet, range_str, data_only)

    if not rows:
        return

    expected_cols = col_max - col_min + 1
    expected_rows = row_max - row_min + 1
    normalized: list[list[Any]] = []
    for row in rows:
        if use_plain:
            values = list(row)
        else:
            values = [_payload_to_python(cell) for cell in row]
        width = len(values)
        if width >= expected_cols:
            normalized.append(values[:expected_cols])
        else:
            normalized.append(values + [None] * (expected_cols - width))

    while len(normalized) < expected_rows:
        normalized.append([None] * expected_cols)

    for col_offset in range(expected_cols):
        yield tuple(
            normalized[row_offset][col_offset]
            for row_offset in range(expected_rows)
        )


def iter_rows_bulk(
    ws: Worksheet,
    min_row: int | None,
    max_row: int | None,
    min_col: int | None,
    max_col: int | None,
) -> Iterator[tuple[Any, ...]]:
    """Bulk-read row values through one Rust FFI call."""
    from wolfxl._cell import _payload_to_python

    reader = ws._workbook._rust_reader  # noqa: SLF001
    sheet = ws._title  # noqa: SLF001
    data_only = getattr(ws._workbook, "_data_only", False)  # noqa: SLF001

    row_min = min_row or 1
    row_max = max_row or ws._max_row()  # noqa: SLF001
    col_min = min_col or 1
    col_max = max_col or ws._max_col()  # noqa: SLF001
    range_str = f"{rowcol_to_a1(row_min, col_min)}:{rowcol_to_a1(row_max, col_max)}"

    use_plain = hasattr(reader, "read_sheet_values_plain")
    if use_plain:
        rows = reader.read_sheet_values_plain(sheet, range_str, data_only)
    else:
        rows = reader.read_sheet_values(sheet, range_str, data_only)

    if not rows:
        return

    expected_cols = col_max - col_min + 1
    for row in rows:
        if use_plain:
            values = list(row)
        else:
            values = [_payload_to_python(cell) for cell in row]
        width = len(values)
        if width >= expected_cols:
            yield tuple(values[:expected_cols])
        else:
            yield tuple(values) + (None,) * (expected_cols - width)
