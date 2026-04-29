"""Worksheet append and bulk-write buffer helpers."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def materialize_append_buffer(ws: Worksheet) -> None:
    """Convert a worksheet append buffer into dirty Cell objects."""
    start = ws._append_buffer_start  # noqa: SLF001
    buffer = ws._append_buffer  # noqa: SLF001
    if not buffer:
        return
    ws._append_buffer = []  # noqa: SLF001
    for row_offset, row_values in enumerate(buffer):
        row_index = start + row_offset
        for col_index, value in enumerate(row_values, start=1):
            ws.cell(row=row_index, column=col_index, value=value)


def materialize_bulk_writes(ws: Worksheet) -> None:
    """Convert bulk write buffers into dirty Cell objects."""
    writes = ws._bulk_writes  # noqa: SLF001
    if not writes:
        return
    ws._bulk_writes = []  # noqa: SLF001
    for grid, start_row, start_col in writes:
        for row_offset, row_values in enumerate(grid):
            for col_offset, value in enumerate(row_values):
                if value is not None:
                    ws.cell(
                        row=start_row + row_offset,
                        column=start_col + col_offset,
                        value=value,
                    )


def extract_non_batchable(
    grid: list[list[Any]],
    start_row: int,
    start_col: int,
) -> list[tuple[int, int, Any]]:
    """Extract values that require per-cell writes from a batch grid."""
    individual: list[tuple[int, int, Any]] = []
    for row_offset, row_values in enumerate(grid):
        for col_offset, value in enumerate(row_values):
            if value is not None and (
                isinstance(value, bool)
                or (isinstance(value, str) and value.startswith("="))
                or not isinstance(value, (int, float, str))
            ):
                individual.append(
                    (start_row + row_offset, start_col + col_offset, value)
                )
                row_values[col_offset] = None
    return individual


def batch_write_dicts(
    ws: Worksheet,
    batch_fn: Any,
    entries: list[tuple[int, int, dict[str, Any]]],
) -> None:
    """Build a bounding-box grid of dicts and call a batch Rust method."""
    min_row = entries[0][0]
    min_col = entries[0][1]
    max_row = min_row
    max_col = min_col
    for row, col, _payload in entries:
        if row < min_row:
            min_row = row
        if row > max_row:
            max_row = row
        if col < min_col:
            min_col = col
        if col > max_col:
            max_col = col

    num_rows = max_row - min_row + 1
    num_cols = max_col - min_col + 1
    grid: list[list[Any]] = [[None] * num_cols for _ in range(num_rows)]
    for row, col, payload in entries:
        grid[row - min_row][col - min_col] = payload

    from wolfxl._utils import rowcol_to_a1

    start = rowcol_to_a1(min_row, min_col)
    batch_fn(ws._title, start, grid)  # noqa: SLF001
