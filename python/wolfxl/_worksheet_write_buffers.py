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
