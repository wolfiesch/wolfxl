"""Worksheet structural operation queueing helpers."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from wolfxl.utils.cell import column_index_from_string, range_boundaries

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def coerce_col_idx(idx: int | str, op: str) -> int:
    """Accept either a 1-based int or an Excel column letter for col ops."""
    if isinstance(idx, str):
        try:
            col_idx = column_index_from_string(idx)
        except Exception as exc:
            raise ValueError(
                f"{op}: idx {idx!r} is not a valid column letter"
            ) from exc
    elif isinstance(idx, int) and not isinstance(idx, bool):
        col_idx = idx
    else:
        raise ValueError(f"{op}: idx must be int or str, got {idx!r}")
    if col_idx < 1:
        raise ValueError(f"{op}: idx must be >= 1, got {idx!r}")
    return col_idx


def insert_rows(ws: Worksheet, idx: int, amount: int = 1) -> None:
    """Queue an insert-rows operation for modify-mode save processing."""
    if not isinstance(idx, int) or idx < 1:
        raise ValueError(
            f"insert_rows: idx must be a positive integer (>=1), got {idx!r}"
        )
    if not isinstance(amount, int) or amount < 1:
        raise ValueError(
            f"insert_rows: amount must be a positive integer (>=1), got {amount!r}"
        )
    ws._workbook._pending_axis_shifts.append((ws.title, "row", idx, amount))  # noqa: SLF001


def delete_rows(ws: Worksheet, idx: int, amount: int = 1) -> None:
    """Queue a delete-rows operation for modify-mode save processing."""
    if not isinstance(idx, int) or idx < 1:
        raise ValueError(
            f"delete_rows: idx must be a positive integer (>=1), got {idx!r}"
        )
    if not isinstance(amount, int) or amount < 1:
        raise ValueError(
            f"delete_rows: amount must be a positive integer (>=1), got {amount!r}"
        )
    ws._workbook._pending_axis_shifts.append((ws.title, "row", idx, -amount))  # noqa: SLF001


def insert_cols(ws: Worksheet, idx: int | str, amount: int = 1) -> None:
    """Queue an insert-columns operation for modify-mode save processing."""
    col_idx = coerce_col_idx(idx, "insert_cols")
    if not isinstance(amount, int) or amount < 0:
        raise ValueError(
            f"insert_cols: amount must be an integer >= 0, got {amount!r}"
        )
    if amount == 0:
        return
    ws._workbook._pending_axis_shifts.append((ws.title, "col", col_idx, amount))  # noqa: SLF001


def delete_cols(ws: Worksheet, idx: int | str, amount: int = 1) -> None:
    """Queue a delete-columns operation for modify-mode save processing."""
    col_idx = coerce_col_idx(idx, "delete_cols")
    if not isinstance(amount, int) or amount < 0:
        raise ValueError(
            f"delete_cols: amount must be an integer >= 0, got {amount!r}"
        )
    if amount == 0:
        return
    ws._workbook._pending_axis_shifts.append((ws.title, "col", col_idx, -amount))  # noqa: SLF001


def move_range(
    ws: Worksheet,
    cell_range: Any,
    rows: int = 0,
    cols: int = 0,
    translate: bool = False,
) -> None:
    """Queue a rectangular range move operation for modify-mode save processing."""
    if not isinstance(rows, int) or isinstance(rows, bool):
        raise TypeError(
            f"move_range: rows must be an int, got {type(rows).__name__}"
        )
    if not isinstance(cols, int) or isinstance(cols, bool):
        raise TypeError(
            f"move_range: cols must be an int, got {type(cols).__name__}"
        )
    if not isinstance(cell_range, str):
        cell_range = str(cell_range)
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    except Exception as exc:
        raise ValueError(
            f"move_range: cell_range must be a valid A1 range string, "
            f"got {cell_range!r}: {exc}"
        ) from exc
    if min_col is None or min_row is None or max_col is None or max_row is None:
        raise ValueError(
            f"move_range: cell_range must have all four corners "
            f"(rows + cols), got {cell_range!r}"
        )
    if rows == 0 and cols == 0:
        return

    dst_min_row = min_row + rows
    dst_max_row = max_row + rows
    dst_min_col = min_col + cols
    dst_max_col = max_col + cols
    if dst_min_row < 1 or dst_max_row > 1_048_576:
        raise ValueError(
            f"move_range: destination row range "
            f"[{dst_min_row}, {dst_max_row}] is out of bounds "
            f"(must be in [1, 1048576])"
        )
    if dst_min_col < 1 or dst_max_col > 16_384:
        raise ValueError(
            f"move_range: destination column range "
            f"[{dst_min_col}, {dst_max_col}] is out of bounds "
            f"(must be in [1, 16384])"
        )
    ws._workbook._pending_range_moves.append(  # noqa: SLF001
        (
            ws.title,
            int(min_col),
            int(min_row),
            int(max_col),
            int(max_row),
            int(rows),
            int(cols),
            bool(translate),
        )
    )
