"""Write-mode worksheet cell flush helpers for the native Rust writer."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from wolfxl._utils import rowcol_to_a1

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def flush_to_writer(
    ws: Worksheet,
    writer: Any,
    python_value_to_payload: Any,
    font_to_format_dict: Any,
    fill_to_format_dict: Any,
    alignment_to_format_dict: Any,
    border_to_rust_dict: Any,
    rich_text_to_runs_payload: Any,
) -> None:
    """Flush dirty worksheet cells to the native write-mode backend.

    Args:
        ws: Worksheet whose pending write-mode state should be drained.
        writer: Native workbook writer exposed by the Rust extension.
        python_value_to_payload: Converter for per-cell value payloads.
        font_to_format_dict: Converter for font payload fragments.
        fill_to_format_dict: Converter for fill payload fragments.
        alignment_to_format_dict: Converter for alignment payload fragments.
        border_to_rust_dict: Converter for border payloads.
        rich_text_to_runs_payload: Converter for CellRichText runs.
    """
    _flush_append_buffer(ws, writer, python_value_to_payload)
    _flush_bulk_writes(ws, writer, python_value_to_payload)

    batch_values, individual_values, format_cells = _partition_dirty_cells(ws)
    _write_batch_values(ws, writer, batch_values)
    _write_individual_values(
        ws,
        writer,
        individual_values,
        python_value_to_payload,
        rich_text_to_runs_payload,
    )
    _flush_spill_child_placeholders(ws, writer)
    _flush_format_cells(
        ws,
        writer,
        format_cells,
        font_to_format_dict,
        fill_to_format_dict,
        alignment_to_format_dict,
        border_to_rust_dict,
    )


def _flush_append_buffer(
    ws: Worksheet,
    writer: Any,
    python_value_to_payload: Any,
) -> None:
    """Flush rows queued via ``append()`` using batch writes where possible."""
    if not ws._append_buffer:  # noqa: SLF001
        return

    buffer = ws._append_buffer  # noqa: SLF001
    start_row = ws._append_buffer_start  # noqa: SLF001
    start_a1 = rowcol_to_a1(start_row, 1)
    individual_values = ws._extract_non_batchable(buffer, start_row, 1)  # noqa: SLF001

    writer.write_sheet_values(ws._title, start_a1, buffer)  # noqa: SLF001

    for row, col, value in individual_values:
        coord = rowcol_to_a1(row, col)
        payload = python_value_to_payload(value)
        writer.write_cell_value(ws._title, coord, payload)  # noqa: SLF001

    ws._append_buffer = []  # noqa: SLF001


def _flush_bulk_writes(
    ws: Worksheet,
    writer: Any,
    python_value_to_payload: Any,
) -> None:
    """Flush rows queued via ``write_rows()`` using batch writes where possible."""
    for grid, start_row, start_col in ws._bulk_writes:  # noqa: SLF001
        start_a1 = rowcol_to_a1(start_row, start_col)
        individual_values = ws._extract_non_batchable(  # noqa: SLF001
            grid, start_row, start_col
        )
        writer.write_sheet_values(ws._title, start_a1, grid)  # noqa: SLF001
        for row, col, value in individual_values:
            coord = rowcol_to_a1(row, col)
            payload = python_value_to_payload(value)
            writer.write_cell_value(ws._title, coord, payload)  # noqa: SLF001

    ws._bulk_writes = []  # noqa: SLF001


def _partition_dirty_cells(
    ws: Worksheet,
) -> tuple[list[tuple[int, int, Any]], list[tuple[int, int, Any]], list[tuple[int, int, Any]]]:
    """Split dirty cells into batchable values, individual values, and formats."""
    batch_values: list[tuple[int, int, Any]] = []
    individual_values: list[tuple[int, int, Any]] = []
    format_cells: list[tuple[int, int, Any]] = []

    for row, col in ws._dirty:  # noqa: SLF001
        cell = ws._cells.get((row, col))  # noqa: SLF001
        if cell is None:
            continue

        if cell._value_dirty:  # noqa: SLF001
            value = cell._value  # noqa: SLF001
            if value is None or (
                isinstance(value, (int, float, str))
                and not isinstance(value, bool)
                and not (isinstance(value, str) and value.startswith("="))
            ):
                batch_values.append((row, col, value))
            else:
                individual_values.append((row, col, cell))

        if cell._format_dirty:  # noqa: SLF001
            format_cells.append((row, col, cell))

    return batch_values, individual_values, format_cells


def _write_batch_values(
    ws: Worksheet,
    writer: Any,
    batch_values: list[tuple[int, int, Any]],
) -> None:
    """Write simple dirty values as one rectangular batch when present."""
    if not batch_values:
        return

    min_row = batch_values[0][0]
    min_col = batch_values[0][1]
    max_row = min_row
    max_col = min_col
    for row, col, _value in batch_values:
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
    for row, col, value in batch_values:
        grid[row - min_row][col - min_col] = value

    start = rowcol_to_a1(min_row, min_col)
    writer.write_sheet_values(ws._title, start, grid)  # noqa: SLF001


def _write_individual_values(
    ws: Worksheet,
    writer: Any,
    individual_values: list[tuple[int, int, Any]],
    python_value_to_payload: Any,
    rich_text_to_runs_payload: Any,
) -> None:
    """Write non-batchable dirty values with type-preserving writer calls."""
    from wolfxl.cell.cell import ArrayFormula, DataTableFormula
    from wolfxl.cell.rich_text import CellRichText

    for _row, _col, cell in individual_values:
        coord = rowcol_to_a1(cell._row, cell._col)  # noqa: SLF001
        value = cell._value  # noqa: SLF001
        if isinstance(value, ArrayFormula):
            if hasattr(writer, "write_cell_array_formula"):
                writer.write_cell_array_formula(
                    ws._title,  # noqa: SLF001
                    coord,
                    {"kind": "array", "ref": value.ref, "text": value.text},
                )
            else:
                payload = python_value_to_payload(f"={value.text}")
                writer.write_cell_value(ws._title, coord, payload)  # noqa: SLF001
            continue
        if isinstance(value, DataTableFormula):
            if hasattr(writer, "write_cell_array_formula"):
                writer.write_cell_array_formula(
                    ws._title,  # noqa: SLF001
                    coord,
                    {
                        "kind": "data_table",
                        "ref": value.ref,
                        "ca": value.ca,
                        "dt2D": value.dt2D,
                        "dtr": value.dtr,
                        "r1": value.r1,
                        "r2": value.r2,
                    },
                )
            continue
        if isinstance(value, CellRichText):
            if hasattr(writer, "write_cell_rich_text"):
                runs_payload = rich_text_to_runs_payload(value)
                writer.write_cell_rich_text(ws._title, coord, runs_payload)  # noqa: SLF001
            else:
                payload = python_value_to_payload(str(value))
                writer.write_cell_value(ws._title, coord, payload)  # noqa: SLF001
            continue

        payload = python_value_to_payload(value)
        writer.write_cell_value(ws._title, coord, payload)  # noqa: SLF001


def _flush_spill_child_placeholders(ws: Worksheet, writer: Any) -> None:
    """Flush placeholder cells for array-formula spill ranges."""
    for (row, col), (kind, _payload) in ws._pending_array_formulas.items():  # noqa: SLF001
        if kind != "spill_child":
            continue
        if (row, col) in ws._dirty:  # noqa: SLF001
            continue
        coord = rowcol_to_a1(row, col)
        if hasattr(writer, "write_cell_array_formula"):
            writer.write_cell_array_formula(
                ws._title,  # noqa: SLF001
                coord,
                {"kind": "spill_child"},
            )


def _flush_format_cells(
    ws: Worksheet,
    writer: Any,
    format_cells: list[tuple[int, int, Any]],
    font_to_format_dict: Any,
    fill_to_format_dict: Any,
    alignment_to_format_dict: Any,
    border_to_rust_dict: Any,
) -> None:
    """Flush dirty cell format and border payloads."""
    from wolfxl._cell import _UNSET

    if not format_cells:
        return

    format_entries: list[tuple[int, int, dict[str, Any]]] = []
    border_entries: list[tuple[int, int, dict[str, Any]]] = []

    for _row, _col, cell in format_cells:
        row = cell._row  # noqa: SLF001
        col = cell._col  # noqa: SLF001
        fmt: dict[str, Any] = {}

        if cell._font is not _UNSET and cell._font is not None:  # noqa: SLF001
            fmt.update(font_to_format_dict(cell._font))  # noqa: SLF001
        if cell._fill is not _UNSET and cell._fill is not None:  # noqa: SLF001
            fmt.update(fill_to_format_dict(cell._fill))  # noqa: SLF001
        if cell._alignment is not _UNSET and cell._alignment is not None:  # noqa: SLF001
            fmt.update(alignment_to_format_dict(cell._alignment))  # noqa: SLF001
        if cell._number_format is not _UNSET and cell._number_format is not None:  # noqa: SLF001
            fmt["number_format"] = cell._number_format  # noqa: SLF001

        if fmt:
            format_entries.append((row, col, fmt))

        if cell._border is not _UNSET and cell._border is not None:  # noqa: SLF001
            border = border_to_rust_dict(cell._border)  # noqa: SLF001
            if border:
                border_entries.append((row, col, border))

    if len(format_entries) > 1:
        ws._batch_write_dicts(writer.write_sheet_formats, format_entries)  # noqa: SLF001
    else:
        for row, col, fmt in format_entries:
            writer.write_cell_format(ws._title, rowcol_to_a1(row, col), fmt)  # noqa: SLF001

    if len(border_entries) > 1:
        ws._batch_write_dicts(writer.write_sheet_borders, border_entries)  # noqa: SLF001
    else:
        for row, col, border in border_entries:
            writer.write_cell_border(ws._title, rowcol_to_a1(row, col), border)  # noqa: SLF001
