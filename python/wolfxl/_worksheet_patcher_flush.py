"""Modify-mode worksheet cell flush helpers for the Rust patcher."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from wolfxl._utils import rowcol_to_a1

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def flush_to_patcher(
    ws: Worksheet,
    patcher: Any,
    python_value_to_payload: Any,
    font_to_format_dict: Any,
    fill_to_format_dict: Any,
    alignment_to_format_dict: Any,
    border_to_rust_dict: Any,
    rich_text_to_runs_payload: Any,
    protection_to_format_dict: Any,
) -> None:
    """Flush dirty worksheet cells to the ``XlsxPatcher`` backend."""
    from wolfxl._cell import _UNSET
    from wolfxl.cell.cell import ArrayFormula, DataTableFormula
    from wolfxl.cell.rich_text import CellRichText

    spill_children: set[tuple[int, int]] = {
        key
        for key, (kind, _payload) in ws._pending_array_formulas.items()  # noqa: SLF001
        if kind == "spill_child" and key not in ws._dirty  # noqa: SLF001
    }

    for row, col in ws._dirty:  # noqa: SLF001
        cell = ws._cells.get((row, col))  # noqa: SLF001
        if cell is None:
            continue
        coord = rowcol_to_a1(row, col)

        if cell._value_dirty:  # noqa: SLF001
            value = cell._value  # noqa: SLF001
            if isinstance(value, ArrayFormula):
                patcher.queue_array_formula(
                    ws._title,  # noqa: SLF001
                    coord,
                    {"kind": "array", "ref": value.ref, "text": value.text},
                )
            elif isinstance(value, DataTableFormula):
                patcher.queue_array_formula(
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
                        "del1": value.del1,
                        "del2": value.del2,
                    },
                )
            elif isinstance(value, CellRichText):
                runs_payload = rich_text_to_runs_payload(value)
                patcher.queue_rich_text_value(ws._title, coord, runs_payload)  # noqa: SLF001
            else:
                payload = python_value_to_payload(value)
                patcher.queue_value(ws._title, coord, payload)  # noqa: SLF001

        if cell._format_dirty:  # noqa: SLF001
            fmt: dict[str, Any] = {}

            if cell._font is not _UNSET and cell._font is not None:  # noqa: SLF001
                fmt.update(font_to_format_dict(cell._font))  # noqa: SLF001
            if cell._fill is not _UNSET and cell._fill is not None:  # noqa: SLF001
                fmt.update(fill_to_format_dict(cell._fill))  # noqa: SLF001
            if cell._alignment is not _UNSET and cell._alignment is not None:  # noqa: SLF001
                fmt.update(alignment_to_format_dict(cell._alignment))  # noqa: SLF001
            if cell._number_format is not _UNSET and cell._number_format is not None:  # noqa: SLF001
                fmt["number_format"] = cell._number_format  # noqa: SLF001
            if cell._protection is not _UNSET and cell._protection is not None:  # noqa: SLF001
                fmt.update(protection_to_format_dict(cell._protection))  # noqa: SLF001

            if fmt:
                patcher.queue_format(ws._title, coord, fmt)  # noqa: SLF001

            if cell._border is not _UNSET and cell._border is not None:  # noqa: SLF001
                border = border_to_rust_dict(cell._border)  # noqa: SLF001
                if border:
                    patcher.queue_border(ws._title, coord, border)  # noqa: SLF001

    for row, col in spill_children:
        coord = rowcol_to_a1(row, col)
        patcher.queue_array_formula(
            ws._title,  # noqa: SLF001
            coord,
            {"kind": "spill_child"},
        )
