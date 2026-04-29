"""Worksheet collection proxy objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


class AutoFilter:
    """Proxy for ``ws.auto_filter``.

    The object mirrors openpyxl's worksheet filter surface and keeps filter
    column plus sort-state serialization in one place.
    """

    __slots__ = ("_ref", "filter_columns", "sort_state")

    def __init__(self) -> None:
        from wolfxl.worksheet.filters import FilterColumn  # noqa: F401

        self._ref: str | None = None
        self.filter_columns: list[Any] = []
        self.sort_state: Any = None

    @property
    def ref(self) -> str | None:
        return self._ref

    @ref.setter
    def ref(self, value: str | None) -> None:
        self._ref = value

    def add_filter_column(
        self,
        col_id: int,
        filter: Any = None,
        *,
        hidden_button: bool = False,
        show_button: bool = True,
        date_group_items: Any = None,
    ) -> Any:
        """Append and return a ``FilterColumn`` entry."""
        from wolfxl.worksheet.filters import FilterColumn

        fc = FilterColumn(
            col_id=col_id,
            hidden_button=hidden_button,
            show_button=show_button,
            filter=filter,
            date_group_items=list(date_group_items) if date_group_items else [],
        )
        self.filter_columns.append(fc)
        return fc

    def add_sort_condition(
        self,
        ref: str,
        descending: bool = False,
        sort_by: str = "value",
        *,
        custom_list: str | None = None,
        dxf_id: int | None = None,
        icon_set: str | None = None,
        icon_id: int | None = None,
    ) -> Any:
        """Append and return a ``SortCondition`` entry."""
        from wolfxl.worksheet.filters import SortCondition, SortState

        if self.sort_state is None:
            self.sort_state = SortState(ref=ref)
        sc = SortCondition(
            ref=ref,
            descending=descending,
            sort_by=sort_by,
            custom_list=custom_list,
            dxf_id=dxf_id,
            icon_set=icon_set,
            icon_id=icon_id,
        )
        self.sort_state.sort_conditions.append(sc)
        return sc

    def to_rust_dict(self) -> dict[str, Any]:
        """Serialize the filter state for the native save pipeline."""
        from wolfxl.worksheet.filters import AutoFilter as _AF

        af = _AF(
            ref=self._ref,
            filter_columns=list(self.filter_columns),
            sort_state=self.sort_state,
        )
        return af.to_rust_dict()


class MergedCellsProxy:
    """Openpyxl-shaped proxy for ``Worksheet.merged_cells``."""

    __slots__ = ("_ws",)

    def __init__(self, ws: Worksheet) -> None:
        self._ws = ws

    @property
    def ranges(self) -> list[str]:
        ws = self._ws
        wb = ws._workbook  # noqa: SLF001
        if wb._rust_reader is None:  # noqa: SLF001
            return list(ws._merged_ranges)  # noqa: SLF001
        try:
            return wb._rust_reader.read_merged_ranges(ws._title)  # noqa: SLF001
        except Exception:
            return list(ws._merged_ranges)  # noqa: SLF001

    def __iter__(self):  # type: ignore[no-untyped-def]
        return iter(self.ranges)

    def __len__(self) -> int:
        return len(self.ranges)


def merge_cells(ws: Worksheet, range_string: str) -> None:
    """Merge a cell range through the write-mode Rust backend."""
    wb = ws._workbook  # noqa: SLF001
    if wb._rust_writer is None:  # noqa: SLF001
        raise RuntimeError("merge_cells requires write mode")
    wb._rust_writer.merge_cells(ws._title, range_string)  # noqa: SLF001
    ws._merged_ranges.add(range_string)  # noqa: SLF001


def unmerge_cells(ws: Worksheet, range_string: str) -> None:
    """Forget a merged range from the worksheet's pending merge set."""
    ws._merged_ranges.discard(range_string)  # noqa: SLF001
