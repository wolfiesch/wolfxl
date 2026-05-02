"""Worksheet construction-state helpers."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from wolfxl._worksheet_collections import AutoFilter

if TYPE_CHECKING:
    from wolfxl._workbook import Workbook
    from wolfxl._worksheet import Worksheet


def initialize_worksheet_state(
    ws: Worksheet,
    workbook: Workbook,
    title: str,
) -> None:
    """Initialize worksheet caches, buffers, and pending mutation queues."""
    ws._workbook = workbook  # noqa: SLF001
    ws._title = title  # noqa: SLF001
    ws._cells: dict[tuple[int, int], Any] = {}  # noqa: SLF001
    ws._dirty: set[tuple[int, int]] = set()  # noqa: SLF001
    ws._dimensions: tuple[int, int] | None = None  # noqa: SLF001
    ws._max_col_idx = 0  # noqa: SLF001
    ws._next_append_row = 1  # noqa: SLF001

    ws._append_buffer: list[list[Any]] = []  # noqa: SLF001
    ws._append_buffer_start = 1  # noqa: SLF001
    ws._bulk_writes: list[tuple[list[list[Any]], int, int]] = []  # noqa: SLF001

    ws._freeze_panes: str | None = None  # noqa: SLF001
    ws._auto_filter = AutoFilter()  # noqa: SLF001
    ws._row_heights: dict[int, float | None] = {}  # noqa: SLF001
    ws._col_widths: dict[str, float | None] = {}  # noqa: SLF001
    ws._sheet_state: str | None = None  # noqa: SLF001
    ws._merged_ranges: set[str] = set()  # noqa: SLF001
    ws._print_area: str | None = None  # noqa: SLF001
    ws._sheet_visibility_cache: dict[str, Any] | None = None  # noqa: SLF001

    ws._comments_cache: dict[str, Any] | None = None  # noqa: SLF001
    ws._hyperlinks_cache: dict[str, Any] | None = None  # noqa: SLF001
    ws._defined_names_cache: dict[str, Any] | None = None  # noqa: SLF001
    ws._tables_cache: dict[str, Any] | None = None  # noqa: SLF001
    ws._data_validations_cache: Any | None = None  # noqa: SLF001
    ws._conditional_formatting_cache: Any | None = None  # noqa: SLF001
    ws._images_cache: list[Any] | None = None  # noqa: SLF001
    ws._charts_cache: list[Any] | None = None  # noqa: SLF001

    ws._pending_comments: dict[str, Any] = {}  # noqa: SLF001
    ws._pending_hyperlinks: dict[str, Any] = {}  # noqa: SLF001
    ws._pending_rich_text: dict[tuple[int, int], Any] = {}  # noqa: SLF001
    ws._pending_array_formulas: dict[  # noqa: SLF001
        tuple[int, int], tuple[str, dict[str, Any]]
    ] = {}
    ws._pending_tables: list[Any] = []  # noqa: SLF001
    ws._pending_data_validations: list[Any] = []  # noqa: SLF001
    ws._pending_conditional_formats: list[tuple[str, Any]] = []  # noqa: SLF001
    ws._pending_images: list[Any] = []  # noqa: SLF001
    ws._pending_charts: list[Any] = []  # noqa: SLF001
    ws._pending_pivot_tables: list[Any] = []  # noqa: SLF001
    ws._pending_slicers: list[Any] = []  # noqa: SLF001

    ws._page_setup: Any = None  # noqa: SLF001
    ws._page_margins: Any = None  # noqa: SLF001
    ws._print_options: Any = None  # noqa: SLF001
    ws._header_footer: Any = None  # noqa: SLF001
    ws._sheet_properties: Any = None  # noqa: SLF001
    ws._sheet_view: Any = None  # noqa: SLF001
    ws._protection: Any = None  # noqa: SLF001
    ws._print_title_rows: str | None = None  # noqa: SLF001
    ws._print_title_cols: str | None = None  # noqa: SLF001
    ws._row_breaks: Any = None  # noqa: SLF001
    ws._col_breaks: Any = None  # noqa: SLF001
    ws._sheet_format: Any = None  # noqa: SLF001
