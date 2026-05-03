"""wolfxl <-> openpyxl compatibility spec - source of truth.

Editing protocol: change ``ENTRIES`` and ``CATEGORIES`` here, then run
``python scripts/render_compat_matrix.py`` to regenerate
``docs/migration/compatibility-matrix.md``. The
``tests/test_openpyxl_compat_oracle.py`` harness imports
``ENTRIES`` directly to drive its parametrised probes; entries with a
``probe`` field run live against wolfxl + openpyxl in CI.

Status values:

* ``supported``    ``+``  implemented; covered by tests / fixtures
* ``partial``      ``~``  works for the common case, documented caveats
* ``not_yet``      ``-``  not implemented; tracked under a gap (gap_id)
* ``out_of_scope`` ``x``  explicitly out of roadmap

``gap_id`` matches the G01-G28 rows in
``Plans/openpyxl-parity-program.md``.
"""

from __future__ import annotations

from typing import TypedDict


class Category(TypedDict):
    id: str
    title: str


class _EntryRequired(TypedDict):
    id: str
    category: str
    openpyxl: str
    wolfxl: str
    status: str  # supported | partial | not_yet | out_of_scope


class Entry(_EntryRequired, total=False):
    gap_id: str  # G01..G28
    probe: str  # probe id wired into the oracle harness
    notes: str


STATUS_DISPLAY: dict[str, str] = {
    "supported": "[+] Supported",
    "partial": "[~] Partial",
    "not_yet": "[-] Not Yet",
    "out_of_scope": "[x] Out of Scope",
}


CATEGORIES: list[Category] = [
    {"id": "workbook", "title": "Workbook + Worksheet"},
    {"id": "cell_styles", "title": "Cell + style API"},
    {"id": "charts", "title": "Charts"},
    {"id": "pivots", "title": "Pivot tables"},
    {"id": "images", "title": "Images + drawings"},
    {"id": "structural", "title": "Worksheet structural ops"},
    {"id": "modify", "title": "Modify-mode mutations"},
    {"id": "read", "title": "Read-side parity"},
    {"id": "utils", "title": "Utility functions"},
    {"id": "streaming", "title": "Streaming write (`write_only=True`)"},
    {"id": "protection", "title": "Protection"},
    {"id": "external_links", "title": "External links"},
    {"id": "vba", "title": "VBA macros"},
    {"id": "legacy_formats", "title": "Legacy formats (`.xlsb` / `.xls` / `.ods`)"},
    {"id": "comments", "title": "Comments"},
    {"id": "rich_text", "title": "Rich text"},
    {"id": "cf", "title": "Conditional formatting"},
    {"id": "defined_names", "title": "Defined names"},
    {"id": "print_settings", "title": "Print settings + page setup"},
    {"id": "array_formulas", "title": "Array + data-table formulas"},
    {"id": "calc_chain", "title": "Calc chain"},
    {"id": "slicers", "title": "Slicers"},
]


ENTRIES: list[Entry] = [
    # --- Workbook + Worksheet ----------------------------------------------
    {
        "id": "workbook.open.basic",
        "category": "workbook",
        "openpyxl": "Workbook()",
        "wolfxl": "wolfxl.Workbook()",
        "status": "supported",
        "probe": "workbook_open_basic",
        "notes": "Write-mode default.",
    },
    {
        "id": "workbook.load.basic",
        "category": "workbook",
        "openpyxl": "load_workbook(path)",
        "wolfxl": "wolfxl.load_workbook(path)",
        "status": "supported",
        "probe": "workbook_load_basic",
    },
    {
        "id": "workbook.load.data_only",
        "category": "workbook",
        "openpyxl": "load_workbook(path, data_only=True)",
        "wolfxl": "wolfxl.load_workbook(path, data_only=True)",
        "status": "supported",
        "probe": "workbook_load_data_only",
    },
    {
        "id": "workbook.load.modify",
        "category": "workbook",
        "openpyxl": "Implicit (full DOM rewrite)",
        "wolfxl": "wolfxl.load_workbook(path, modify=True)",
        "status": "supported",
        "probe": "workbook_load_modify",
        "notes": "Surgical patcher; faster than DOM rewrite.",
    },
    {
        "id": "workbook.load.read_only",
        "category": "workbook",
        "openpyxl": "load_workbook(path, read_only=True)",
        "wolfxl": "wolfxl.load_workbook(path, read_only=True)",
        "status": "supported",
        "probe": "workbook_load_read_only",
        "notes": "Streaming reads (auto-engages > 50k rows).",
    },
    {
        "id": "workbook.save.basic",
        "category": "workbook",
        "openpyxl": "wb.save(path)",
        "wolfxl": "wb.save(path)",
        "status": "supported",
        "probe": "workbook_save_basic",
    },
    {
        "id": "workbook.sheet_access",
        "category": "workbook",
        "openpyxl": 'wb["Sheet"], wb.active, wb.sheetnames',
        "wolfxl": 'wb["Sheet"], wb.active, wb.sheetnames',
        "status": "supported",
        "probe": "workbook_sheet_access",
    },
    {
        "id": "workbook.create_sheet",
        "category": "workbook",
        "openpyxl": "wb.create_sheet(title)",
        "wolfxl": "wb.create_sheet(title)",
        "status": "supported",
        "probe": "workbook_create_sheet",
    },
    {
        "id": "workbook.copy_worksheet",
        "category": "workbook",
        "openpyxl": "wb.copy_worksheet(ws)",
        "wolfxl": "wb.copy_worksheet(ws)",
        "status": "supported",
        "probe": "workbook_copy_worksheet",
        "notes": "Diverges from openpyxl in 5 documented ways (always more preservation).",
    },
    {
        "id": "workbook.write_only",
        "category": "streaming",
        "openpyxl": "Workbook(write_only=True)",
        "wolfxl": "wolfxl.Workbook(write_only=True)",
        "status": "partial",
        "gap_id": "G20",
        "probe": "workbook_write_only_streaming",
        "notes": "Accepted today as a kwarg but routes through the standard in-memory writer. S7 introduces a true bounded-memory append-only path.",
    },
    # --- Cell + style API --------------------------------------------------
    {
        "id": "cell.basic_value",
        "category": "cell_styles",
        "openpyxl": 'ws["A1"].value, ws.cell(r, c).value',
        "wolfxl": 'ws["A1"].value, ws.cell(r, c).value',
        "status": "supported",
        "probe": "cell_basic_value",
    },
    {
        "id": "cell.font_fill_border_alignment",
        "category": "cell_styles",
        "openpyxl": "cell.font / fill / border / alignment / number_format",
        "wolfxl": "cell.font / fill / border / alignment / number_format",
        "status": "supported",
        "probe": "cell_font_fill_border_alignment",
        "notes": "All five attributes (plus protection and border) round-trip together on one cell. Python flush layer merges format and border keys into a single dict so the native writer interns one combined xf record, instead of minting two ids that overwrite each other.",
    },
    {
        "id": "cell.diagonal_borders",
        "category": "cell_styles",
        "openpyxl": "Border(diagonal=Side(...), diagonalUp=True)",
        "wolfxl": "Border(diagonal=Side(...), diagonalUp=True)",
        "status": "supported",
        "probe": "cell_diagonal_borders",
        "notes": "Round-trips through both write mode and modify mode; diagonalUp/diagonalDown attrs and the <diagonal> child of <border> persist via the patcher's BorderSpec and the writer's intern path.",
    },
    {
        "id": "cell.protection",
        "category": "cell_styles",
        "openpyxl": "cell.protection = Protection(...)",
        "wolfxl": "cell.protection = Protection(...)",
        "status": "supported",
        "probe": "cell_protection",
        "notes": "Round-trips through wolfxl reload; locked/hidden flags persist via <protection> child of <xf> with applyProtection=\"1\".",
    },
    {
        "id": "cell.named_style",
        "category": "cell_styles",
        "openpyxl": "NamedStyle(name=...) + wb.add_named_style(...)",
        "wolfxl": "NamedStyle(name=...) + wb.add_named_style(...)",
        "status": "supported",
        "probe": "cell_named_style",
        "notes": "cell.style binds a registered NamedStyle and round-trips through cellStyleXfs/cellStyles + the xfId attr on <xf>; reader resurfaces the name via cell.style.",
    },
    {
        "id": "cell.gradient_fill",
        "category": "cell_styles",
        "openpyxl": "GradientFill(stop=(...))",
        "wolfxl": "GradientFill(stop=(...))",
        "status": "partial",
        "gap_id": "G05",
        "probe": "cell_gradient_fill",
    },
    # --- Charts ------------------------------------------------------------
    {
        "id": "charts.basic_2d",
        "category": "charts",
        "openpyxl": "BarChart / LineChart / PieChart / DoughnutChart",
        "wolfxl": "BarChart / LineChart / PieChart / DoughnutChart",
        "status": "supported",
        "probe": "charts_basic_2d",
    },
    {
        "id": "charts.advanced_2d",
        "category": "charts",
        "openpyxl": "AreaChart / ScatterChart / BubbleChart / RadarChart",
        "wolfxl": "AreaChart / ScatterChart / BubbleChart / RadarChart",
        "status": "supported",
        "probe": "charts_advanced_2d",
    },
    {
        "id": "charts.3d",
        "category": "charts",
        "openpyxl": "BarChart3D / LineChart3D / PieChart3D / AreaChart3D",
        "wolfxl": "BarChart3D / LineChart3D / PieChart3D / AreaChart3D",
        "status": "supported",
    },
    {
        "id": "charts.surface_stock_projected",
        "category": "charts",
        "openpyxl": "SurfaceChart / SurfaceChart3D / StockChart / ProjectedPieChart",
        "wolfxl": "SurfaceChart / SurfaceChart3D / StockChart / ProjectedPieChart",
        "status": "supported",
    },
    {
        "id": "charts.add_remove_replace",
        "category": "charts",
        "openpyxl": "ws.add_chart / remove_chart / replace_chart",
        "wolfxl": "ws.add_chart / remove_chart / replace_chart",
        "status": "supported",
        "probe": "charts_add_remove_replace",
        "notes": "remove/replace shipped in v1.7.",
    },
    {
        "id": "charts.combination",
        "category": "charts",
        "openpyxl": "bar + line on shared category axis with secondary value axis",
        "wolfxl": "bar + line on shared category axis with secondary value axis",
        "status": "not_yet",
        "gap_id": "G15",
        "probe": "charts_combination",
    },
    {
        "id": "charts.label_rich_text",
        "category": "charts",
        "openpyxl": "data label + axis label rich-text runs",
        "wolfxl": "data label + axis label rich-text runs",
        "status": "partial",
        "gap_id": "G10",
        "probe": "charts_label_rich_text",
        "notes": "Title rich text shipped in v1.7; labels and axis-label runs are next.",
    },
    {
        "id": "charts.pivot_chart_per_point",
        "category": "charts",
        "openpyxl": "pivot chart with per-point overrides",
        "wolfxl": "pivot chart with per-point overrides",
        "status": "partial",
        "gap_id": "G16",
    },
    # --- Pivot tables ------------------------------------------------------
    {
        "id": "pivots.construction",
        "category": "pivots",
        "openpyxl": "PivotCache + PivotTable construction",
        "wolfxl": "wolfxl.pivot.PivotCache + PivotTable",
        "status": "supported",
        "probe": "pivots_construction",
        "notes": "v2.0 ships pre-aggregated records emit (no refresh-on-open required).",
    },
    {
        "id": "pivots.linked_chart",
        "category": "pivots",
        "openpyxl": "chart.pivot_source = pt",
        "wolfxl": "chart.pivot_source = pt",
        "status": "supported",
    },
    {
        "id": "pivots.in_place_edit",
        "category": "pivots",
        "openpyxl": "edit existing pivot's source range, field order, etc.",
        "wolfxl": "edit existing pivot's source range, field order, etc.",
        "status": "not_yet",
        "gap_id": "G17",
        "probe": "pivots_in_place_edit",
    },
    {
        "id": "pivots.copy_worksheet",
        "category": "pivots",
        "openpyxl": "copy_worksheet of pivot-bearing sheet (drops in openpyxl)",
        "wolfxl": "copy_worksheet of pivot-bearing sheet (deep-clone)",
        "status": "supported",
    },
    # --- Images + drawings -------------------------------------------------
    {
        "id": "images.basic",
        "category": "images",
        "openpyxl": 'Image("logo.png") + ws.add_image(img, "B5")',
        "wolfxl": 'Image("logo.png") + ws.add_image(img, "B5")',
        "status": "supported",
        "probe": "images_basic",
    },
    {
        "id": "images_replace_remove",
        "category": "images",
        "openpyxl": "ws.replace_image / remove_image",
        "wolfxl": "ws.replace_image / remove_image",
        "status": "supported",
        "probe": "images_replace_remove",
    },
    # --- Structural ops ----------------------------------------------------
    {
        "id": "structural.insert_delete_rows",
        "category": "structural",
        "openpyxl": "ws.insert_rows / delete_rows",
        "wolfxl": "ws.insert_rows / delete_rows",
        "status": "supported",
        "probe": "structural_insert_delete_rows",
    },
    {
        "id": "structural.insert_delete_cols",
        "category": "structural",
        "openpyxl": "ws.insert_cols / delete_cols",
        "wolfxl": "ws.insert_cols / delete_cols",
        "status": "supported",
    },
    {
        "id": "structural.move_range",
        "category": "structural",
        "openpyxl": "ws.move_range(range, rows, cols)",
        "wolfxl": "ws.move_range(range, rows, cols)",
        "status": "supported",
    },
    # --- Modify-mode mutations ---------------------------------------------
    {
        "id": "modify.document_properties",
        "category": "modify",
        "openpyxl": "wb.properties.title = ...",
        "wolfxl": "wb.properties.title = ...",
        "status": "supported",
    },
    {
        "id": "modify.defined_names",
        "category": "modify",
        "openpyxl": "wb.defined_names[name] = DefinedName(...)",
        "wolfxl": "wb.defined_names[name] = DefinedName(...)",
        "status": "supported",
        "probe": "modify_defined_names",
    },
    {
        "id": "modify.tables",
        "category": "modify",
        "openpyxl": "ws.add_table(Table(...))",
        "wolfxl": "ws.add_table(Table(...))",
        "status": "supported",
        "probe": "modify_tables",
    },
    {
        "id": "modify.data_validations",
        "category": "modify",
        "openpyxl": "ws.data_validations.append(...)",
        "wolfxl": "ws.data_validations.append(...)",
        "status": "supported",
        "probe": "modify_data_validations",
    },
    # --- Read-side ---------------------------------------------------------
    {
        "id": "read.xlsx",
        "category": "read",
        "openpyxl": "load_workbook(path)",
        "wolfxl": "load_workbook(path)",
        "status": "supported",
    },
    {
        "id": "read.xlsb",
        "category": "read",
        "openpyxl": "not supported in openpyxl",
        "wolfxl": "native BIFF12; values, cached formulas, read-side styles",
        "status": "supported",
    },
    {
        "id": "read.xls",
        "category": "read",
        "openpyxl": "not supported in openpyxl",
        "wolfxl": "calamine; value-only, styles raise",
        "status": "partial",
        "notes": "Style accessors raise; values + formulas read.",
    },
    {
        "id": "read.ods",
        "category": "read",
        "openpyxl": "not supported in openpyxl",
        "wolfxl": "not supported (out of scope)",
        "status": "out_of_scope",
        "gap_id": "G27",
    },
    # --- Utility functions -------------------------------------------------
    {
        "id": "utils.get_column_letter",
        "category": "utils",
        "openpyxl": "openpyxl.utils.get_column_letter",
        "wolfxl": "wolfxl.utils.cell.get_column_letter",
        "status": "supported",
        "probe": "utils_get_column_letter",
    },
    {
        "id": "utils.column_index_from_string",
        "category": "utils",
        "openpyxl": "openpyxl.utils.column_index_from_string",
        "wolfxl": "wolfxl.utils.cell.column_index_from_string",
        "status": "supported",
        "probe": "utils_column_index_from_string",
    },
    {
        "id": "utils.range_boundaries",
        "category": "utils",
        "openpyxl": "openpyxl.utils.range_boundaries",
        "wolfxl": "wolfxl.utils.cell.range_boundaries",
        "status": "supported",
        "probe": "utils_range_boundaries",
    },
    {
        "id": "utils.coordinate_to_tuple",
        "category": "utils",
        "openpyxl": "openpyxl.utils.coordinate_to_tuple",
        "wolfxl": "wolfxl.utils.cell.coordinate_to_tuple",
        "status": "supported",
    },
    # --- Protection -------------------------------------------------------
    {
        "id": "protection.sheet",
        "category": "protection",
        "openpyxl": "ws.protection = SheetProtection(...)",
        "wolfxl": "ws.protection = SheetProtection(...)",
        "status": "supported",
        "probe": "protection_sheet",
        "notes": "Round-trips through openpyxl reload (sheet flag, formatCells override, password hash).",
    },
    {
        "id": "protection.workbook",
        "category": "protection",
        "openpyxl": "wb.security = WorkbookProtection(...)",
        "wolfxl": "wb.security = WorkbookProtection(...)",
        "status": "supported",
        "probe": "protection_workbook",
        "notes": "camelCase aliases (lockStructure, workbookPassword, etc.) accepted alongside snake_case; round-trips through openpyxl reload.",
    },
    # --- External links ---------------------------------------------------
    {
        "id": "external_links.workbook_collection",
        "category": "external_links",
        "openpyxl": "wb._external_links + xl/externalLinks/* parts",
        "wolfxl": "wb._external_links + xl/externalLinks/* parts",
        "status": "not_yet",
        "gap_id": "G18",
        "probe": "external_links_collection",
    },
    # --- VBA --------------------------------------------------------------
    {
        "id": "vba.preserve",
        "category": "vba",
        "openpyxl": ".xlsm preserved on read+write",
        "wolfxl": ".xlsm preserved on modify-mode save",
        "status": "supported",
    },
    {
        "id": "vba.inspect",
        "category": "vba",
        "openpyxl": "workbook.vba_archive (read-only inspection)",
        "wolfxl": "workbook.vba_archive (read-only inspection)",
        "status": "not_yet",
        "gap_id": "G19",
    },
    {
        "id": "vba.author",
        "category": "vba",
        "openpyxl": "not supported in openpyxl",
        "wolfxl": "macro authoring from Python",
        "status": "not_yet",
        "gap_id": "G28",
        "notes": "Decision-gated (S11).",
    },
    # --- Legacy formats ---------------------------------------------------
    {
        "id": "legacy_formats.xlsb_write",
        "category": "legacy_formats",
        "openpyxl": "not supported",
        "wolfxl": "write `.xlsb`",
        "status": "not_yet",
        "gap_id": "G25",
        "notes": "Decision-gated (S9).",
    },
    {
        "id": "legacy_formats.xls_write",
        "category": "legacy_formats",
        "openpyxl": "not supported",
        "wolfxl": "write `.xls`",
        "status": "not_yet",
        "gap_id": "G26",
        "notes": "Decision-gated (S9).",
    },
    {
        "id": "legacy_formats.ods_read_write",
        "category": "legacy_formats",
        "openpyxl": "not supported",
        "wolfxl": "read+write `.ods`",
        "status": "not_yet",
        "gap_id": "G27",
        "notes": "Decision-gated (S10).",
    },
    # --- Comments ---------------------------------------------------------
    {
        "id": "comments.basic",
        "category": "comments",
        "openpyxl": "cell.comment = Comment(text, author)",
        "wolfxl": "cell.comment = Comment(text, author)",
        "status": "supported",
        "probe": "comments_basic",
    },
    {
        "id": "comments.threaded",
        "category": "comments",
        "openpyxl": "threaded comment write + modify (xl/threadedComments)",
        "wolfxl": "threaded comment write + modify (xl/threadedComments)",
        "status": "not_yet",
        "gap_id": "G08",
        "probe": "comments_threaded",
    },
    # --- Rich text --------------------------------------------------------
    {
        "id": "rich_text.cell",
        "category": "rich_text",
        "openpyxl": "cell.value = CellRichText(...)",
        "wolfxl": "cell.value = CellRichText(...)",
        "status": "supported",
        "probe": "rich_text_cell",
    },
    {
        "id": "rich_text.headers_footers",
        "category": "rich_text",
        "openpyxl": "ws.oddHeader / oddFooter rich-text runs",
        "wolfxl": "ws.oddHeader / oddFooter rich-text runs",
        "status": "not_yet",
        "gap_id": "G09",
        "probe": "rich_text_headers_footers",
    },
    # --- Conditional formatting -------------------------------------------
    {
        "id": "cf.basic_rules",
        "category": "cf",
        "openpyxl": "cellIs / containsText / expression / colorScale (basic)",
        "wolfxl": "cellIs / containsText / expression / colorScale (basic)",
        "status": "supported",
        "probe": "cf_basic_rules",
    },
    {
        "id": "cf.icon_sets",
        "category": "cf",
        "openpyxl": "IconSetRule (3 / 4 / 5 icons + percentile / number ladders)",
        "wolfxl": "IconSetRule (3 / 4 / 5 icons + percentile / number ladders)",
        "status": "not_yet",
        "gap_id": "G11",
        "probe": "cf_icon_sets",
    },
    {
        "id": "cf.data_bars",
        "category": "cf",
        "openpyxl": "DataBarRule (gradient + solid; min / max / percent / formula)",
        "wolfxl": "DataBarRule (gradient + solid; min / max / percent / formula)",
        "status": "supported",
        "probe": "cf_data_bars",
        "notes": "Round-trips through openpyxl reload (cfvo min/max types preserved). Edge cases like percent / formula cfvo not yet probed.",
    },
    {
        "id": "cf.color_scales_advanced",
        "category": "cf",
        "openpyxl": "ColorScaleRule with 3-stop, percentile / formula / number cfvo",
        "wolfxl": "ColorScaleRule with 3-stop, percentile / formula / number cfvo",
        "status": "partial",
        "gap_id": "G13",
        "probe": "cf_color_scales_advanced",
    },
    {
        "id": "cf.stop_if_true_priority",
        "category": "cf",
        "openpyxl": "stopIfTrue + explicit priority + dxf integration",
        "wolfxl": "stopIfTrue + explicit priority + dxf integration",
        "status": "partial",
        "gap_id": "G14",
        "probe": "cf_stop_if_true_priority",
    },
    # --- Defined names ----------------------------------------------------
    {
        "id": "defined_names.basic",
        "category": "defined_names",
        "openpyxl": "wb.defined_names + DefinedName(name=, value=)",
        "wolfxl": "wb.defined_names + DefinedName(name=, value=)",
        "status": "supported",
        "probe": "defined_names_basic",
    },
    {
        "id": "defined_names.edge_cases",
        "category": "defined_names",
        "openpyxl": "hidden, comment, custom_menu, function, function_group_id, shortcut_key",
        "wolfxl": "hidden, comment, custom_menu, function, function_group_id, shortcut_key",
        "status": "partial",
        "gap_id": "G22",
        "probe": "defined_names_edge_cases",
    },
    # --- Print settings ---------------------------------------------------
    {
        "id": "print_settings.basic",
        "category": "print_settings",
        "openpyxl": "ws.print_options / page_setup / page_margins / print_title_rows",
        "wolfxl": "ws.print_options / page_setup / page_margins / print_title_rows",
        "status": "supported",
        "probe": "print_settings_basic",
    },
    {
        "id": "print_settings.depth",
        "category": "print_settings",
        "openpyxl": "full PageSetup / PrintOptions surface (~30 attrs)",
        "wolfxl": "full PageSetup / PrintOptions surface (~30 attrs)",
        "status": "partial",
        "gap_id": "G24",
        "probe": "print_settings_depth",
    },
    # --- Array + data-table formulas --------------------------------------
    {
        "id": "array_formulas.array_formula",
        "category": "array_formulas",
        "openpyxl": "ArrayFormula(ref, text)",
        "wolfxl": "ArrayFormula(ref, text)",
        "status": "supported",
        "probe": "array_formula_basic",
        "notes": "Round-trips through openpyxl reload (ref + text preserved as openpyxl ArrayFormula).",
    },
    {
        "id": "array_formulas.data_table",
        "category": "array_formulas",
        "openpyxl": "DataTableFormula(...)",
        "wolfxl": "DataTableFormula(...)",
        "status": "supported",
        "probe": "array_formula_data_table",
    },
    {
        "id": "array_formulas.spill",
        "category": "array_formulas",
        "openpyxl": "dynamic-array spill metadata",
        "wolfxl": "dynamic-array spill metadata",
        "status": "partial",
        "gap_id": "G07",
    },
    # --- Calc chain -------------------------------------------------------
    {
        "id": "calc_chain.basic",
        "category": "calc_chain",
        "openpyxl": "calcChain rebuild on modify",
        "wolfxl": "calcChain rebuild on modify",
        "status": "supported",
    },
    {
        "id": "calc_chain.edge_cases",
        "category": "calc_chain",
        "openpyxl": "cross-sheet calc-chain ordering, deleted-cell pruning",
        "wolfxl": "cross-sheet calc-chain ordering, deleted-cell pruning",
        "status": "partial",
        "gap_id": "G23",
    },
    # --- Slicers ----------------------------------------------------------
    {
        "id": "slicers.with_pivot",
        "category": "slicers",
        "openpyxl": "Slicer + SlicerCache wired to a PivotCache",
        "wolfxl": "Slicer + SlicerCache wired to a PivotCache",
        "status": "supported",
    },
    {
        "id": "slicers.standalone",
        "category": "slicers",
        "openpyxl": "Slicer outside pivot context (table-driven, etc.)",
        "wolfxl": "Slicer outside pivot context (table-driven, etc.)",
        "status": "not_yet",
        "gap_id": "G21",
    },
]


def entries_by_category() -> dict[str, list[Entry]]:
    """Return entries grouped by category id, preserving definition order."""
    grouped: dict[str, list[Entry]] = {c["id"]: [] for c in CATEGORIES}
    for entry in ENTRIES:
        grouped.setdefault(entry["category"], []).append(entry)
    return grouped


def status_totals() -> dict[str, int]:
    """Return per-status totals across all entries."""
    totals: dict[str, int] = {}
    for entry in ENTRIES:
        totals[entry["status"]] = totals.get(entry["status"], 0) + 1
    return totals


def probes() -> list[Entry]:
    """Return entries that carry a `probe` field, in declaration order."""
    return [e for e in ENTRIES if e.get("probe")]
