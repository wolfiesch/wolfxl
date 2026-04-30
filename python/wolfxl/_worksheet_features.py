"""Lazy worksheet feature collection loaders."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def _strip_sheet_prefix(refers_to: str, sheet_name: str) -> str:
    if refers_to.startswith("="):
        refers_to = refers_to[1:]
    if "!" not in refers_to:
        return refers_to
    prefix, _, tail = refers_to.partition("!")
    if prefix.strip("'").replace("''", "'") == sheet_name:
        return tail
    return refers_to


def get_defined_names(ws: Worksheet) -> dict[str, Any]:
    """Return worksheet-scoped defined names for ``ws``."""
    if ws._defined_names_cache is not None:  # noqa: SLF001
        return ws._defined_names_cache  # noqa: SLF001
    from wolfxl.workbook import DefinedNameDict
    from wolfxl.workbook.defined_name import DefinedName

    names = DefinedNameDict()
    wb = ws._workbook  # noqa: SLF001
    if wb._rust_reader is not None:  # noqa: SLF001
        try:
            entries = wb._rust_reader.read_named_ranges(ws._title)  # noqa: SLF001
        except Exception:
            entries = []
        for entry in entries:
            if entry.get("scope") != "sheet":
                continue
            name = entry["name"]
            refers_to = _strip_sheet_prefix(entry["refers_to"], ws._title)
            dict.__setitem__(
                names,
                name,
                DefinedName(name=name, value=refers_to, localSheetId=None),
            )
    ws._defined_names_cache = names  # noqa: SLF001
    return names


def get_comments_map(ws: Worksheet) -> dict[str, Any]:
    """Return ``{cell_ref: Comment}`` for ``ws``, cached on the worksheet."""
    if ws._comments_cache is not None:  # noqa: SLF001
        return ws._comments_cache  # noqa: SLF001
    from wolfxl.comments import Comment

    wb = ws._workbook  # noqa: SLF001
    if wb._rust_reader is None:  # noqa: SLF001
        ws._comments_cache = {}  # noqa: SLF001
        return ws._comments_cache  # noqa: SLF001
    try:
        entries = wb._rust_reader.read_comments(ws._title)  # noqa: SLF001
    except Exception:
        entries = []
    result: dict[str, Any] = {}
    for entry in entries:
        cell_ref = entry.get("cell")
        if not cell_ref:
            continue
        comment = Comment(
            text=entry.get("text", ""),
            author=entry.get("author") or None,
        )
        comment.bind(ws[cell_ref])
        result[cell_ref] = comment
    ws._comments_cache = result  # noqa: SLF001
    return result


def get_hyperlinks_map(ws: Worksheet) -> dict[str, Any]:
    """Return ``{cell_ref: Hyperlink}`` for ``ws``, cached on the worksheet."""
    if ws._hyperlinks_cache is not None:  # noqa: SLF001
        return ws._hyperlinks_cache  # noqa: SLF001
    from wolfxl.worksheet.hyperlink import Hyperlink

    wb = ws._workbook  # noqa: SLF001
    if wb._rust_reader is None:  # noqa: SLF001
        ws._hyperlinks_cache = {}  # noqa: SLF001
        return ws._hyperlinks_cache  # noqa: SLF001
    try:
        entries = wb._rust_reader.read_hyperlinks(ws._title)  # noqa: SLF001
    except Exception:
        entries = []
    result: dict[str, Any] = {}
    for entry in entries:
        cell_ref = entry.get("cell")
        if not cell_ref:
            continue
        is_internal = bool(entry.get("internal", False))
        raw_target = entry.get("target")
        result[cell_ref] = Hyperlink(
            ref=cell_ref,
            target=None if is_internal else raw_target,
            location=raw_target if is_internal else None,
            display=entry.get("display") or None,
            tooltip=entry.get("tooltip") or None,
        )
    ws._hyperlinks_cache = result  # noqa: SLF001
    return result


def get_tables_map(ws: Worksheet) -> dict[str, Any]:
    """Return ``{table_name: Table}`` for ``ws``, cached on the worksheet."""
    if ws._tables_cache is not None:  # noqa: SLF001
        return ws._tables_cache  # noqa: SLF001
    from wolfxl.worksheet.table import Table, TableColumn, TableStyleInfo

    wb = ws._workbook  # noqa: SLF001
    if wb._rust_reader is None:  # noqa: SLF001
        ws._tables_cache = {}  # noqa: SLF001
        return ws._tables_cache  # noqa: SLF001
    try:
        entries = wb._rust_reader.read_tables(ws._title)  # noqa: SLF001
    except Exception:
        entries = []
    result: dict[str, Any] = {}
    for entry in entries:
        name = entry.get("name") or entry.get("displayName")
        if not name:
            continue
        style_name = entry.get("style") or entry.get("style_name")
        table_style_info = (
            TableStyleInfo(
                name=style_name,
                showRowStripes=bool(entry.get("show_row_stripes", False)),
                showColumnStripes=bool(entry.get("show_column_stripes", False)),
                showFirstColumn=bool(entry.get("show_first_column", False)),
                showLastColumn=bool(entry.get("show_last_column", False)),
            )
            if style_name is not None
            else None
        )
        columns_raw = entry.get("columns") or []
        table_columns = [
            TableColumn(id=index + 1, name=str(column))
            for index, column in enumerate(columns_raw)
        ]
        result[name] = Table(
            name=name,
            displayName=entry.get("displayName") or name,
            ref=entry.get("ref", ""),
            comment=entry.get("comment"),
            tableType=entry.get("table_type"),
            headerRowCount=1 if entry.get("header_row", True) else 0,
            totalsRowCount=1 if entry.get("totals_row", False) else 0,
            totalsRowShown=entry.get("totals_row_shown"),
            tableStyleInfo=table_style_info,
            tableColumns=table_columns,
        )
    ws._tables_cache = result  # noqa: SLF001
    return result


def get_data_validations(ws: Worksheet) -> Any:
    """Return the ``DataValidationList`` for ``ws``, cached on the worksheet."""
    if ws._data_validations_cache is not None:  # noqa: SLF001
        return ws._data_validations_cache  # noqa: SLF001
    from wolfxl.worksheet.datavalidation import DataValidation, DataValidationList

    wb = ws._workbook  # noqa: SLF001
    validation_list = DataValidationList(ws=ws)
    if wb._rust_reader is None:  # noqa: SLF001
        ws._data_validations_cache = validation_list  # noqa: SLF001
        return validation_list
    try:
        entries = wb._rust_reader.read_data_validations(ws._title)  # noqa: SLF001
    except Exception:
        entries = []
    for entry in entries:
        validation_list.dataValidation.append(
            DataValidation(
                type=entry.get("validation_type") or entry.get("type"),
                operator=entry.get("operator"),
                formula1=entry.get("formula1"),
                formula2=entry.get("formula2"),
                allowBlank=bool(entry.get("allow_blank", False)),
                showErrorMessage=bool(entry.get("show_error_message", False)),
                showInputMessage=bool(entry.get("show_input_message", False)),
                error=entry.get("error"),
                errorTitle=entry.get("error_title"),
                prompt=entry.get("prompt"),
                promptTitle=entry.get("prompt_title"),
                sqref=entry.get("range") or entry.get("sqref") or "",
            )
        )
    ws._data_validations_cache = validation_list  # noqa: SLF001
    return validation_list


def get_conditional_formatting(ws: Worksheet) -> Any:
    """Return the ``ConditionalFormattingList`` for ``ws``, cached on the worksheet."""
    if ws._conditional_formatting_cache is not None:  # noqa: SLF001
        return ws._conditional_formatting_cache  # noqa: SLF001
    from wolfxl.formatting import ConditionalFormatting, ConditionalFormattingList
    from wolfxl.formatting.rule import Rule

    wb = ws._workbook  # noqa: SLF001
    formatting_list = ConditionalFormattingList(ws=ws)
    if wb._rust_reader is None:  # noqa: SLF001
        ws._conditional_formatting_cache = formatting_list  # noqa: SLF001
        return formatting_list
    try:
        entries = wb._rust_reader.read_conditional_formats(ws._title)  # noqa: SLF001
    except Exception:
        entries = []
    grouped: dict[str, list[Rule]] = {}
    order: list[str] = []
    for entry in entries:
        sqref = entry.get("range") or entry.get("sqref") or ""
        if sqref not in grouped:
            grouped[sqref] = []
            order.append(sqref)
        formula = entry.get("formula")
        if formula is None:
            formula_list: list[str] = []
        elif isinstance(formula, list):
            formula_list = [str(item) for item in formula]
        else:
            formula_list = [str(formula)]
        grouped[sqref].append(
            Rule(
                type=entry.get("rule_type") or entry.get("type") or "expression",
                operator=entry.get("operator"),
                formula=formula_list,
                stopIfTrue=bool(entry.get("stop_if_true", False)),
                priority=int(entry.get("priority", 1)),
            )
        )
    for sqref in order:
        formatting_list._append_entry(  # noqa: SLF001
            ConditionalFormatting(sqref=sqref, rules=grouped[sqref])
        )
    ws._conditional_formatting_cache = formatting_list  # noqa: SLF001
    return formatting_list


def add_table(ws: Worksheet, table: Any) -> None:
    """Attach a table to ``ws`` and queue it for save-time flushing."""
    from wolfxl.worksheet.table import Table

    if not isinstance(table, Table):
        raise TypeError(
            f"add_table() expects a wolfxl.worksheet.table.Table, got {type(table).__name__}"
        )
    if ws._tables_cache is None:  # noqa: SLF001
        ws._tables_cache = {}  # noqa: SLF001
    ws._tables_cache[table.name] = table  # noqa: SLF001
    ws._pending_tables.append(table)  # noqa: SLF001


def add_data_validation(ws: Worksheet, validation: Any) -> None:
    """Openpyxl-style alias for ``ws.data_validations.append(validation)``."""
    get_data_validations(ws).append(validation)
