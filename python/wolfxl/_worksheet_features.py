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


def get_threaded_comments_map(ws: Worksheet) -> dict[str, Any]:
    """Return ``{cell_ref: ThreadedComment}`` for ``ws``, cached on the worksheet.

    Reassembles the flat OOXML payload into a tree: top-level threads
    keyed by ``cell_ref``, each with their replies attached via
    ``ThreadedComment.replies``. Reply-to-parent links use the GUID
    chain from ``parentId``. Persons are resolved through
    ``wb.persons.by_id`` so the same ``Person`` instance appears across
    threads (matching openpyxl).
    """
    if ws._threaded_comments_cache is not None:  # noqa: SLF001
        return ws._threaded_comments_cache  # noqa: SLF001
    from datetime import datetime

    from wolfxl.comments import ThreadedComment

    wb = ws._workbook  # noqa: SLF001
    if wb._rust_reader is None or not hasattr(  # noqa: SLF001
        wb._rust_reader, "read_threaded_comments"  # noqa: SLF001
    ):
        ws._threaded_comments_cache = {}  # noqa: SLF001
        return ws._threaded_comments_cache  # noqa: SLF001
    try:
        entries = wb._rust_reader.read_threaded_comments(ws._title)  # noqa: SLF001
    except Exception:
        entries = []

    # First pass: build ``ThreadedComment`` instances keyed by GUID so we
    # can wire parent->reply links in pass two without worrying about
    # document order. Persons are resolved via the workbook registry,
    # which is itself hydrated lazily on first ``wb.persons`` access.
    persons_registry = wb.persons
    by_guid: dict[str, Any] = {}
    raw_by_guid: dict[str, dict[str, Any]] = {}
    for entry in entries:
        guid = entry.get("id")
        if not guid:
            continue
        person_id = entry.get("person_id") or ""
        person = persons_registry.by_id(person_id)
        if person is None:
            # personList is missing or stale — synthesize a placeholder so
            # the thread is still legible. Idempotent on the synthetic id.
            from wolfxl.comments._person import Person

            person = Person(name="", user_id="", provider_id="None", id=person_id or guid)
            persons_registry._seed(person)  # noqa: SLF001

        created_raw = entry.get("created")
        created: datetime | None = None
        if isinstance(created_raw, str) and created_raw:
            try:
                # Excel writes UTC ISO; ``fromisoformat`` accepts the wolfxl
                # canonical ``YYYY-MM-DDTHH:MM:SS.sss`` shape.
                created = datetime.fromisoformat(created_raw.rstrip("Z"))
            except ValueError:
                created = None
        tc = ThreadedComment(
            text=entry.get("text", "") or "",
            person=person,
            created=created,
            done=bool(entry.get("done", False)),
            id=guid,
        )
        by_guid[guid] = tc
        raw_by_guid[guid] = entry

    # Pass two: wire reply chains and pick out top-level threads.
    result: dict[str, Any] = {}
    for guid, tc in by_guid.items():
        raw = raw_by_guid[guid]
        parent_id = raw.get("parent_id")
        if parent_id is None:
            cell_ref = raw.get("cell")
            if cell_ref:
                result[cell_ref] = tc
            continue
        parent = by_guid.get(parent_id)
        if parent is None:
            # Orphan reply — treat it as top-level so the user can still
            # see the comment text rather than dropping it silently.
            cell_ref = raw.get("cell")
            if cell_ref:
                result[cell_ref] = tc
            continue
        tc.parent = parent
        parent.replies.append(tc)

    ws._threaded_comments_cache = result  # noqa: SLF001
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


class _Cfvo:
    """openpyxl-shaped cfvo anchor — ``type`` / ``val`` attributes only.

    The reader produces a flat dict (``{"type": ..., "val": ...}``) per
    cfvo and we pivot that into an attribute object so callers can mirror
    openpyxl's ``rule.colorScale.cfvo[i].type`` access shape.
    """

    __slots__ = ("type", "val")

    def __init__(self, cfvo_type: str, val: Any = None) -> None:
        self.type = cfvo_type
        self.val = val


class _ColorScaleProxy:
    """Round-tripped ``<colorScale>`` block exposing ``.cfvo`` + ``.color``.

    Mirrors the openpyxl ``ColorScale`` value object on the loaded
    :class:`~wolfxl.formatting.rule.Rule` so probes can poke
    ``rule.colorScale.cfvo[i].type`` after a save+load cycle.
    """

    __slots__ = ("cfvo", "color")

    def __init__(self, cfvo: list[_Cfvo], color: list[str]) -> None:
        self.cfvo = cfvo
        self.color = color


def _build_color_scale_proxy(payload: Any) -> _ColorScaleProxy | None:
    """Build a :class:`_ColorScaleProxy` from a Rust-side dict, or ``None``.

    The Rust reader hands us ``{"cfvo": [...], "colors": [...]}`` when a
    rule had a ``<colorScale>`` block; everything else is omitted.
    """
    if not isinstance(payload, dict):
        return None
    raw_cfvo = payload.get("cfvo") or []
    raw_colors = payload.get("colors") or []
    cfvo = [
        _Cfvo(
            cfvo_type=str(entry.get("type", "")) if isinstance(entry, dict) else "",
            val=entry.get("val") if isinstance(entry, dict) else None,
        )
        for entry in raw_cfvo
    ]
    colors = [str(c) for c in raw_colors]
    return _ColorScaleProxy(cfvo=cfvo, color=colors)


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
        rule = Rule(
            type=entry.get("rule_type") or entry.get("type") or "expression",
            operator=entry.get("operator"),
            formula=formula_list,
            stopIfTrue=bool(entry.get("stop_if_true", False)),
            priority=int(entry.get("priority", 1)),
        )
        # Attach openpyxl-shaped colorScale shim when the Rust reader
        # surfaced any cfvo/color pairs (G13).
        color_scale_payload = entry.get("color_scale")
        proxy = _build_color_scale_proxy(color_scale_payload)
        if proxy is not None:
            rule.colorScale = proxy  # type: ignore[attr-defined]
            # Also stash the round-trippable form into ``extra`` so a
            # save-time payload can rebuild the gradient without consulting
            # the proxy. Keeps the patch -> save path symmetrical.
            #
            # 2-stop maps (start, end) -> cfvo[0..2]; 3-stop maps
            # (start, mid, end) -> cfvo[0..3]. Anything else (rare) is
            # treated as the 3-stop case for the first three entries.
            extra = dict(rule.extra or {})
            n = len(proxy.cfvo)
            if n <= 2:
                index_map = (("start", 0), ("end", 1))
            else:
                index_map = (("start", 0), ("mid", 1), ("end", 2))
            for prefix, idx in index_map:
                if idx < len(proxy.cfvo):
                    cfvo_entry = proxy.cfvo[idx]
                    extra[f"{prefix}_type"] = cfvo_entry.type or None
                    extra[f"{prefix}_value"] = cfvo_entry.val
                if idx < len(proxy.color):
                    extra[f"{prefix}_color"] = proxy.color[idx]
            rule.extra = extra
        grouped[sqref].append(rule)
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
