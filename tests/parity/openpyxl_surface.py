"""The openpyxl API surface that SynthGL (and the broader target) depends on.

This module is the single source of truth for what "parity" means. Each entry
documents:

* the openpyxl import path + symbol
* where SynthGL uses it (grepped reality, not speculation)
* the category it belongs to
* a short parity note (how to check, known caveats)

The parity harness (``test_surface_smoke.py``, ``test_read_parity.py``,
``test_utils_parity.py``, ``test_write_parity.py``) consumes
``SURFACE_ENTRIES`` to decide what to test.

When SynthGL's openpyxl usage grows, add a new entry here BEFORE wiring tests.
When a WolfXL release closes a gap, flip ``wolfxl_supported`` and remove the
entry from ``tests/parity/KNOWN_GAPS.md``.

Grep once per quarter with::

    rg -n '^(from openpyxl.*import|import openpyxl)' \
       /Users/wolfgangschoenberger/Projects/SynthGL

to catch drift.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class SurfaceCategory(str, Enum):
    """Logical grouping of the openpyxl API surface.

    ``str`` mixin keeps ``SurfaceCategory.WORKBOOK_OPEN == "workbook_open"``
    true on Python 3.9-3.10 (``enum.StrEnum`` is 3.11+)."""

    WORKBOOK_OPEN = "workbook_open"
    SHEET_ACCESS = "sheet_access"
    CELL_READ = "cell_read"
    RANGE_LAYOUT = "range_layout"
    DEFINED_NAMES = "defined_names"
    WRITE = "write"
    STYLES = "styles"
    UTILS = "utils"


@dataclass(frozen=True)
class SurfaceEntry:
    """A single openpyxl symbol that must have a wolfxl equivalent."""

    openpyxl_path: str
    """Import path, e.g. ``openpyxl.utils.cell.get_column_letter``."""

    wolfxl_path: str | None
    """Equivalent wolfxl import path, or ``None`` if no equivalent exists yet.

    ``None`` means the symbol must appear in ``KNOWN_GAPS.md`` and is a HARD
    failure on the parity smoke test until a later phase closes the gap.
    """

    category: SurfaceCategory
    synthgl_usage: tuple[str, ...]
    """SynthGL file paths that import this symbol (relative to repo root)."""

    parity_note: str
    """How to check parity + known caveats."""

    wolfxl_supported: bool = True
    """Flip to ``False`` when a release ships the symbol."""

    write_api: bool = False
    """If True, used on the write path — parity check is oracle-based
    (wolfxl writes, openpyxl re-reads). If False, wolfxl reads an openpyxl-
    authored file."""

    tags: frozenset[str] = field(default_factory=frozenset)
    """Free-form labels (``"hard"``, ``"soft"``, ``"info"``) matching the
    scoring tiers in the plan. Default is ``"hard"`` — mismatches block CI."""


# ---------------------------------------------------------------------------
# WORKBOOK OPEN — SynthGL reads and writes xlsx files via load_workbook + Workbook
# ---------------------------------------------------------------------------
_OPEN_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="openpyxl.load_workbook",
        wolfxl_path="wolfxl.load_workbook",
        category=SurfaceCategory.WORKBOOK_OPEN,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
            "packages/synthgl-lrbench/src/synthgl/lrbench/generators/f3_normalization.py",
        ),
        parity_note=(
            "Accept ``data_only``, ``read_only``, ``keep_links`` kwargs without "
            "error. ``data_only=True`` must return cached formula results when present. "
            "Return a workbook with iterable ``sheetnames`` and ``__getitem__``. "
            "Password kwarg lands in Phase 2."
        ),
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.Workbook",
        wolfxl_path="wolfxl.Workbook",
        category=SurfaceCategory.WORKBOOK_OPEN,
        synthgl_usage=(
            "src/synthgl/app/services/ingest_confidence_generated.py",
            "tests/ingest/test_round_trip_compound.py",
        ),
        parity_note=(
            "``Workbook()`` constructs a write-mode workbook with a default 'Sheet'. "
            "``wb.active`` is the default sheet, ``wb.create_sheet(title)`` adds one."
        ),
        write_api=True,
        tags=frozenset({"hard"}),
    ),
)

# ---------------------------------------------------------------------------
# SHEET ACCESS
# ---------------------------------------------------------------------------
_SHEET_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="openpyxl.worksheet.worksheet.Worksheet",
        wolfxl_path="wolfxl.Worksheet",
        category=SurfaceCategory.SHEET_ACCESS,
        synthgl_usage=(
            "src/synthgl/app/adapters/excel_parser.py",
        ),
        parity_note=(
            "Re-exported at top level via ``wolfxl.Worksheet`` for type-hint "
            "narrowing. Backed by ``wolfxl._worksheet.Worksheet``."
        ),
        tags=frozenset({"hard", "type-import"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.cell.cell.Cell",
        wolfxl_path="wolfxl.Cell",
        category=SurfaceCategory.SHEET_ACCESS,
        synthgl_usage=(
            "src/synthgl/app/adapters/excel_parser.py",
        ),
        parity_note=(
            "Re-exported at top level via ``wolfxl.Cell`` for type-hint "
            "narrowing. Backed by ``wolfxl._cell.Cell``."
        ),
        tags=frozenset({"hard", "type-import"}),
    ),
)

# ---------------------------------------------------------------------------
# CELL READ — value, number_format, font, fill, border, alignment
# ---------------------------------------------------------------------------
_CELL_READ_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="Cell.value",
        wolfxl_path="wolfxl._cell.Cell.value",
        category=SurfaceCategory.CELL_READ,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
            "src/synthgl/app/adapters/excel_parser.py",
        ),
        parity_note=(
            "Values MUST match byte-for-byte for string/number/bool/None. Dates "
            "+ datetimes must compare equal via ``==``. Midnight date cells must "
            "surface as ``datetime`` objects (matching openpyxl's read contract). "
            "Formulas return formula text by default and cached values under "
            "``data_only=True``."
        ),
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Cell.number_format",
        wolfxl_path="wolfxl._cell.Cell.number_format",
        category=SurfaceCategory.CELL_READ,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
        ),
        parity_note=(
            "Round-trip the raw Excel format code. openpyxl returns ``'General'`` "
            "for unformatted cells; wolfxl should too. Critical for the date-vs-"
            "number heuristic in ``is_date_format``."
        ),
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Cell.font",
        wolfxl_path="wolfxl._cell.Cell.font",
        category=SurfaceCategory.CELL_READ,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
        ),
        parity_note=(
            "Parity on ``name, size, bold, italic, underline, color``. "
            "Strike + family/scheme currently INFO-tier."
        ),
        tags=frozenset({"soft"}),
    ),
    SurfaceEntry(
        openpyxl_path="Cell.fill",
        wolfxl_path="wolfxl._cell.Cell.fill",
        category=SurfaceCategory.CELL_READ,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
        ),
        parity_note="Compare ``patternType`` + ``fgColor`` hex.",
        tags=frozenset({"soft"}),
    ),
    SurfaceEntry(
        openpyxl_path="Cell.border",
        wolfxl_path="wolfxl._cell.Cell.border",
        category=SurfaceCategory.CELL_READ,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
        ),
        parity_note="Compare each edge's ``style`` + ``color``. Diagonal is INFO.",
        tags=frozenset({"soft"}),
    ),
    SurfaceEntry(
        openpyxl_path="Cell.alignment",
        wolfxl_path="wolfxl._cell.Cell.alignment",
        category=SurfaceCategory.CELL_READ,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
        ),
        parity_note="Parity on ``horizontal, vertical, wrap_text``. Indent + rotation SOFT.",
        tags=frozenset({"soft"}),
    ),
)

# ---------------------------------------------------------------------------
# RANGE / LAYOUT — dimensions, merged cells, freeze, column widths
# ---------------------------------------------------------------------------
_RANGE_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="Worksheet.max_row",
        wolfxl_path="wolfxl._worksheet.Worksheet.max_row",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
        ),
        parity_note=(
            "Shipped as a ``@property`` wrapper around ``_max_row()`` on the "
            "Worksheet class (in both read and write modes)."
        ),
        tags=frozenset({"hard", "api-shape"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.max_column",
        wolfxl_path="wolfxl._worksheet.Worksheet.max_column",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
        ),
        parity_note=(
            "Shipped as a ``@property`` wrapper around ``_max_col()`` on the "
            "Worksheet class. Note: openpyxl uses ``max_column`` (not "
            "``max_col``); we mirror the longer name."
        ),
        tags=frozenset({"hard", "api-shape"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.merged_cells",
        wolfxl_path="wolfxl._worksheet.Worksheet.merged_cells",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
        ),
        parity_note=(
            "Shipped as a ``@property`` returning ``_MergedCellsProxy``. The "
            "proxy exposes ``.ranges`` (list of range strings) backed by the "
            "Rust reader's ``read_merged_ranges`` in read mode and the "
            "in-memory ``_merged_ranges`` set in write mode."
        ),
        tags=frozenset({"hard", "api-shape"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.freeze_panes",
        wolfxl_path="wolfxl._worksheet.Worksheet.freeze_panes",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
        ),
        parity_note="Cell-reference string (e.g. ``'B2'``) or ``None``.",
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.iter_rows",
        wolfxl_path="wolfxl._worksheet.Worksheet.iter_rows",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
            "src/synthgl/app/adapters/excel_parser.py",
        ),
        parity_note=(
            "Signature: ``iter_rows(min_row, max_row, min_col, max_col, "
            "values_only)``. ``values_only=True`` yields tuples of raw values."
        ),
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.column_dimensions[letter].width",
        wolfxl_path="wolfxl._worksheet.Worksheet.column_dimensions",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
        ),
        parity_note=(
            "Property-of-proxy access. Smoke test verifies the proxy attribute "
            "exists; round-trip behavior covered by ``test_write_parity``. "
            "Note: wolfxl's column width on round-trip differs from openpyxl by "
            "~0.7 (rust_xlsxwriter applies an additional padding constant); "
            "tolerance widened in the write parity test pending Phase 0 "
            "investigation."
        ),
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.row_dimensions[i].height",
        wolfxl_path="wolfxl._worksheet.Worksheet.row_dimensions",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
        ),
        parity_note="INFO-tier per plan; not currently exercised in parity.",
        tags=frozenset({"info"}),
    ),
)

# ---------------------------------------------------------------------------
# DEFINED NAMES
# ---------------------------------------------------------------------------
_DEFINED_NAME_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="Workbook.defined_names",
        wolfxl_path="wolfxl._workbook.Workbook.defined_names",
        category=SurfaceCategory.DEFINED_NAMES,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
        ),
        parity_note=(
            "Read: openpyxl returns ``DefinedNameDict``; wolfxl returns "
            "``dict[str, str]``. SynthGL only needs ``name -> refers_to`` lookup + "
            "iteration, so dict is acceptable. Write: Phase 1 work."
        ),
        tags=frozenset({"hard"}),
    ),
)

# ---------------------------------------------------------------------------
# WRITE — cell assignment, style assignment, save
# ---------------------------------------------------------------------------
_WRITE_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="Cell.value = x",
        wolfxl_path="wolfxl._cell.Cell.value (setter)",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(
            "src/synthgl/app/services/ingest_confidence_generated.py",
            "tests/ingest/test_round_trip_compound.py",
        ),
        parity_note=(
            "Setting a Python value should round-trip: write, reopen via openpyxl, "
            "read the same value. Covered by ``test_write_parity``."
        ),
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.append(iterable)",
        wolfxl_path="wolfxl._worksheet.Worksheet.append",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(
            "tests/ingest/test_round_trip_compound.py",
        ),
        parity_note="Values appended to next free row starting at column A.",
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.merge_cells(range)",
        wolfxl_path="wolfxl._worksheet.Worksheet.merge_cells",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
        ),
        parity_note="After save + reopen, range must appear in ``merged_cells``.",
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Workbook.save(path)",
        wolfxl_path="wolfxl._workbook.Workbook.save",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(
            "src/synthgl/app/services/ingest_confidence_generated.py",
            "tests/ingest/test_round_trip_compound.py",
        ),
        parity_note="Produces a valid xlsx file openable by openpyxl.",
        write_api=True,
        tags=frozenset({"hard"}),
    ),
)

# ---------------------------------------------------------------------------
# STYLE CONSTRUCTION — used on write paths
# ---------------------------------------------------------------------------
_STYLE_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="openpyxl.styles.Font",
        wolfxl_path="wolfxl.Font",
        category=SurfaceCategory.STYLES,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
            "tests/ingest/test_excel_adapter.py",
        ),
        parity_note=(
            "Frozen dataclass accepting ``name, size, bold, italic, underline, "
            "color``. Must accept ``color=Color(rgb=...)`` and raw hex strings."
        ),
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.styles.PatternFill",
        wolfxl_path="wolfxl.PatternFill",
        category=SurfaceCategory.STYLES,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
        ),
        parity_note=(
            "Must accept both ``patternType=`` and ``fill_type=`` kwargs "
            "(openpyxl alias). Already handled in wolfxl 0.3.2."
        ),
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.styles.Alignment",
        wolfxl_path="wolfxl.Alignment",
        category=SurfaceCategory.STYLES,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
            "tests/ingest/test_excel_adapter.py",
        ),
        parity_note="``horizontal, vertical, wrap_text, text_rotation, indent`` kwargs.",
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.styles.Border",
        wolfxl_path="wolfxl.Border",
        category=SurfaceCategory.STYLES,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
            "tests/ingest/test_excel_adapter.py",
        ),
        parity_note="Four-edge dataclass with ``left/right/top/bottom: Side``.",
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.styles.Side",
        wolfxl_path="wolfxl.Side",
        category=SurfaceCategory.STYLES,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
            "tests/ingest/test_excel_adapter.py",
        ),
        parity_note="``style, color`` kwargs.",
        write_api=True,
        tags=frozenset({"hard"}),
    ),
)

# ---------------------------------------------------------------------------
# UTILS — the gnarliest area because openpyxl's public/private line is fuzzy
# ---------------------------------------------------------------------------
_UTILS_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="openpyxl.utils.cell.get_column_letter",
        wolfxl_path="wolfxl.utils.cell.get_column_letter",
        category=SurfaceCategory.UTILS,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
            "packages/synthgl-lrbench-agent/src/synthgl/lrbench_agent/formula_awareness.py",
            "packages/synthgl-lrbench-agent/src/synthgl/lrbench_agent/server.py",
            "packages/synthgl-lrbench-agent/src/synthgl/lrbench_agent/xlsx_reader.py",
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
        ),
        parity_note=(
            "Shipped via ``wolfxl.utils.cell.get_column_letter``; mirrors "
            "openpyxl's 1..18278 (ZZZ) bound and ``ValueError`` message."
        ),
        tags=frozenset({"hard", "api-rename"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.utils.cell.column_index_from_string",
        wolfxl_path="wolfxl.utils.cell.column_index_from_string",
        category=SurfaceCategory.UTILS,
        synthgl_usage=(),
        parity_note=(
            "Shipped via ``wolfxl.utils.cell.column_index_from_string`` — "
            "inverse of ``get_column_letter``."
        ),
        tags=frozenset({"hard", "api-rename"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.utils.cell.range_boundaries",
        wolfxl_path="wolfxl.utils.cell.range_boundaries",
        category=SurfaceCategory.UTILS,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
            "packages/synthgl-lrbench-agent/src/synthgl/lrbench_agent/server.py",
        ),
        parity_note=(
            "Shipped via ``wolfxl.utils.cell.range_boundaries`` — accepts "
            "absolute refs (``$A$1:$D$10``) and whole-column/row refs (``A:A``, "
            "``1:1``) with the same return shape as openpyxl."
        ),
        tags=frozenset({"hard", "missing"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.utils.cell.coordinate_to_tuple",
        wolfxl_path="wolfxl.utils.cell.coordinate_to_tuple",
        category=SurfaceCategory.UTILS,
        synthgl_usage=(),
        parity_note=(
            "Shipped via ``wolfxl.utils.cell.coordinate_to_tuple`` — returns "
            "``(row, col)`` 1-based, matching openpyxl."
        ),
        tags=frozenset({"hard", "api-rename"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.styles.numbers.is_date_format",
        wolfxl_path="wolfxl.utils.numbers.is_date_format",
        category=SurfaceCategory.UTILS,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
        ),
        parity_note=(
            "Shipped via ``wolfxl.utils.numbers.is_date_format`` — ports "
            "openpyxl's STRIP_RE + DATE_TOKEN_RE bug-for-bug, including the "
            "``[locale]`` vs ``[h]/[mm]/[ss]`` distinction."
        ),
        tags=frozenset({"hard", "missing"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.utils.datetime.from_excel",
        wolfxl_path="wolfxl.utils.datetime.from_excel",
        category=SurfaceCategory.UTILS,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
        ),
        parity_note=(
            "Shipped via ``wolfxl.utils.datetime.from_excel`` — reproduces "
            "the 1900 leap-year bug (epoch=1899-12-30, +1 day for serials in "
            "(0, 60))."
        ),
        tags=frozenset({"hard", "missing"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.utils.datetime.CALENDAR_WINDOWS_1900",
        wolfxl_path="wolfxl.utils.datetime.CALENDAR_WINDOWS_1900",
        category=SurfaceCategory.UTILS,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
        ),
        parity_note=(
            "Shipped via ``wolfxl.utils.datetime.CALENDAR_WINDOWS_1900`` as a "
            "``datetime(1899, 12, 30)`` sentinel (matches openpyxl after its "
            "internal re-bind)."
        ),
        tags=frozenset({"hard", "missing"}),
    ),
)


SURFACE_ENTRIES: tuple[SurfaceEntry, ...] = (
    *_OPEN_ENTRIES,
    *_SHEET_ENTRIES,
    *_CELL_READ_ENTRIES,
    *_RANGE_ENTRIES,
    *_DEFINED_NAME_ENTRIES,
    *_WRITE_ENTRIES,
    *_STYLE_ENTRIES,
    *_UTILS_ENTRIES,
)


def entries_by_category(category: SurfaceCategory) -> tuple[SurfaceEntry, ...]:
    """Filter entries by category — used by targeted test files."""
    return tuple(e for e in SURFACE_ENTRIES if e.category == category)


def supported_entries() -> tuple[SurfaceEntry, ...]:
    """Entries WolfXL currently claims to support — these run in CI."""
    return tuple(e for e in SURFACE_ENTRIES if e.wolfxl_supported)


def known_gap_entries() -> tuple[SurfaceEntry, ...]:
    """Entries tracked in KNOWN_GAPS.md — skipped with xfail on the harness."""
    return tuple(e for e in SURFACE_ENTRIES if not e.wolfxl_supported)
