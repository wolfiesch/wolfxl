"""Native ``.xlsb`` read parity against committed sidecar goldens."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

import wolfxl

FIXTURES_DIR = Path(__file__).parent / "fixtures" / "xlsb"


def _all_fixtures() -> list[Path]:
    return sorted(FIXTURES_DIR.glob("*.xlsb"))


_FIXTURES = _all_fixtures()


pytestmark = pytest.mark.skipif(
    not _FIXTURES,
    reason="No .xlsb fixtures present (Sprint Κ Pod-γ)",
)


def _coerce(v: object) -> object:
    """Normalize workbook values for JSON-sidecar equality."""
    if hasattr(v, "isoformat"):
        return v.isoformat()  # type: ignore[no-any-return]
    return v


def _trim_trailing_empty(rows: list[list[object]]) -> list[list[object]]:
    while rows and all(value is None for value in rows[-1]):
        rows.pop()
    return rows


def _sheet_values(ws: object) -> list[list[object]]:
    rows = [
        [_coerce(cell.value) for cell in row]
        for row in ws.iter_rows()  # type: ignore[attr-defined]
    ]
    return _trim_trailing_empty(rows)


@pytest.mark.parametrize("fixture", _FIXTURES, ids=lambda p: p.name)
def test_xlsb_values_match_committed_goldens(fixture: Path) -> None:
    """Native xlsb reads match committed dependency-free value sidecars."""
    wb = wolfxl.load_workbook(str(fixture), data_only=True)
    expected = json.loads(fixture.with_suffix(".golden.json").read_text())
    actual = {sheet_name: _sheet_values(wb[sheet_name]) for sheet_name in wb.sheetnames}
    assert actual == expected


def test_xlsb_modify_raises() -> None:
    fixture = _FIXTURES[0]
    with pytest.raises(NotImplementedError, match="transcribe"):
        wolfxl.load_workbook(str(fixture), modify=True)


def test_xlsb_read_only_raises() -> None:
    fixture = _FIXTURES[0]
    with pytest.raises(NotImplementedError, match="xlsx-only"):
        wolfxl.load_workbook(str(fixture), read_only=True)


def test_xlsb_password_raises() -> None:
    fixture = _FIXTURES[0]
    with pytest.raises(NotImplementedError, match="xlsx-only"):
        wolfxl.load_workbook(str(fixture), password="anything")


def test_xlsb_cell_styles_are_readable() -> None:
    """Native xlsb exposes read-side style accessors."""
    fixture = _FIXTURES[0]
    wb = wolfxl.load_workbook(str(fixture))
    ws = wb.active
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                assert cell.font is not None
                assert cell.fill is not None
                assert cell.border is not None
                assert cell.alignment is not None
                _ = cell.number_format
                return
    pytest.fail("no non-empty cells in fixture")


def test_xlsb_from_bytes() -> None:
    fixture = _FIXTURES[0]
    data = fixture.read_bytes()
    wb_bytes = wolfxl.load_workbook(data)
    wb_path = wolfxl.load_workbook(str(fixture))
    assert wb_bytes.sheetnames == wb_path.sheetnames


def test_xlsb_classify_format() -> None:
    """``wolfxl.classify_file_format`` reports 'xlsb' for this fixture
    both as a path and as bytes.

    Note: ``wolfxl.classify_format`` (without ``_file_``) is a separate,
    long-standing SynthGL number-format archetype classifier. The
    Sprint Κ file-format detector lives at
    ``wolfxl.classify_file_format`` (re-exported from
    ``wolfxl._rust.classify_file_format``).
    """
    fixture = _FIXTURES[0]
    fmt_path = wolfxl.classify_file_format(str(fixture))
    assert fmt_path == "xlsb", f"path -> {fmt_path!r}"
    fmt_bytes = wolfxl.classify_file_format(fixture.read_bytes())
    assert fmt_bytes == "xlsb", f"bytes -> {fmt_bytes!r}"
