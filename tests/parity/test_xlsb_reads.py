"""Sprint Κ Pod-γ — ``.xlsb`` read parity vs ``pandas + calamine``.

These tests assert that ``wolfxl.load_workbook`` returns the same cell
values as ``pandas.read_excel(engine="calamine")`` for every committed
``.xlsb`` fixture. They also pin down the fail-fast behaviour for
xlsx-only options (``modify=``, ``read_only=``, ``password=``) and the
``NotImplementedError`` raised by style accessors on legacy formats.

Will go green once Pod-α (xls reader) and Pod-β (xlsb reader) land.
"""

from __future__ import annotations

from pathlib import Path

import pytest

import wolfxl

pd = pytest.importorskip("pandas")
pytest.importorskip("python_calamine")

FIXTURES_DIR = Path(__file__).parent / "fixtures" / "xlsb"


def _all_fixtures() -> list[Path]:
    return sorted(FIXTURES_DIR.glob("*.xlsb"))


_FIXTURES = _all_fixtures()


pytestmark = pytest.mark.skipif(
    not _FIXTURES,
    reason="No .xlsb fixtures present (Sprint Κ Pod-γ)",
)


def _coerce(v: object) -> object:
    """Normalise Python values for cross-engine equality."""
    # pandas-calamine surfaces empty cells as NaN; wolfxl uses None.
    if v is None:
        return None
    if isinstance(v, float):
        # NaN equality is not reflexive — treat as None.
        if v != v:  # noqa: PLR0124
            return None
    return v


@pytest.mark.parametrize("fixture", _FIXTURES, ids=lambda p: p.name)
def test_xlsb_values_match_pandas_calamine(fixture: Path) -> None:
    """wolfxl.load_workbook reads same cell values as pandas+calamine."""
    wb = wolfxl.load_workbook(str(fixture))
    df = pd.read_excel(
        str(fixture), engine="calamine", sheet_name=None, header=None,
    )

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if sheet_name not in df:
            # pandas+calamine and wolfxl might disagree on chart-only sheets
            # being "sheets". If wolfxl exposes one but pandas doesn't, the
            # wolfxl-side sheet must therefore be empty.
            assert all(
                cell.value is None
                for row in ws.iter_rows()
                for cell in row
            ), f"{fixture.name}: {sheet_name!r} unique to wolfxl with content"
            continue

        df_sheet = df[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                cv = _coerce(cell.value)
                if cv is None:
                    continue
                # df_sheet is 0-indexed, cell.row/column are 1-indexed.
                if (cell.row - 1) >= df_sheet.shape[0]:
                    continue
                if (cell.column - 1) >= df_sheet.shape[1]:
                    continue
                df_value = _coerce(df_sheet.iat[cell.row - 1, cell.column - 1])
                if df_value is None:
                    continue
                # Numeric: float compare with tolerance.
                if isinstance(cv, (int, float)) and isinstance(
                    df_value, (int, float)
                ):
                    assert abs(float(cv) - float(df_value)) < 1e-9, (
                        f"{fixture.name}!{sheet_name}!{cell.coordinate}: "
                        f"wolfxl={cv} pandas={df_value}"
                    )
                else:
                    assert cv == df_value or str(cv) == str(df_value), (
                        f"{fixture.name}!{sheet_name}!{cell.coordinate}: "
                        f"wolfxl={cv!r} pandas={df_value!r}"
                    )


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


def test_xlsb_cell_font_raises() -> None:
    """Style accessors are xlsx-only; xlsb must surface NotImplementedError."""
    fixture = _FIXTURES[0]
    wb = wolfxl.load_workbook(str(fixture))
    ws = wb.active
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                with pytest.raises(NotImplementedError, match="xlsx-only"):
                    _ = cell.font
                return
    pytest.fail("no non-empty cells in fixture")


def test_xlsb_from_bytes() -> None:
    fixture = _FIXTURES[0]
    data = fixture.read_bytes()
    wb_bytes = wolfxl.load_workbook(data)
    wb_path = wolfxl.load_workbook(str(fixture))
    assert wb_bytes.sheetnames == wb_path.sheetnames


def test_xlsb_classify_format() -> None:
    """``wolfxl.classify_format`` (file-format detection variant) reports
    'xlsb' for this fixture both as a path and as bytes.

    Pod-β is responsible for adding a path/bytes-aware format classifier.
    The existing ``_rust.classify_format`` is the SynthGL archetype
    classifier and is unrelated to file format.
    """
    fixture = _FIXTURES[0]
    fmt_path = wolfxl.classify_format(str(fixture))
    assert fmt_path == "xlsb", f"path -> {fmt_path!r}"
    fmt_bytes = wolfxl.classify_format(fixture.read_bytes())
    assert fmt_bytes == "xlsb", f"bytes -> {fmt_bytes!r}"
