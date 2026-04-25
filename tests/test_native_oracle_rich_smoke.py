"""W4B smoke gate: NativeWorkbook wires 8 rich-feature pymethods.

Builds the SAME workbook under WOLFXL_WRITER=oracle and WOLFXL_WRITER=native,
unzips both outputs, and structurally diffs the rich-feature OOXML parts.

Goal: confirm that all 8 implemented pymethods produce valid OOXML. The
comparison is structural (presence of key elements + semantic content), not
byte-level, because the two backends produce legitimately different part
layouts (e.g. oracle places comments at ``xl/comments1.xml`` while native
uses ``xl/comments/comments1.xml``).
"""
from __future__ import annotations

import os
import subprocess
import sys
import textwrap
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest


# ---------------------------------------------------------------------------
# Fixture script — same workbook for both backends
# ---------------------------------------------------------------------------

_FIXTURE_SCRIPT = textwrap.dedent(
    r"""
    import sys, os
    import wolfxl

    wb = wolfxl.Workbook()
    ws = wb.active
    ws.title = "Data"

    # Populate a few cells so CF / DV ranges have content
    ws["A1"].value = 42
    ws["B1"].value = "hello"
    ws["C1"].value = "world"
    for r in range(2, 6):
        ws.cell(row=r, column=4, value=r * 10)
        ws.cell(row=r, column=5, value=r * 20)
        ws.cell(row=r, column=6, value=r * 30)

    w = wb._rust_writer

    # (1) Hyperlink: A1 → external
    w.add_hyperlink("Data", {
        "cell": "A1",
        "target": "https://example.com",
        "display": "Example",
        "tooltip": "Click",
    })

    # (2) Comment on B1 from Alice, (3) Comment on C1 from Bob.
    # These two in order prove IndexMap author-order preservation.
    w.add_comment("Data", {"cell": "B1", "text": "Alice comment", "author": "Alice"})
    w.add_comment("Data", {"cell": "C1", "text": "Bob comment",   "author": "Bob"})

    # (4) Print area
    ws.print_area = "A1:D10"

    # (5) Conditional format: cellIs > 100, yellow bg
    w.add_conditional_format("Data", {
        "range": "A1:A10",
        "rule_type": "cellIs",
        "operator": "greaterThan",
        "formula": "100",
        "format": {"bg_color": "#FFFF00"},
    })

    # (6) Data validation: list on B1:B10
    w.add_data_validation("Data", {
        "range": "B1:B10",
        "validation_type": "list",
        "formula1": '"Red,Green,Blue"',
    })

    # (7) Named range: workbook scope
    w.add_named_range("Data", {
        "name":      "MyRange",
        "scope":     "workbook",
        "refers_to": "Data!$A$1:$A$10",
    })

    # (8) Table: D1:F5 with 3 columns
    w.add_table("Data", {
        "name":       "MyTable",
        "ref":        "D1:F5",
        "style":      "TableStyleMedium9",
        "columns":    ["Col A", "Col B", "Col C"],
        "header_row": True,
        "totals_row": False,
    })

    # Properties (native-only — oracle has no set_properties method)
    if hasattr(w, "set_properties"):
        w.set_properties({
            "title":   "Test",
            "creator": "Claude",
            "subject": "W4B smoke",
        })

    wb.save(sys.argv[1])
    """
)


def _save_under(env_value: str, target: Path) -> None:
    """Run the fixture script under the given WOLFXL_WRITER backend.

    ``WOLFXL_TEST_EPOCH=0`` forces ZIP entry mtimes to the Unix epoch in the
    native writer so two runs produce byte-identical xlsx output. Required for
    differential / golden-file tests per CLAUDE.md.
    """
    env = {**os.environ, "WOLFXL_WRITER": env_value, "WOLFXL_TEST_EPOCH": "0"}
    result = subprocess.run(
        [sys.executable, "-c", _FIXTURE_SCRIPT, str(target)],
        env=env,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        pytest.fail(
            f"WOLFXL_WRITER={env_value} save failed:\n"
            f"stdout: {result.stdout}\nstderr: {result.stderr}"
        )


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------

def _find_part(namelist: list[str], *suffixes: str) -> str | None:
    """Return the first part name that ends with any of the given suffixes."""
    for suffix in suffixes:
        for name in namelist:
            if name.endswith(suffix):
                return name
    return None


def _parse(zf: zipfile.ZipFile, part: str) -> ET.Element:
    return ET.fromstring(zf.read(part))


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

@pytest.mark.smoke
def test_rich_features_part_presence(tmp_path: Path) -> None:
    """Every rich-feature part the oracle emits must also exist in native output."""
    oracle_path = tmp_path / "oracle.xlsx"
    native_path = tmp_path / "native.xlsx"
    _save_under("oracle", oracle_path)
    _save_under("native", native_path)

    with zipfile.ZipFile(oracle_path) as oz, zipfile.ZipFile(native_path) as nz:
        oracle_parts = set(oz.namelist())
        native_parts = set(nz.namelist())

    # Core structural parts both backends must produce
    required = [
        "xl/workbook.xml",
        "xl/styles.xml",
        "xl/worksheets/sheet1.xml",
        "docProps/core.xml",
    ]
    for part in required:
        assert part in oracle_parts, f"oracle missing required part: {part}"
        assert part in native_parts, f"native missing required part: {part}"

    # Table part (both backends register a table for D1:F5)
    oracle_table = _find_part(list(oracle_parts), "table1.xml")
    native_table = _find_part(list(native_parts), "table1.xml")
    assert oracle_table is not None, f"oracle missing table part; parts={sorted(oracle_parts)}"
    assert native_table is not None, f"native missing table part; parts={sorted(native_parts)}"

    # VML drawing (both backends produce comments → VML)
    oracle_vml = _find_part(list(oracle_parts), "vmlDrawing1.vml")
    native_vml = _find_part(list(native_parts), "vmlDrawing1.vml")
    assert oracle_vml is not None, f"oracle missing VML; parts={sorted(oracle_parts)}"
    assert native_vml is not None, f"native missing VML; parts={sorted(native_parts)}"


@pytest.mark.smoke
def test_rich_features_comments_author_order(tmp_path: Path) -> None:
    """Comments must preserve IndexMap author order: Alice=0, Bob=1.

    This specifically validates the BTreeMap-vs-IndexMap author ordering fix
    that motivated the native writer rewrite.
    """
    native_path = tmp_path / "native.xlsx"
    _save_under("native", native_path)

    with zipfile.ZipFile(native_path) as nz:
        # Native uses xl/comments/comments1.xml
        comments_part = _find_part(nz.namelist(), "comments1.xml")
        assert comments_part is not None, (
            f"native missing comments part; parts={sorted(nz.namelist())}"
        )
        root = _parse(nz, comments_part)

    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    authors = [a.text or "" for a in root.findall(f".//{{{ns}}}author")]
    assert len(authors) == 2, f"expected 2 authors, got {authors}"
    assert authors[0] == "Alice", f"author[0] should be Alice (insertion order), got {authors!r}"
    assert authors[1] == "Bob",   f"author[1] should be Bob  (insertion order), got {authors!r}"


@pytest.mark.smoke
def test_rich_features_comments_oracle_author_order(tmp_path: Path) -> None:
    """Oracle output must contain both Alice and Bob authors, Alice first.

    Note: rust_xlsxwriter (the oracle backend) inserts a default "Author"
    entry in addition to the named authors. The key invariant is that Alice
    comes before Bob in insertion order (the IndexMap fix is on the native
    side; the oracle's behavior is a reference point, not a strict model).
    """
    oracle_path = tmp_path / "oracle.xlsx"
    _save_under("oracle", oracle_path)

    with zipfile.ZipFile(oracle_path) as oz:
        # Oracle uses xl/comments1.xml (in package root, not a subdirectory)
        comments_part = _find_part(oz.namelist(), "comments1.xml")
        assert comments_part is not None, (
            f"oracle missing comments part; parts={sorted(oz.namelist())}"
        )
        root = _parse(oz, comments_part)

    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    authors = [a.text or "" for a in root.findall(f".//{{{ns}}}author")]
    assert "Alice" in authors, f"oracle authors missing Alice; got {authors!r}"
    assert "Bob" in authors, f"oracle authors missing Bob; got {authors!r}"
    alice_idx = authors.index("Alice")
    bob_idx = authors.index("Bob")
    assert alice_idx < bob_idx, (
        f"oracle: Alice must appear before Bob (insertion order); got {authors!r}"
    )


@pytest.mark.smoke
def test_rich_features_workbook_defined_names(tmp_path: Path) -> None:
    """Both backends must emit MyRange and a Print_Area defined name."""
    oracle_path = tmp_path / "oracle.xlsx"
    native_path = tmp_path / "native.xlsx"
    _save_under("oracle", oracle_path)
    _save_under("native", native_path)

    for label, path in (("oracle", oracle_path), ("native", native_path)):
        with zipfile.ZipFile(path) as zf:
            wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
        # Both backends emit the OOXML built-in name `_xlnm.Print_Area`.
        # Asserting the exact form catches casing regressions in the emitter.
        assert "_xlnm.Print_Area" in wb_xml, (
            f"{label} workbook.xml missing _xlnm.Print_Area defined name"
        )
        assert "MyRange" in wb_xml, (
            f"{label} workbook.xml missing MyRange defined name"
        )


@pytest.mark.smoke
def test_rich_features_sheet_cf_and_dv(tmp_path: Path) -> None:
    """Both backends must emit conditionalFormatting + dataValidation with the
    same semantic payload (cellIs > 100, list ``"Red,Green,Blue"``)."""
    oracle_path = tmp_path / "oracle.xlsx"
    native_path = tmp_path / "native.xlsx"
    _save_under("oracle", oracle_path)
    _save_under("native", native_path)

    for label, path in (("oracle", oracle_path), ("native", native_path)):
        with zipfile.ZipFile(path) as zf:
            sheet_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert "conditionalFormatting" in sheet_xml, (
            f"{label} sheet1.xml missing <conditionalFormatting>"
        )
        # Semantic CF payload: cellIs operator + threshold formula.
        assert 'type="cellIs"' in sheet_xml, (
            f"{label} sheet1.xml CF missing type=\"cellIs\""
        )
        assert 'operator="greaterThan"' in sheet_xml, (
            f"{label} sheet1.xml CF missing operator=\"greaterThan\""
        )
        assert "<formula>100</formula>" in sheet_xml, (
            f"{label} sheet1.xml CF missing <formula>100</formula>"
        )

        assert "dataValidation" in sheet_xml, (
            f"{label} sheet1.xml missing <dataValidation>"
        )
        # Semantic DV payload: list type + the literal "Red,Green,Blue" value
        # set. The list literal is XML-escaped (&quot;) inside <formula1>.
        assert 'type="list"' in sheet_xml, (
            f"{label} sheet1.xml DV missing type=\"list\""
        )
        assert "Red,Green,Blue" in sheet_xml, (
            f"{label} sheet1.xml DV missing list values 'Red,Green,Blue'"
        )


@pytest.mark.smoke
def test_rich_features_hyperlink_presence(tmp_path: Path) -> None:
    """Both backends must emit an external hyperlink on A1 with ref + r:id, and
    the target URL must surface in the sheet's external relationships file.

    External hyperlinks use ``r:id`` (Target lives in
    ``xl/worksheets/_rels/sheet1.xml.rels``); internal hyperlinks would carry a
    ``location`` attribute on ``<hyperlink>`` instead. The fixture uses
    ``https://example.com`` so external is the asserted shape.
    """
    oracle_path = tmp_path / "oracle.xlsx"
    native_path = tmp_path / "native.xlsx"
    _save_under("oracle", oracle_path)
    _save_under("native", native_path)

    for label, path in (("oracle", oracle_path), ("native", native_path)):
        with zipfile.ZipFile(path) as zf:
            sheet_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
            rels_part = _find_part(zf.namelist(), "worksheets/_rels/sheet1.xml.rels")
            assert rels_part is not None, (
                f"{label} missing sheet1 rels; parts={sorted(zf.namelist())}"
            )
            rels_xml = zf.read(rels_part).decode("utf-8")

        # Hyperlink element must exist on A1 (the cell the fixture set).
        assert "<hyperlink" in sheet_xml, (
            f"{label} sheet1.xml missing <hyperlink> element"
        )
        assert 'ref="A1"' in sheet_xml, (
            f"{label} sheet1.xml hyperlink missing ref=\"A1\""
        )
        # External target → r:id attr on <hyperlink>, target URL in the rels.
        assert "r:id=" in sheet_xml, (
            f"{label} external hyperlink missing r:id attr (would indicate "
            f"misclassified as internal)"
        )
        assert "https://example.com" in rels_xml, (
            f"{label} sheet1.xml.rels missing target https://example.com; "
            f"rels={rels_xml!r}"
        )


@pytest.mark.smoke
def test_rich_features_table_structure(tmp_path: Path) -> None:
    """Both backends must emit a table part with name 'MyTable', the D1:F5
    range, and all three header column names ('Col A', 'Col B', 'Col C').
    """
    oracle_path = tmp_path / "oracle.xlsx"
    native_path = tmp_path / "native.xlsx"
    _save_under("oracle", oracle_path)
    _save_under("native", native_path)

    for label, path in (("oracle", oracle_path), ("native", native_path)):
        with zipfile.ZipFile(path) as zf:
            table_part = _find_part(zf.namelist(), "table1.xml")
            assert table_part is not None, (
                f"{label} missing table1.xml; parts={sorted(zf.namelist())}"
            )
            table_xml = zf.read(table_part).decode("utf-8")
        assert 'name="MyTable"' in table_xml, (
            f"{label} table1.xml missing name=\"MyTable\""
        )
        assert 'ref="D1:F5"' in table_xml, (
            f"{label} table1.xml missing ref=\"D1:F5\""
        )
        for col_name in ("Col A", "Col B", "Col C"):
            assert col_name in table_xml, (
                f"{label} table1.xml missing column '{col_name}'; "
                f"table={table_xml!r}"
            )


@pytest.mark.smoke
def test_rich_features_native_doc_props(tmp_path: Path) -> None:
    """Native backend must emit title/creator/subject in docProps/core.xml."""
    native_path = tmp_path / "native.xlsx"
    _save_under("native", native_path)

    with zipfile.ZipFile(native_path) as nz:
        core_xml = nz.read("docProps/core.xml").decode("utf-8")

    # The native set_properties call sets title="Test", creator="Claude",
    # subject="W4B smoke". Anchor the title assertion to its element bounds
    # (>Test<) so the bare 4-letter literal can't collide with namespace URIs
    # or tag names in core.xml.
    assert ">Test<" in core_xml, f"native core.xml missing title 'Test'; core={core_xml!r}"
    assert "Claude" in core_xml, f"native core.xml missing creator 'Claude'; core={core_xml!r}"
    assert "W4B smoke" in core_xml, (
        f"native core.xml missing subject 'W4B smoke'; core={core_xml!r}"
    )


@pytest.mark.smoke
def test_rich_features_print_area_roundtrip(tmp_path: Path) -> None:
    """Native set_print_area must produce a Print_Area defined name in workbook.xml."""
    native_path = tmp_path / "native.xlsx"
    _save_under("native", native_path)

    with zipfile.ZipFile(native_path) as nz:
        wb_xml = nz.read("xl/workbook.xml").decode("utf-8")

    assert "_xlnm.Print_Area" in wb_xml, (
        f"native workbook.xml missing _xlnm.Print_Area; workbook={wb_xml!r}"
    )
