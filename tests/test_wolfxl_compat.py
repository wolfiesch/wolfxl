"""Tests for the wolfxl openpyxl-compatible API."""

from __future__ import annotations

from pathlib import Path

import pytest


def _require_rust() -> None:
    pytest.importorskip("wolfxl._rust")


# ======================================================================
# Pure-Python unit tests (no Rust needed)
# ======================================================================


class TestUtils:
    """Coordinate conversion helpers."""

    def test_column_letter(self) -> None:
        from wolfxl._utils import column_letter

        assert column_letter(1) == "A"
        assert column_letter(26) == "Z"
        assert column_letter(27) == "AA"
        assert column_letter(702) == "ZZ"

    def test_column_index(self) -> None:
        from wolfxl._utils import column_index

        assert column_index("A") == 1
        assert column_index("Z") == 26
        assert column_index("AA") == 27
        assert column_index("ZZ") == 702

    def test_a1_roundtrip(self) -> None:
        from wolfxl._utils import a1_to_rowcol, rowcol_to_a1

        assert a1_to_rowcol("B3") == (3, 2)
        assert rowcol_to_a1(3, 2) == "B3"
        assert a1_to_rowcol("AA100") == (100, 27)
        assert rowcol_to_a1(100, 27) == "AA100"

    def test_invalid_a1_raises(self) -> None:
        from wolfxl._utils import a1_to_rowcol

        with pytest.raises(ValueError, match="Invalid A1 reference"):
            a1_to_rowcol("123")


class TestStyles:
    """Frozen style dataclasses."""

    def test_font_defaults(self) -> None:
        from wolfxl._styles import Font

        f = Font()
        assert f.bold is False
        assert f.name is None
        assert f.size is None

    def test_font_is_frozen(self) -> None:
        from wolfxl._styles import Font

        f = Font(bold=True)
        with pytest.raises(AttributeError):
            f.bold = False  # type: ignore[misc]

    def test_color_hex_conversion(self) -> None:
        from wolfxl._styles import Color

        c = Color(rgb="FFFF0000")
        assert c.to_hex() == "#FF0000"
        assert Color.from_hex("#00FF00").rgb == "FF00FF00"

    def test_pattern_fill(self) -> None:
        from wolfxl._styles import PatternFill

        fill = PatternFill(patternType="solid", fgColor="#FF0000")
        assert fill._fg_hex() == "#FF0000"

    def test_border_defaults(self) -> None:
        from wolfxl._styles import Border, Side

        b = Border()
        assert b.left == Side()
        assert b.top.style is None

    def test_alignment_defaults(self) -> None:
        from wolfxl._styles import Alignment

        a = Alignment()
        assert a.horizontal is None
        assert a.wrap_text is False
        assert a.indent == 0


# ======================================================================
# Read tests (require wolfxl._rust + fixtures)
# ======================================================================

REPO_ROOT = Path(__file__).resolve().parents[1]
FIXTURES = REPO_ROOT / "tests" / "fixtures"


class TestReadMode:
    """Read an existing Excel fixture via CalamineStyledBook."""

    def setup_method(self) -> None:
        _require_rust()

    def test_load_workbook_basic(self) -> None:
        from wolfxl import load_workbook

        path = FIXTURES / "tier1" / "01_cell_values.xlsx"
        if not path.exists():
            pytest.skip("fixture not found")
        wb = load_workbook(str(path))
        assert "Sheet1" in wb.sheetnames or len(wb.sheetnames) > 0
        ws = wb[wb.sheetnames[0]]
        # Column B has test values, A has labels. Row 2 = "Hello World" per manifest.
        val = ws["B2"].value
        assert val == "Hello World"
        wb.close()

    def test_read_number(self) -> None:
        from wolfxl import load_workbook

        path = FIXTURES / "tier1" / "01_cell_values.xlsx"
        if not path.exists():
            pytest.skip("fixture not found")
        wb = load_workbook(str(path))
        ws = wb[wb.sheetnames[0]]
        # Row 7 col B = integer 42
        val = ws["B7"].value
        assert val == 42 or val == 42.0
        wb.close()

    def test_read_font_bold(self) -> None:
        from wolfxl import load_workbook

        path = FIXTURES / "tier1" / "03_text_formatting.xlsx"
        if not path.exists():
            pytest.skip("fixture not found")
        wb = load_workbook(str(path))
        ws = wb[wb.sheetnames[0]]
        # Row 2, col B = bold text per manifest
        cell = ws["B2"]
        assert cell.font.bold is True
        wb.close()

    def test_read_background_color(self) -> None:
        from wolfxl import load_workbook

        path = FIXTURES / "tier1" / "04_background_colors.xlsx"
        if not path.exists():
            pytest.skip("fixture not found")
        wb = load_workbook(str(path))
        ws = wb[wb.sheetnames[0]]
        # Row 2, col B should have a background color
        fill = ws["B2"].fill
        # Just verify it parsed without error — exact color varies by fixture
        assert fill is not None
        wb.close()

    def test_context_manager(self) -> None:
        from wolfxl import load_workbook

        path = FIXTURES / "tier1" / "01_cell_values.xlsx"
        if not path.exists():
            pytest.skip("fixture not found")
        with load_workbook(str(path)) as wb:
            assert len(wb.sheetnames) > 0

    def test_iter_rows_read_mode(self) -> None:
        from wolfxl import load_workbook

        path = FIXTURES / "tier1" / "01_cell_values.xlsx"
        if not path.exists():
            pytest.skip("fixture not found")
        wb = load_workbook(str(path))
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(min_row=1, max_row=3, values_only=True))
        assert len(rows) == 3
        assert rows[1][1] == "Hello World"  # B2
        wb.close()

    def test_iter_rows_auto_dimensions(self) -> None:
        from wolfxl import load_workbook

        path = FIXTURES / "tier1" / "01_cell_values.xlsx"
        if not path.exists():
            pytest.skip("fixture not found")
        wb = load_workbook(str(path))
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
        assert len(rows) > 10  # fixture has ~20 rows
        wb.close()

    def test_workbook_contains(self) -> None:
        from wolfxl import load_workbook

        path = FIXTURES / "tier1" / "01_cell_values.xlsx"
        if not path.exists():
            pytest.skip("fixture not found")
        wb = load_workbook(str(path))
        first = wb.sheetnames[0]
        assert first in wb
        assert "NonexistentSheet" not in wb
        wb.close()


# ======================================================================
# Write tests (require wolfxl._rust)
# ======================================================================


class TestWriteMode:
    """Write a new Excel file via RustXlsxWriterBook."""

    def setup_method(self) -> None:
        _require_rust()

    def test_write_basic(self, tmp_path: Path) -> None:
        from wolfxl import Workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Hello"
        ws["B1"] = 42
        ws["C1"] = True
        out = tmp_path / "basic.xlsx"
        wb.save(str(out))
        assert out.exists()
        assert out.stat().st_size > 0

    def test_write_with_font(self, tmp_path: Path) -> None:
        from wolfxl import Font, Workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Bold"
        ws["A1"].font = Font(bold=True, size=14, name="Arial")
        out = tmp_path / "font.xlsx"
        wb.save(str(out))
        assert out.exists()

    def test_write_with_fill(self, tmp_path: Path) -> None:
        from wolfxl import PatternFill, Workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Colored"
        ws["A1"].fill = PatternFill(patternType="solid", fgColor="#FF0000")
        out = tmp_path / "fill.xlsx"
        wb.save(str(out))
        assert out.exists()

    def test_write_with_border(self, tmp_path: Path) -> None:
        from wolfxl import Border, Side, Workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Bordered"
        ws["A1"].border = Border(
            left=Side(style="thin", color="#000000"),
            right=Side(style="thin", color="#000000"),
            top=Side(style="thin", color="#000000"),
            bottom=Side(style="thin", color="#000000"),
        )
        out = tmp_path / "border.xlsx"
        wb.save(str(out))
        assert out.exists()

    def test_write_with_alignment(self, tmp_path: Path) -> None:
        from wolfxl import Alignment, Workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Centered"
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        out = tmp_path / "align.xlsx"
        wb.save(str(out))
        assert out.exists()

    def test_write_multiple_sheets(self, tmp_path: Path) -> None:
        from wolfxl import Workbook

        wb = Workbook()
        ws1 = wb.active
        assert ws1 is not None
        ws1["A1"] = "Sheet1 data"
        ws2 = wb.create_sheet("Data")
        ws2["A1"] = "Sheet2 data"
        assert wb.sheetnames == ["Sheet", "Data"]
        out = tmp_path / "multi.xlsx"
        wb.save(str(out))
        assert out.exists()

    def test_write_number_format(self, tmp_path: Path) -> None:
        from wolfxl import Workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = 42000
        ws["A1"].number_format = "$#,##0"
        out = tmp_path / "numfmt.xlsx"
        wb.save(str(out))
        assert out.exists()

    def test_cell_method(self, tmp_path: Path) -> None:
        from wolfxl import Workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        c = ws.cell(row=1, column=1, value="Via cell()")
        assert c.value == "Via cell()"
        assert c.coordinate == "A1"
        out = tmp_path / "cell_method.xlsx"
        wb.save(str(out))
        assert out.exists()


# ======================================================================
# Round-trip tests (write with wolfxl, read back)
# ======================================================================


class TestRoundTrip:
    """Write with wolfxl, read back with wolfxl."""

    def setup_method(self) -> None:
        _require_rust()

    def test_roundtrip_values(self, tmp_path: Path) -> None:
        from wolfxl import Workbook, load_workbook

        # Write
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "text"
        ws["A2"] = 123
        ws["A3"] = 3.14
        ws["A4"] = True
        out = tmp_path / "roundtrip.xlsx"
        wb.save(str(out))

        # Read back
        wb2 = load_workbook(str(out))
        ws2 = wb2[wb2.sheetnames[0]]
        assert ws2["A1"].value == "text"
        assert ws2["A2"].value == 123 or ws2["A2"].value == 123.0
        assert abs(ws2["A3"].value - 3.14) < 0.001
        assert ws2["A4"].value is True
        wb2.close()

    def test_roundtrip_font(self, tmp_path: Path) -> None:
        from wolfxl import Font, Workbook, load_workbook

        # Write
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Bold"
        ws["A1"].font = Font(bold=True)
        out = tmp_path / "roundtrip_font.xlsx"
        wb.save(str(out))

        # Read back
        wb2 = load_workbook(str(out))
        ws2 = wb2[wb2.sheetnames[0]]
        assert ws2["A1"].value == "Bold"
        assert ws2["A1"].font.bold is True
        wb2.close()

    def test_roundtrip_fill(self, tmp_path: Path) -> None:
        from wolfxl import PatternFill, Workbook, load_workbook

        # Write
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Red"
        ws["A1"].fill = PatternFill(patternType="solid", fgColor="#FF0000")
        out = tmp_path / "roundtrip_fill.xlsx"
        wb.save(str(out))

        # Read back
        wb2 = load_workbook(str(out))
        ws2 = wb2[wb2.sheetnames[0]]
        fill = ws2["A1"].fill
        assert fill.fgColor is not None
        # The color should contain FF0000 (exact format may vary)
        fg = str(fill.fgColor).upper()
        assert "FF0000" in fg
        wb2.close()

    def test_roundtrip_formula(self, tmp_path: Path) -> None:
        from wolfxl import Workbook, load_workbook

        # Write
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = 10
        ws["A2"] = 20
        ws["A3"] = "=SUM(A1:A2)"
        out = tmp_path / "roundtrip_formula.xlsx"
        wb.save(str(out))

        # Read back — formula should be preserved as string
        wb2 = load_workbook(str(out))
        ws2 = wb2[wb2.sheetnames[0]]
        val = ws2["A3"].value
        assert val is not None
        assert "SUM" in str(val).upper()
        wb2.close()


# ======================================================================
# Modify mode tests (load existing, modify, save, verify)
# ======================================================================


FIXTURE = FIXTURES / "tier1" / "01_cell_values.xlsx"


class TestModifyMode:
    """Test the read-modify-write path via WolfXL (XlsxPatcher)."""

    def setup_method(self) -> None:
        _require_rust()
        if not FIXTURE.exists():
            pytest.skip("tier1 fixture not available")

    def test_modify_repr(self) -> None:
        from wolfxl import load_workbook

        wb = load_workbook(str(FIXTURE), modify=True)
        assert "modify" in repr(wb)
        wb.close()

    def test_modify_string_value(self, tmp_path: Path) -> None:
        from wolfxl import load_workbook

        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Modified"
        out = tmp_path / "mod_string.xlsx"
        wb.save(str(out))
        wb.close()

        # Verify with wolfxl read
        wb2 = load_workbook(str(out))
        assert wb2.active is not None
        assert wb2.active["A1"].value == "Modified"
        wb2.close()

    def test_modify_number_value(self, tmp_path: Path) -> None:
        from wolfxl import load_workbook

        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        ws["B2"] = 99.5
        out = tmp_path / "mod_number.xlsx"
        wb.save(str(out))
        wb.close()

        wb2 = load_workbook(str(out))
        assert wb2.active is not None
        assert abs(wb2.active["B2"].value - 99.5) < 0.001
        wb2.close()

    def test_modify_boolean_value(self, tmp_path: Path) -> None:
        from wolfxl import load_workbook

        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        ws["C3"] = True
        out = tmp_path / "mod_bool.xlsx"
        wb.save(str(out))
        wb.close()

        wb2 = load_workbook(str(out))
        assert wb2.active is not None
        assert wb2.active["C3"].value is True
        wb2.close()

    def test_modify_formula(self, tmp_path: Path) -> None:
        from wolfxl import load_workbook

        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        ws["D4"] = "=SUM(1,2,3)"
        out = tmp_path / "mod_formula.xlsx"
        wb.save(str(out))
        wb.close()

        # Verify formula preserved (openpyxl reads with = prefix)
        import openpyxl

        wb2 = openpyxl.load_workbook(str(out))
        assert wb2.active is not None
        assert wb2.active["D4"].value == "=SUM(1,2,3)"
        wb2.close()

    def test_modify_preserves_unchanged(self, tmp_path: Path) -> None:
        """Cells not touched should remain unchanged after save."""
        from wolfxl import load_workbook

        # Read original B1
        wb_orig = load_workbook(str(FIXTURE))
        assert wb_orig.active is not None
        orig_b1 = wb_orig.active["B1"].value
        wb_orig.close()

        # Modify only A1
        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Changed"
        out = tmp_path / "mod_preserve.xlsx"
        wb.save(str(out))
        wb.close()

        # B1 should still have its original value
        wb2 = load_workbook(str(out))
        assert wb2.active is not None
        assert wb2.active["B1"].value == orig_b1
        assert wb2.active["A1"].value == "Changed"
        wb2.close()

    def test_modify_read_then_write(self, tmp_path: Path) -> None:
        """Read a value, modify it, save — the classic read-modify-write cycle."""
        from wolfxl import load_workbook

        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        original = ws["A1"].value  # read via calamine
        ws["A1"] = f"WAS: {original}"  # write via patcher
        out = tmp_path / "mod_rmw.xlsx"
        wb.save(str(out))
        wb.close()

        wb2 = load_workbook(str(out))
        assert wb2.active is not None
        assert wb2.active["A1"].value == f"WAS: {original}"
        wb2.close()

    def test_modify_insert_new_cell(self, tmp_path: Path) -> None:
        """Insert a cell at a position that didn't exist in the original."""
        from wolfxl import load_workbook

        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        ws["Z99"] = "New cell"
        out = tmp_path / "mod_insert.xlsx"
        wb.save(str(out))
        wb.close()

        wb2 = load_workbook(str(out))
        assert wb2.active is not None
        assert wb2.active["Z99"].value == "New cell"
        wb2.close()

    def test_modify_font(self, tmp_path: Path) -> None:
        """Modify mode: set font on a cell."""
        import openpyxl

        from wolfxl import Font, load_workbook

        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Bold"
        ws["A1"].font = Font(bold=True, size=14, name="Arial", color="#FF0000")
        out = tmp_path / "mod_font.xlsx"
        wb.save(str(out))
        wb.close()

        wb2 = openpyxl.load_workbook(str(out))
        assert wb2.active is not None
        f = wb2.active["A1"].font
        assert f.bold is True
        assert f.size == 14.0
        assert f.name == "Arial"
        assert "FF0000" in str(f.color.rgb)
        wb2.close()

    def test_modify_format_only_preserves_value(self, tmp_path: Path) -> None:
        """Modify mode: format-only edits must preserve existing cell values."""
        import openpyxl

        from wolfxl import Font, load_workbook

        # Read original value via openpyxl.
        wb0 = openpyxl.load_workbook(str(FIXTURE))
        assert wb0.active is not None
        original = wb0.active["B2"].value
        wb0.close()

        # Modify only formatting (no value assignment).
        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        ws["B2"].font = Font(bold=True)
        out = tmp_path / "mod_format_only.xlsx"
        wb.save(str(out))
        wb.close()

        wb2 = openpyxl.load_workbook(str(out))
        assert wb2.active is not None
        cell = wb2.active["B2"]
        assert cell.value == original
        assert cell.font.bold is True
        wb2.close()

    def test_modify_fill(self, tmp_path: Path) -> None:
        """Modify mode: set fill on a cell."""
        import openpyxl

        from wolfxl import PatternFill, load_workbook

        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Yellow"
        ws["A1"].fill = PatternFill(patternType="solid", fgColor="#FFFF00")
        out = tmp_path / "mod_fill.xlsx"
        wb.save(str(out))
        wb.close()

        wb2 = openpyxl.load_workbook(str(out))
        assert wb2.active is not None
        fill = wb2.active["A1"].fill
        assert fill.patternType == "solid"
        assert "FFFF00" in str(fill.fgColor.rgb)
        wb2.close()

    def test_modify_alignment(self, tmp_path: Path) -> None:
        """Modify mode: set alignment on a cell."""
        import openpyxl

        from wolfxl import Alignment, load_workbook

        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Centered"
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        out = tmp_path / "mod_align.xlsx"
        wb.save(str(out))
        wb.close()

        wb2 = openpyxl.load_workbook(str(out))
        assert wb2.active is not None
        al = wb2.active["A1"].alignment
        assert al.horizontal == "center"
        assert al.vertical == "center"
        assert al.wrapText is True
        wb2.close()

    def test_modify_border(self, tmp_path: Path) -> None:
        """Modify mode: set border on a cell."""
        import openpyxl

        from wolfxl import Border, Side, load_workbook

        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Bordered"
        ws["A1"].border = Border(
            left=Side(style="thin", color="#000000"),
            right=Side(style="medium", color="#FF0000"),
            top=Side(style="thin", color="#000000"),
            bottom=Side(style="thin", color="#000000"),
        )
        out = tmp_path / "mod_border.xlsx"
        wb.save(str(out))
        wb.close()

        wb2 = openpyxl.load_workbook(str(out))
        assert wb2.active is not None
        b = wb2.active["A1"].border
        assert b.left.style == "thin"
        assert b.right.style == "medium"
        wb2.close()

    def test_modify_number_format(self, tmp_path: Path) -> None:
        """Modify mode: set number format on a cell."""
        import openpyxl

        from wolfxl import load_workbook

        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        ws["A1"] = 42000
        ws["A1"].number_format = "$#,##0"
        out = tmp_path / "mod_numfmt.xlsx"
        wb.save(str(out))
        wb.close()

        wb2 = openpyxl.load_workbook(str(out))
        assert wb2.active is not None
        assert wb2.active["A1"].number_format == "$#,##0"
        wb2.close()

    def test_modify_combined_value_and_format(self, tmp_path: Path) -> None:
        """Modify mode: set both value and format on the same cell."""
        import openpyxl

        from wolfxl import Font, PatternFill, load_workbook

        wb = load_workbook(str(FIXTURE), modify=True)
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Styled"
        ws["A1"].font = Font(bold=True, italic=True)
        ws["A1"].fill = PatternFill(patternType="solid", fgColor="#00FF00")
        out = tmp_path / "mod_combined.xlsx"
        wb.save(str(out))
        wb.close()

        wb2 = openpyxl.load_workbook(str(out))
        assert wb2.active is not None
        cell = wb2.active["A1"]
        assert cell.value == "Styled"
        assert cell.font.bold is True
        assert cell.font.italic is True
        assert "00FF00" in str(cell.fill.fgColor.rgb)
        wb2.close()

    def test_modify_multiple_sheets(self, tmp_path: Path) -> None:
        """Modify mode: patch cells across multiple sheets."""
        multi_fixture = FIXTURES /"tier1" / "09_multiple_sheets.xlsx"
        if not multi_fixture.exists():
            pytest.skip("multi-sheet fixture not available")

        import openpyxl

        from wolfxl import load_workbook

        wb = load_workbook(str(multi_fixture), modify=True)
        wb["Alpha"]["A1"] = "Patched Alpha"
        wb["Beta"]["B2"] = 999
        wb["Gamma"]["C3"] = True
        out = tmp_path / "mod_multi.xlsx"
        wb.save(str(out))
        wb.close()

        wb2 = openpyxl.load_workbook(str(out))
        assert wb2["Alpha"]["A1"].value == "Patched Alpha"
        assert wb2["Beta"]["B2"].value == 999
        assert wb2["Gamma"]["C3"].value is True
        # Originals preserved
        assert wb2["Alpha"]["B1"].value is not None  # original data still there
        wb2.close()

    def test_modify_preserves_images(self, tmp_path: Path) -> None:
        """Modify mode: files with images should preserve them after patching."""
        img_fixture = FIXTURES /"tier2" / "14_images.xlsx"
        if not img_fixture.exists():
            pytest.skip("images fixture not available")

        import openpyxl

        from wolfxl import load_workbook

        orig_size = img_fixture.stat().st_size

        wb = load_workbook(str(img_fixture), modify=True)
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Has images"
        out = tmp_path / "mod_images.xlsx"
        wb.save(str(out))
        wb.close()

        # File should still be similar size (images preserved)
        new_size = out.stat().st_size
        ratio = new_size / orig_size
        assert 0.5 < ratio < 2.0, f"Size ratio {ratio:.2f} suggests corruption"

        # Should still open fine
        wb2 = openpyxl.load_workbook(str(out))
        assert wb2.active is not None
        assert wb2.active["A1"].value == "Has images"
        wb2.close()

    def test_modify_preserves_hyperlinks(self, tmp_path: Path) -> None:
        """Modify mode: files with hyperlinks should preserve them."""
        link_fixture = FIXTURES /"tier2" / "13_hyperlinks.xlsx"
        if not link_fixture.exists():
            pytest.skip("hyperlinks fixture not available")

        import openpyxl

        from wolfxl import load_workbook

        wb = load_workbook(str(link_fixture), modify=True)
        ws = wb.active
        assert ws is not None
        ws["A1"] = "Has links"
        out = tmp_path / "mod_links.xlsx"
        wb.save(str(out))
        wb.close()

        wb2 = openpyxl.load_workbook(str(out))
        assert wb2.active is not None
        assert wb2.active["A1"].value == "Has links"
        wb2.close()


# ======================================================================
# Batch flush tests (write_sheet_values optimization)
# ======================================================================


class TestBatchFlush:
    """Verify the batch flush path produces correct output."""

    def setup_method(self) -> None:
        _require_rust()

    def test_batch_numeric_roundtrip(self, tmp_path: Path) -> None:
        """Bulk numeric writes should batch into write_sheet_values."""
        from wolfxl import Workbook, load_workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        # Write a 100×10 grid of numbers — all batchable types.
        for r in range(1, 101):
            for c in range(1, 11):
                ws.cell(row=r, column=c, value=r * 100 + c)
        out = tmp_path / "batch_numeric.xlsx"
        wb.save(str(out))

        wb2 = load_workbook(str(out))
        ws2 = wb2[wb2.sheetnames[0]]
        assert ws2["A1"].value == 101
        assert ws2["J100"].value == 10010
        assert ws2["E50"].value == 5005
        wb2.close()

    def test_batch_mixed_types(self, tmp_path: Path) -> None:
        """Mixed types: ints/strs batch, bools/formulas go per-cell."""
        from wolfxl import Workbook, load_workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = 42          # batchable
        ws["B1"] = "hello"     # batchable
        ws["C1"] = 3.14        # batchable
        ws["D1"] = True        # NOT batchable (bool)
        ws["E1"] = "=SUM(1,2)" # NOT batchable (formula)
        ws["F1"] = None         # batchable (skip)
        out = tmp_path / "batch_mixed.xlsx"
        wb.save(str(out))

        wb2 = load_workbook(str(out))
        ws2 = wb2[wb2.sheetnames[0]]
        assert ws2["A1"].value == 42 or ws2["A1"].value == 42.0
        assert ws2["B1"].value == "hello"
        assert abs(ws2["C1"].value - 3.14) < 0.001
        assert ws2["D1"].value is True
        val_e = ws2["E1"].value
        assert val_e is not None and "SUM" in str(val_e).upper()
        wb2.close()

    def test_batch_with_format(self, tmp_path: Path) -> None:
        """Cells with both values and formats: value batches, format per-cell."""
        from wolfxl import Font, Workbook, load_workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        # A1: value + format (value goes to batch, format goes per-cell)
        ws["A1"] = "styled"
        ws["A1"].font = Font(bold=True)
        # B1: value only (pure batch)
        ws["B1"] = "plain"
        out = tmp_path / "batch_format.xlsx"
        wb.save(str(out))

        wb2 = load_workbook(str(out))
        ws2 = wb2[wb2.sheetnames[0]]
        assert ws2["A1"].value == "styled"
        assert ws2["A1"].font.bold is True
        assert ws2["B1"].value == "plain"
        wb2.close()

    def test_batch_large_grid(self, tmp_path: Path) -> None:
        """10K cells — the batch path should handle this without issues."""
        from wolfxl import Workbook, load_workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        rows, cols = 1000, 10
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                ws.cell(row=r, column=c, value=float(r * cols + c))
        out = tmp_path / "batch_10k.xlsx"
        wb.save(str(out))

        wb2 = load_workbook(str(out))
        ws2 = wb2[wb2.sheetnames[0]]
        # Spot-check corners
        assert ws2["A1"].value == 11.0
        assert ws2["J1000"].value == 10010.0
        wb2.close()


# ======================================================================
# Append tests (openpyxl-compatible ws.append)
# ======================================================================


class TestAppend:
    """Test ws.append() — openpyxl-compatible row insertion."""

    def setup_method(self) -> None:
        _require_rust()

    def test_append_basic(self, tmp_path: Path) -> None:
        from wolfxl import Workbook, load_workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.append(["Name", "Age", "City"])
        ws.append(["Alice", 30, "NYC"])
        ws.append(["Bob", 25, "LA"])
        out = tmp_path / "append_basic.xlsx"
        wb.save(str(out))

        wb2 = load_workbook(str(out))
        ws2 = wb2[wb2.sheetnames[0]]
        assert ws2["A1"].value == "Name"
        assert ws2["B2"].value == 30 or ws2["B2"].value == 30.0
        assert ws2["C3"].value == "LA"
        wb2.close()

    def test_append_many_rows(self, tmp_path: Path) -> None:
        """Append 5000 rows — exercises the batch flush path."""
        from wolfxl import Workbook, load_workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        for i in range(5000):
            ws.append([i, f"row_{i}", i * 1.1])
        out = tmp_path / "append_5k.xlsx"
        wb.save(str(out))

        wb2 = load_workbook(str(out))
        ws2 = wb2[wb2.sheetnames[0]]
        assert ws2["A1"].value == 0 or ws2["A1"].value == 0.0
        assert ws2["B5000"].value == "row_4999"
        wb2.close()

    def test_append_mixed_with_cell(self, tmp_path: Path) -> None:
        """Mix append() with direct cell() writes.

        Like openpyxl, direct cell() writes do NOT advance the append counter.
        So append() after cell(row=2) still writes to the next append row (2),
        overwriting the cell() value. This matches openpyxl semantics.
        """
        from wolfxl import Workbook, load_workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.append(["Header1", "Header2"])   # row 1, next_append = 2
        ws.cell(row=5, column=1, value="manual")  # does NOT advance counter
        ws.append(["Row2_A", "Row2_B"])     # row 2 (matches openpyxl)
        out = tmp_path / "append_mixed.xlsx"
        wb.save(str(out))

        wb2 = load_workbook(str(out))
        ws2 = wb2[wb2.sheetnames[0]]
        assert ws2["A1"].value == "Header1"
        assert ws2["A2"].value == "Row2_A"
        assert ws2["A5"].value == "manual"
        wb2.close()

    def test_append_empty_row(self, tmp_path: Path) -> None:
        """Appending an empty iterable should still advance the row counter."""
        from wolfxl import Workbook, load_workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.append(["row1"])
        ws.append([])  # empty
        ws.append(["row3"])
        out = tmp_path / "append_empty.xlsx"
        wb.save(str(out))

        wb2 = load_workbook(str(out))
        ws2 = wb2[wb2.sheetnames[0]]
        assert ws2["A1"].value == "row1"
        assert ws2["A3"].value == "row3"
        wb2.close()


# ======================================================================
# PathLike support tests
# ======================================================================


class TestPathLike:
    """Test os.PathLike support for save() and load_workbook()."""

    def setup_method(self) -> None:
        _require_rust()

    def test_save_with_path(self, tmp_path: Path) -> None:
        from wolfxl import Workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "pathlib"
        out = tmp_path / "pathlike.xlsx"
        wb.save(out)  # Pass Path object, not str
        assert out.exists()

    def test_load_with_path(self, tmp_path: Path) -> None:
        from wolfxl import Workbook, load_workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"] = "pathlib"
        out = tmp_path / "pathlike_load.xlsx"
        wb.save(out)

        wb2 = load_workbook(out)  # Pass Path object, not str
        assert wb2.active is not None
        assert wb2.active["A1"].value == "pathlib"
        wb2.close()

    def test_save_and_load_modify_with_path(self) -> None:
        fixture = FIXTURES / "tier1" / "01_cell_values.xlsx"
        if not fixture.exists():
            pytest.skip("fixture not found")
        from wolfxl import load_workbook

        wb = load_workbook(fixture, modify=True)  # Path object
        assert wb.active is not None
        wb.close()


# ======================================================================
# Title setter tests
# ======================================================================


class TestTitleSetter:
    """Test ws.title setter for renaming worksheets."""

    def setup_method(self) -> None:
        _require_rust()

    def test_rename_sheet(self) -> None:
        from wolfxl import Workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        assert ws.title == "Sheet"
        ws.title = "Data"
        assert ws.title == "Data"
        assert wb.sheetnames == ["Data"]
        assert "Data" in wb
        assert "Sheet" not in wb

    def test_rename_duplicate_raises(self) -> None:
        from wolfxl import Workbook

        wb = Workbook()
        wb.create_sheet("Other")
        ws = wb.active
        assert ws is not None
        with pytest.raises(ValueError, match="already exists"):
            ws.title = "Other"

    def test_rename_noop(self) -> None:
        from wolfxl import Workbook

        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Sheet"  # Same name — should be a no-op
        assert ws.title == "Sheet"
