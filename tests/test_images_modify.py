"""Sprint Λ Pod-β (RFC-045) — modify-mode image tests.

Counterpart to ``test_images_write.py``. Verifies that
``ws.add_image(Image(...))`` works in modify mode (load → mutate →
save) and produces a valid xlsx with all the same parts. Also
exercises the v1.5 limit: appending to a sheet that already has a
drawing part raises ``NotImplementedError``.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

import wolfxl
from wolfxl import load_workbook
from wolfxl.drawing import Image

FIXTURES = Path(__file__).parent / "fixtures" / "images"
PNG_PATH = FIXTURES / "tiny_red_dot.png"
JPG_PATH = FIXTURES / "tiny_blue_dot.jpg"


def _baseline(tmp_path: Path) -> Path:
    """Write a baseline xlsx with no images."""
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "baseline"
    out = tmp_path / "baseline.xlsx"
    wb.save(out)
    return out


def test_modify_add_first_image(tmp_path: Path) -> None:
    base = _baseline(tmp_path)
    wb = load_workbook(base, modify=True)
    wb.active.add_image(Image(PNG_PATH), "B5")
    out = tmp_path / "with_image.xlsx"
    wb.save(out)

    with zipfile.ZipFile(out) as z:
        names = set(z.namelist())
        assert "xl/drawings/drawing1.xml" in names
        assert "xl/drawings/_rels/drawing1.xml.rels" in names
        assert "xl/media/image1.png" in names

        # Original cell preserved.
        s1 = z.read("xl/worksheets/sheet1.xml").decode()
        assert "<drawing r:id=" in s1

        rels = z.read("xl/worksheets/_rels/sheet1.xml.rels").decode()
        assert "/relationships/drawing" in rels

        ct = z.read("[Content_Types].xml").decode()
        assert 'Extension="png"' in ct
        assert "/xl/drawings/drawing1.xml" in ct


def test_modify_preserves_existing_data(tmp_path: Path) -> None:
    """Adding an image must not corrupt other cells."""
    openpyxl = pytest.importorskip("openpyxl")

    base = _baseline(tmp_path)
    wb = load_workbook(base, modify=True)
    wb.active.add_image(Image(PNG_PATH), "B5")
    out = tmp_path / "with_image.xlsx"
    wb.save(out)

    wb2 = openpyxl.load_workbook(out)
    assert wb2.active["A1"].value == "baseline"
    assert len(wb2.active._images) == 1


def test_modify_multiple_images(tmp_path: Path) -> None:
    base = _baseline(tmp_path)
    wb = load_workbook(base, modify=True)
    wb.active.add_image(Image(PNG_PATH), "B2")
    wb.active.add_image(Image(JPG_PATH), "D4")
    out = tmp_path / "multi.xlsx"
    wb.save(out)

    with zipfile.ZipFile(out) as z:
        names = set(z.namelist())
        assert "xl/media/image1.png" in names
        assert "xl/media/image2.jpeg" in names
        d_rels = z.read("xl/drawings/_rels/drawing1.xml.rels").decode()
        assert d_rels.count("/relationships/image") == 2


def test_modify_append_to_existing_drawing_raises(tmp_path: Path) -> None:
    """v1.5 limit — appending to a sheet that already has a drawing part."""
    # Stage 1: create a workbook WITH an image (so it has a drawing rel).
    wb = wolfxl.Workbook()
    wb.active.add_image(Image(PNG_PATH), "B5")
    base = tmp_path / "with_drawing.xlsx"
    wb.save(base)

    # Stage 2: open in modify mode, try to add another image — should
    # raise NotImplementedError.
    wb2 = load_workbook(base, modify=True)
    wb2.active.add_image(Image(JPG_PATH), "D4")
    with pytest.raises(NotImplementedError, match="(?i)drawing"):
        wb2.save(tmp_path / "boom.xlsx")
