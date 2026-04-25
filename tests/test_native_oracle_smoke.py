"""W4A smoke gate: NativeWorkbook produces structurally-equivalent xlsx
to RustXlsxWriterBook for the cell/format/structure subset.

This is the hand-rolled gate before the diff harness ships in W4C.
Layer 4 (LibreOffice) and Layer 1 (canonical bytes) come later.
"""
from __future__ import annotations

import os
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest


def _build_fixture_script() -> str:
    return r"""
import sys
import wolfxl

wb = wolfxl.Workbook()
ws1 = wb.active
ws1.title = "Data"

# Mixed cell types
ws1.cell(row=1, column=1, value="Name")
ws1.cell(row=1, column=2, value="Count")
ws1.cell(row=2, column=1, value="apples")
ws1.cell(row=2, column=2, value=42)
ws1.cell(row=3, column=1, value="pears")
ws1.cell(row=3, column=2, value=3.5)
ws1.cell(row=4, column=1, value=True)

# Bold red font on a styled cell
from wolfxl.styles import Font
c = ws1.cell(row=5, column=1, value="styled")
c.font = Font(bold=True, color="FFFF0000")

# Sheet 2: a merge + freeze
ws2 = wb.create_sheet("Layout")
ws2.cell(row=1, column=1, value="header")
ws2.merge_cells("A1:C1")
ws2.freeze_panes = "B2"

# Sheet 3: row height + column width
ws3 = wb.create_sheet("Sizes")
ws3.cell(row=1, column=1, value="x")
ws3.row_dimensions[1].height = 30
ws3.column_dimensions["A"].width = 25

wb.save(sys.argv[1])
"""


def _save_under(env_value: str, target: Path) -> None:
    env = {**os.environ, "WOLFXL_WRITER": env_value}
    result = subprocess.run(
        [sys.executable, "-c", _build_fixture_script(), str(target)],
        env=env,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        pytest.fail(
            f"WOLFXL_WRITER={env_value} save failed:\n"
            f"stdout: {result.stdout}\nstderr: {result.stderr}"
        )


@pytest.mark.smoke
def test_native_oracle_structural_smoke(tmp_path: Path) -> None:
    """Same workbook saved twice should produce the same OOXML parts.

    Layer 1 (presence): every part the oracle emits must also exist in
    the native output. Deeper structural equivalence (cell-by-cell,
    style-table-by-style-table) lands in W4C's diff harness.
    """
    oracle_path = tmp_path / "oracle.xlsx"
    native_path = tmp_path / "native.xlsx"
    _save_under("oracle", oracle_path)
    _save_under("native", native_path)

    with zipfile.ZipFile(oracle_path) as oz, zipfile.ZipFile(native_path) as nz:
        oracle_parts = {n: oz.read(n) for n in oz.namelist()}
        native_parts = {n: nz.read(n) for n in nz.namelist()}

    # Key parts must be present in both.
    for required in [
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/sheet2.xml",
        "xl/worksheets/sheet3.xml",
        "xl/styles.xml",
        "xl/sharedStrings.xml",
        "xl/workbook.xml",
    ]:
        assert required in oracle_parts, f"oracle missing {required}"
        assert required in native_parts, f"native missing {required}"

    # Workbook XML must list 3 sheets in both.
    for parts, label in [(oracle_parts, "oracle"), (native_parts, "native")]:
        wb_xml = parts["xl/workbook.xml"].decode("utf-8", errors="replace")
        for name in ("Data", "Layout", "Sizes"):
            assert name in wb_xml, f"{label} workbook.xml missing sheet {name!r}"
