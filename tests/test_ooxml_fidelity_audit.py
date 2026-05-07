from __future__ import annotations

import importlib.util
import sys
import zipfile
from pathlib import Path
from types import ModuleType


def _load_audit_module() -> ModuleType:
    script = Path(__file__).resolve().parents[1] / "scripts" / "audit_ooxml_fidelity.py"
    spec = importlib.util.spec_from_file_location("audit_ooxml_fidelity", script)
    assert spec is not None
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


audit_module = _load_audit_module()


def _write_package(path: Path, entries: dict[str, str | bytes]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


def _base_entries() -> dict[str, str]:
    return {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>""",
        "xl/_rels/workbook.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>""",
        "xl/workbook.xml": """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
        "xl/worksheets/sheet1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>""",
    }


def test_clean_package_has_no_fidelity_issues(tmp_path: Path) -> None:
    before = tmp_path / "before.xlsx"
    after = tmp_path / "after.xlsx"
    entries = _base_entries()
    _write_package(before, entries)
    _write_package(after, entries)

    report = audit_module.audit(before, after)

    assert report["issues"] == []


def test_detects_dangling_chart_relationship_after_modify_save(tmp_path: Path) -> None:
    before = tmp_path / "before.xlsx"
    after = tmp_path / "after.xlsx"
    before_entries = _base_entries()
    before_entries["[Content_Types].xml"] = before_entries["[Content_Types].xml"].replace(
        "</Types>",
        '  <Override PartName="/xl/charts/chart1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>\n'
        "</Types>",
    )
    before_entries["xl/worksheets/_rels/sheet1.xml.rels"] = (
        """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    Target="../charts/chart1.xml"/>
</Relationships>"""
    )
    before_entries["xl/charts/chart1.xml"] = (
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>'
    )
    after_entries = dict(before_entries)
    del after_entries["xl/charts/chart1.xml"]

    _write_package(before, before_entries)
    _write_package(after, after_entries)

    report = audit_module.audit(before, after)
    kinds = {issue["kind"] for issue in report["issues"]}

    assert "missing_part" in kinds
    assert "dangling_relationship" in kinds
    assert "feature_part_loss" in kinds


def test_detects_conditional_formatting_dxf_reference_out_of_range(
    tmp_path: Path,
) -> None:
    before = tmp_path / "before.xlsx"
    after = tmp_path / "after.xlsx"
    entries = _base_entries()
    entries["[Content_Types].xml"] = entries["[Content_Types].xml"].replace(
        "</Types>",
        '  <Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\n'
        "</Types>",
    )
    entries["xl/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dxfs count="0"/>
</styleSheet>"""
    entries["xl/worksheets/sheet1.xml"] = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <conditionalFormatting sqref="A1:A5">
    <cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan">
      <formula>5</formula>
    </cfRule>
  </conditionalFormatting>
</worksheet>"""
    _write_package(before, entries)
    _write_package(after, entries)

    report = audit_module.audit(before, after)

    assert any(
        issue["kind"] == "conditional_formatting_dxf_out_of_range"
        for issue in report["issues"]
    )
