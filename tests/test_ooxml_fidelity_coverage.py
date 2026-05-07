from __future__ import annotations

import importlib.util
import json
import sys
import zipfile
from pathlib import Path
from types import ModuleType


def _load_coverage_module() -> ModuleType:
    script = (
        Path(__file__).resolve().parents[1]
        / "scripts"
        / "audit_ooxml_fidelity_coverage.py"
    )
    spec = importlib.util.spec_from_file_location("audit_ooxml_fidelity_coverage", script)
    assert spec is not None
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


coverage_module = _load_coverage_module()


def test_coverage_audit_reports_missing_real_excel_and_structural_evidence(
    tmp_path: Path,
) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    fixture = fixture_dir / "chart.xlsx"
    _write_chart_fixture(fixture)
    (fixture_dir / "manifest.json").write_text(
        json.dumps(
            {
                "fixtures": [
                    {
                        "filename": fixture.name,
                        "fixture_id": "chart",
                        "tool": "excelize",
                    }
                ]
            }
        )
    )

    report = coverage_module.audit_coverage(fixture_dir)

    chart = report["surfaces"]["chart_style_color_preservation"]
    assert chart["external_tool_fixtures"] == ["chart.xlsx"]
    assert chart["missing"] == ["real_excel_fixture", "structural_mutation_pass"]
    assert report["ready"] is False


def test_coverage_audit_accepts_real_excel_and_structural_evidence(
    tmp_path: Path,
) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    external = fixture_dir / "external-chart.xlsx"
    excel = fixture_dir / "excel-chart.xlsx"
    _write_chart_fixture(external)
    _write_chart_fixture(excel)
    (fixture_dir / "manifest.json").write_text(
        json.dumps(
            {
                "fixtures": [
                    {
                        "filename": external.name,
                        "fixture_id": "external_chart",
                        "tool": "excelize",
                    },
                    {
                        "filename": excel.name,
                        "fixture_id": "excel_chart",
                        "tool": "excel",
                    },
                ]
            }
        )
    )
    mutation_report = tmp_path / "mutation-report.json"
    mutation_report.write_text(
        json.dumps(
            {
                "results": [
                    {
                        "fixture": external.name,
                        "mutation": "add_remove_chart",
                        "status": "passed",
                    }
                ]
            }
        )
    )

    report = coverage_module.audit_coverage(fixture_dir, reports=[mutation_report])

    chart = report["surfaces"]["chart_style_color_preservation"]
    assert chart["external_tool_fixtures"] == ["external-chart.xlsx"]
    assert chart["real_excel_fixtures"] == ["excel-chart.xlsx"]
    assert chart["structural_mutation_fixtures"] == ["external-chart.xlsx"]
    assert chart["missing"] == []


def test_coverage_audit_does_not_count_plain_worksheet_as_cf_surface(
    tmp_path: Path,
) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    fixture = fixture_dir / "plain.xlsx"
    _write_plain_fixture(fixture)
    (fixture_dir / "manifest.json").write_text(
        json.dumps(
            {
                "fixtures": [
                    {
                        "filename": fixture.name,
                        "fixture_id": "plain",
                        "tool": "synthetic-ooxml",
                    }
                ]
            }
        )
    )

    report = coverage_module.audit_coverage(fixture_dir)

    assert report["fixtures"][0]["surfaces"] == []
    assert (
        report["surfaces"]["conditional_formatting_extension_preservation"][
            "fixtures"
        ]
        == []
    )


def _write_chart_fixture(path: Path) -> None:
    entries = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        "xl/workbook.xml": """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
        "xl/_rels/workbook.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""",
        "xl/worksheets/sheet1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>""",
        "xl/charts/chart1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:barChart/></c:plotArea></c:chart>
</c:chartSpace>""",
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)


def _write_plain_fixture(path: Path) -> None:
    entries = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        "xl/workbook.xml": """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
        "xl/_rels/workbook.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""",
        "xl/worksheets/sheet1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>""",
        "xl/styles.xml": """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>""",
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)
