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


def test_detects_chart_formula_semantic_drift(tmp_path: Path) -> None:
    before = tmp_path / "before.xlsx"
    after = tmp_path / "after.xlsx"
    before_entries = _base_entries()
    before_entries["xl/charts/chart1.xml"] = _chart_xml("Sheet1!$A$1:$A$5")
    after_entries = dict(before_entries)
    after_entries["xl/charts/chart1.xml"] = _chart_xml("Sheet1!$B$1:$B$5")

    _write_package(before, before_entries)
    _write_package(after, after_entries)

    report = audit_module.audit(before, after)

    assert any(issue["kind"] == "charts_semantic_drift" for issue in report["issues"])


def test_detects_conditional_formatting_sqref_semantic_drift(
    tmp_path: Path,
) -> None:
    before = tmp_path / "before.xlsx"
    after = tmp_path / "after.xlsx"
    before_entries = _base_entries()
    before_entries["xl/worksheets/sheet1.xml"] = _cf_xml("A1:A5")
    after_entries = dict(before_entries)
    after_entries["xl/worksheets/sheet1.xml"] = _cf_xml("A1:A4")

    _write_package(before, before_entries)
    _write_package(after, after_entries)

    report = audit_module.audit(before, after)

    assert any(
        issue["kind"] == "conditional_formatting_semantic_drift"
        for issue in report["issues"]
    )


def test_detects_external_link_target_semantic_drift(tmp_path: Path) -> None:
    before = tmp_path / "before.xlsx"
    after = tmp_path / "after.xlsx"
    before_entries = _base_entries()
    before_entries.update(_external_link_entries("../inputs/source-a.xlsx"))
    after_entries = dict(before_entries)
    after_entries.update(_external_link_entries("../inputs/source-b.xlsx"))

    _write_package(before, before_entries)
    _write_package(after, after_entries)

    report = audit_module.audit(before, after)

    assert any(
        issue["kind"] == "external_links_semantic_drift"
        for issue in report["issues"]
    )


def test_detects_pivot_and_slicer_semantic_drift(tmp_path: Path) -> None:
    before = tmp_path / "before.xlsx"
    after = tmp_path / "after.xlsx"
    before_entries = _base_entries()
    before_entries.update(_pivot_slicer_entries(cache_ref="A1:C5", slicer_cache_id="1"))
    after_entries = dict(before_entries)
    after_entries.update(_pivot_slicer_entries(cache_ref="A1:C4", slicer_cache_id="2"))

    _write_package(before, before_entries)
    _write_package(after, after_entries)

    report = audit_module.audit(before, after)
    kinds = {issue["kind"] for issue in report["issues"]}

    assert "pivots_semantic_drift" in kinds
    assert "slicers_semantic_drift" in kinds


def _chart_xml(formula: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:barChart><c:ser>
    <c:val><c:numRef><c:f>{formula}</c:f></c:numRef></c:val>
  </c:ser></c:barChart></c:plotArea></c:chart>
</c:chartSpace>"""


def _cf_xml(sqref: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <conditionalFormatting sqref="{sqref}">
    <cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule>
  </conditionalFormatting>
</worksheet>"""


def _external_link_entries(target: str) -> dict[str, str]:
    return {
        "xl/externalLinks/externalLink1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames></externalBook>
</externalLink>""",
        "xl/externalLinks/_rels/externalLink1.xml.rels": f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
    Target="{target}" TargetMode="External"/>
</Relationships>""",
    }


def _pivot_slicer_entries(cache_ref: str, slicer_cache_id: str) -> dict[str, str]:
    return {
        "xl/pivotCache/pivotCacheDefinition1.xml": f"""<?xml version="1.0" encoding="UTF-8"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  refreshOnLoad="1">
  <cacheSource type="worksheet"><worksheetSource ref="{cache_ref}" sheet="Data"/></cacheSource>
  <cacheFields count="1"><cacheField name="region"/></cacheFields>
</pivotCacheDefinition>""",
        "xl/pivotTables/pivotTable1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1" cacheId="1">
  <rowFields count="1"><field x="0"/></rowFields>
  <dataFields count="1"><dataField name="Sum of revenue" fld="0" subtotal="sum"/></dataFields>
</pivotTableDefinition>""",
        "xl/slicerCaches/slicerCache1.xml": f"""<?xml version="1.0" encoding="UTF-8"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  name="Slicer_region" pivotCacheId="{slicer_cache_id}">
  <data pivotCacheId="{slicer_cache_id}"/>
</slicerCacheDefinition>""",
        "xl/slicers/slicer1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<slicerList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicer name="Slicer_region1" cache="Slicer_region" caption="Region"/>
</slicerList>""",
    }
