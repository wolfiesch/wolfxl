from __future__ import annotations

import importlib.util
import json
import sys
import zipfile
from pathlib import Path
from types import ModuleType


def _load_gap_radar_module() -> ModuleType:
    script = Path(__file__).resolve().parents[1] / "scripts" / "audit_ooxml_gap_radar.py"
    spec = importlib.util.spec_from_file_location("audit_ooxml_gap_radar", script)
    assert spec is not None
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


gap_radar = _load_gap_radar_module()


def test_gap_radar_reports_unclassified_future_surface(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    fixture = fixture_dir / "future.xlsx"
    _write_future_fixture(fixture)
    (fixture_dir / "manifest.json").write_text(
        json.dumps(
            {
                "fixtures": [
                    {
                        "filename": fixture.name,
                        "fixture_id": "future",
                        "tool": "excel",
                    }
                ]
            }
        )
    )

    report = gap_radar.audit_gap_radar(fixture_dir)

    assert report["clear"] is False
    assert report["unknown_part_families"] == {
        "xl/future/future.xml": ["future.xlsx"]
    }
    assert report["unknown_relationship_types"] == {
        "futureFeature": ["future.xlsx"]
    }
    assert list(report["unknown_content_types"]) == [
        "application/vnd.example.future+xml"
    ]
    assert report["unknown_extension_uri_count"] == 0


def test_gap_radar_reports_unknown_extension_uri_in_known_part(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    fixture = fixture_dir / "future-ext.xlsx"
    _write_future_extension_fixture(fixture)
    (fixture_dir / "manifest.json").write_text(
        json.dumps(
            {
                "fixtures": [
                    {
                        "filename": fixture.name,
                        "fixture_id": "future_ext",
                        "tool": "excel",
                    }
                ]
            }
        )
    )

    report = gap_radar.audit_gap_radar(fixture_dir)

    assert report["clear"] is False
    assert report["unknown_part_family_count"] == 0
    assert report["unknown_relationship_type_count"] == 0
    assert report["unknown_content_type_count"] == 0
    assert report["unknown_extension_uris"] == {
        "{11111111-2222-3333-4444-555555555555}": ["future-ext.xlsx"]
    }


def test_gap_radar_is_clear_for_plain_core_workbook(tmp_path: Path) -> None:
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
                        "tool": "excel",
                    }
                ]
            }
        )
    )

    report = gap_radar.audit_gap_radar(fixture_dir)

    assert report["clear"] is True
    assert report["unknown_part_family_count"] == 0
    assert report["unknown_relationship_type_count"] == 0
    assert report["unknown_content_type_count"] == 0
    assert report["unknown_extension_uri_count"] == 0


def test_gap_radar_can_discover_nested_workbooks_without_manifest(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    nested_dir = fixture_dir / "nested"
    nested_dir.mkdir(parents=True)
    fixture = nested_dir / "plain.xlsx"
    _write_plain_fixture(fixture)

    non_recursive = gap_radar.audit_gap_radar(fixture_dir)
    recursive = gap_radar.audit_gap_radar(fixture_dir, recursive=True)

    assert non_recursive["fixture_count"] == 0
    assert recursive["fixture_count"] == 1
    assert recursive["fixtures"][0]["filename"] == "nested/plain.xlsx"
    assert recursive["clear"] is True


def test_gap_radar_classifies_python_and_sheet_metadata_surface(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    fixture = fixture_dir / "python.xlsx"
    _write_python_metadata_fixture(fixture)
    (fixture_dir / "manifest.json").write_text(
        json.dumps(
            {
                "fixtures": [
                    {
                        "filename": fixture.name,
                        "fixture_id": "python",
                        "tool": "excel",
                    }
                ]
            }
        )
    )

    report = gap_radar.audit_gap_radar(fixture_dir)

    assert report["clear"] is True
    assert report["unknown_part_family_count"] == 0
    assert report["unknown_relationship_type_count"] == 0
    assert report["unknown_content_type_count"] == 0
    assert report["unknown_extension_uri_count"] == 0


def test_gap_radar_classifies_real_world_metadata_and_extension_surfaces(
    tmp_path: Path,
) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    fixture = fixture_dir / "metadata.xlsx"
    _write_metadata_extension_fixture(fixture)
    (fixture_dir / "manifest.json").write_text(
        json.dumps(
            {
                "fixtures": [
                    {
                        "filename": fixture.name,
                        "fixture_id": "metadata",
                        "tool": "excel",
                    }
                ]
            }
        )
    )

    report = gap_radar.audit_gap_radar(fixture_dir)

    assert report["clear"] is True
    assert report["unknown_part_family_count"] == 0
    assert report["unknown_relationship_type_count"] == 0
    assert report["unknown_content_type_count"] == 0
    assert report["unknown_extension_uri_count"] == 0


def test_gap_radar_classifies_powerpivot_custom_property_surfaces(
    tmp_path: Path,
) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    fixture = fixture_dir / "powerpivot.xlsx"
    _write_powerpivot_custom_property_fixture(fixture)
    (fixture_dir / "manifest.json").write_text(
        json.dumps(
            {
                "fixtures": [
                    {
                        "filename": fixture.name,
                        "fixture_id": "powerpivot",
                        "tool": "excel",
                    }
                ]
            }
        )
    )

    report = gap_radar.audit_gap_radar(fixture_dir)

    assert report["clear"] is True
    assert report["unknown_part_family_count"] == 0
    assert report["unknown_relationship_type_count"] == 0
    assert report["unknown_content_type_count"] == 0
    assert report["unknown_extension_uri_count"] == 0


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


def _write_powerpivot_custom_property_fixture(path: Path) -> None:
    entries = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/customProperty1.bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>
</Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        "xl/workbook.xml": """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
  <extLst>
    <ext uri="{841E416B-1EF1-43b6-AB56-02D37102CBD5}"><pivotCaches/></ext>
    <ext uri="{983426D0-5260-488c-9760-48F4B6AC55F4}"><pivotTableReferences/></ext>
  </extLst>
</workbook>""",
        "xl/_rels/workbook.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>""",
        "xl/worksheets/sheet1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>""",
        "xl/worksheets/_rels/sheet1.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty" Target="../customProperty1.bin"/>
</Relationships>""",
        "xl/theme/theme1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"/>""",
        "xl/theme/_rels/theme1.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/>
</Relationships>""",
        "xl/pivotTables/pivotTable1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <extLst>
    <ext uri="{44433962-1CF7-4059-B4EE-95C3D5FFCF73}"><pivotTableData/></ext>
    <ext uri="{C510F80B-63DE-4267-81D5-13C33094786E}"><pivotTableServerFormats/></ext>
  </extLst>
</pivotTableDefinition>""",
        "xl/pivotCache/pivotCacheDefinition1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <extLst><ext uri="{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}"><pivotCacheIdVersion/></ext></extLst>
</pivotCacheDefinition>""",
        "xl/slicerCaches/slicerCache1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <extLst><ext uri="{03082B11-2C62-411c-B77F-237D8FCFBE4C}"><slicerCachePivotTables/></ext></extLst>
</slicerCacheDefinition>""",
        "xl/styles.xml": """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>""",
        "xl/customProperty1.bin": b"<Connections/>",
        "xl/media/image1.jpeg": b"jpeg",
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)


def _write_metadata_extension_fixture(path: Path) -> None:
    entries = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="wmf" ContentType="image/x-wmf"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docMetadata/LabelInfo.xml" ContentType="application/vnd.ms-office.classificationlabels+xml"/>
</Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels" Target="docMetadata/LabelInfo.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.wmf"/>
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
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme.xml"/>
  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2018/relationships/jsaProject" Target="jsaProject.bin"/>
</Relationships>""",
        "xl/worksheets/sheet1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <sheetData/>
  <extLst>
    <ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="0"/></ext>
  </extLst>
</worksheet>""",
        "xl/styles.xml": """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>""",
        "xl/theme/theme.xml": """<?xml version="1.0" encoding="UTF-8"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"/>""",
        "xl/drawings/drawing1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
  <xdr:extLst>
    <xdr:ext uri="{AF507438-7753-43E0-B8FC-AC1667EBCBE1}"><a14:hiddenEffects/></xdr:ext>
    <xdr:ext uri="{53640926-AAD7-44D8-BBD7-CCE9431645EC}"><a14:shadowObscured/></xdr:ext>
  </xdr:extLst>
</xdr:wsDr>""",
        "xl/charts/chart1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
  <c:extLst><c:ext uri="{E28EC0CA-F0BB-4C9C-879D-F8772B89E7AC}"><c16:pivotOptions16/></c:ext></c:extLst>
</c:chartSpace>""",
        "docMetadata/LabelInfo.xml": """<?xml version="1.0" encoding="UTF-8"?><LabelInfo/>""",
        "docProps/thumbnail.wmf": b"wmf",
        "xl/jsaProject.bin": b"jsa",
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("docProps/", b"")
        archive.writestr("xl/theme/", b"")
        for name, content in entries.items():
            archive.writestr(name, content)


def _write_future_fixture(path: Path) -> None:
    entries = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/future/future.xml" ContentType="application/vnd.example.future+xml"/>
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
  <Relationship Id="rId3" Type="http://schemas.example.invalid/relationships/futureFeature" Target="future/future.xml"/>
</Relationships>""",
        "xl/worksheets/sheet1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>""",
        "xl/styles.xml": """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>""",
        "xl/future/future.xml": """<?xml version="1.0" encoding="UTF-8"?>
<future xmlns="http://schemas.example.invalid/future"/>""",
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)


def _write_future_extension_fixture(path: Path) -> None:
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
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <sheetData/>
  <extLst><ext uri="{11111111-2222-3333-4444-555555555555}"><x14:futureThing/></ext></extLst>
</worksheet>""",
        "xl/styles.xml": """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>""",
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)


def _write_python_metadata_fixture(path: Path) -> None:
    entries = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/python.xml" ContentType="application/vnd.ms-excel.python+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
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
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>
  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2023/09/relationships/Python" Target="python.xml"/>
</Relationships>""",
        "xl/worksheets/sheet1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>""",
        "xl/styles.xml": """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>""",
        "xl/python.xml": """<?xml version="1.0" encoding="UTF-8"?>
<python xmlns="http://schemas.microsoft.com/office/spreadsheetml/2023/python"><environmentDefinition id="{11111111-2222-3333-4444-555555555555}"/></python>""",
        "xl/metadata.xml": """<?xml version="1.0" encoding="UTF-8"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
  <futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1"/></ext></extLst></bk></futureMetadata>
</metadata>""",
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)
