from __future__ import annotations

import importlib.util
import json
import sys
import zipfile
from pathlib import Path
from types import ModuleType


def _load_interactive_module() -> ModuleType:
    script = (
        Path(__file__).resolve().parents[1]
        / "scripts"
        / "audit_ooxml_interactive_evidence.py"
    )
    spec = importlib.util.spec_from_file_location(
        "audit_ooxml_interactive_evidence", script
    )
    assert spec is not None
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


interactive = _load_interactive_module()


def test_interactive_audit_marks_absent_probes_not_applicable(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    _write_plain_workbook(fixture_dir / "plain.xlsx")
    _write_manifest(fixture_dir, "plain.xlsx")

    report = interactive.audit_interactive_evidence(fixture_dir)

    assert report["ready"] is True
    assert report["probes"]["slicer_selection_state"]["status"] == "not_applicable"
    assert report["probes"]["macro_project_presence"]["status"] == "not_applicable"


def test_interactive_audit_requires_probe_for_applicable_feature(
    tmp_path: Path,
) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    _write_vba_workbook(fixture_dir / "macro.xlsm")
    _write_manifest(fixture_dir, "macro.xlsm")

    report = interactive.audit_interactive_evidence(fixture_dir)

    assert report["ready"] is False
    macro = report["probes"]["macro_project_presence"]
    assert macro["status"] == "missing"
    assert macro["candidate_fixtures"] == ["macro.xlsm"]
    assert macro["missing"] == ["interactive_probe_pass"]


def test_interactive_audit_accepts_passing_probe_report(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    _write_vba_workbook(fixture_dir / "macro.xlsm")
    _write_manifest(fixture_dir, "macro.xlsm")
    probe_report = tmp_path / "interactive-report.json"
    probe_report.write_text(
        json.dumps(
            {
                "results": [
                    {
                        "fixture": "macro.xlsm",
                        "probe": "macro_project_presence",
                        "status": "passed",
                    }
                ]
            }
        )
    )

    report = interactive.audit_interactive_evidence(
        fixture_dir,
        reports=[probe_report],
    )

    assert report["ready"] is True
    macro = report["probes"]["macro_project_presence"]
    assert macro["status"] == "clear"
    assert macro["passed_fixtures"] == ["macro.xlsm"]


def test_interactive_strict_cli_fails_when_probe_missing(
    tmp_path: Path, capsys
) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    _write_vba_workbook(fixture_dir / "macro.xlsm")
    _write_manifest(fixture_dir, "macro.xlsm")

    code = interactive.main([str(fixture_dir), "--strict"])

    captured = capsys.readouterr()
    assert code == 1
    payload = json.loads(captured.out)
    assert payload["ready"] is False


def _write_manifest(fixture_dir: Path, filename: str) -> None:
    fixture_dir.joinpath("manifest.json").write_text(
        json.dumps(
            {
                "fixtures": [
                    {
                        "filename": filename,
                        "fixture_id": Path(filename).stem,
                        "tool": "excel",
                    }
                ]
            }
        )
    )


def _write_plain_workbook(path: Path) -> None:
    entries = _base_entries()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)


def _write_vba_workbook(path: Path) -> None:
    entries = _base_entries()
    entries["xl/vbaProject.bin"] = b"vba-project"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)


def _base_entries() -> dict[str, str | bytes]:
    return {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
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
    }
