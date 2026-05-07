from __future__ import annotations

import importlib.util
import json
import sys
import zipfile
from pathlib import Path
from types import SimpleNamespace
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


def _load_probe_runner_module() -> ModuleType:
    script = (
        Path(__file__).resolve().parents[1]
        / "scripts"
        / "run_ooxml_interactive_probe.py"
    )
    spec = importlib.util.spec_from_file_location("run_ooxml_interactive_probe", script)
    assert spec is not None
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


probe_runner = _load_probe_runner_module()


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


def test_macro_probe_runner_emits_passing_interactive_report(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    _write_vba_workbook(fixture_dir / "macro.xlsm")
    _write_manifest(fixture_dir, "macro.xlsm")

    def fake_smoke_excel(src: Path, _output_dir: Path, _timeout: int):
        return SimpleNamespace(
            status="passed",
            output=str(src),
            message="opened",
        )

    monkeypatch.setattr(
        probe_runner.run_ooxml_app_smoke,
        "_smoke_excel",
        fake_smoke_excel,
    )

    report = probe_runner.run_interactive_probes(fixture_dir, output_dir)

    assert report["failure_count"] == 0
    assert report["results"][0]["fixture"] == "macro.xlsm"
    assert report["results"][0]["probe"] == "macro_project_presence"
    assert report["results"][0]["status"] == "passed"
    audit = interactive.audit_interactive_evidence(
        fixture_dir,
        reports=[output_dir / "interactive-probe-report.json"],
    )
    assert audit["probes"]["macro_project_presence"]["status"] == "clear"


def test_macro_probe_runner_fails_when_vba_project_missing(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    _write_vba_workbook(fixture_dir / "macro.xlsm")
    _write_manifest(fixture_dir, "macro.xlsm")

    def remove_vba_during_smoke(src: Path, _output_dir: Path, _timeout: int):
        _rewrite_without_vba(src)
        return SimpleNamespace(
            status="passed",
            output=str(src),
            message="opened",
        )

    monkeypatch.setattr(
        probe_runner.run_ooxml_app_smoke,
        "_smoke_excel",
        remove_vba_during_smoke,
    )

    report = probe_runner.run_interactive_probes(fixture_dir, output_dir)

    assert report["failure_count"] == 1
    assert report["results"][0]["status"] == "failed"
    assert "missing after Excel open" in report["results"][0]["message"]


def test_embedded_control_probe_runner_emits_passing_interactive_report(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    _write_embedded_control_workbook(fixture_dir / "control.xlsx")
    _write_manifest(fixture_dir, "control.xlsx")

    def fake_smoke_excel(src: Path, _output_dir: Path, _timeout: int):
        return SimpleNamespace(
            status="passed",
            output=str(src),
            message="opened",
        )

    monkeypatch.setattr(
        probe_runner.run_ooxml_app_smoke,
        "_smoke_excel",
        fake_smoke_excel,
    )

    report = probe_runner.run_interactive_probes(
        fixture_dir,
        output_dir,
        probes=("embedded_control_openability",),
    )

    assert report["failure_count"] == 0
    assert report["results"][0]["fixture"] == "control.xlsx"
    assert report["results"][0]["probe"] == "embedded_control_openability"
    assert report["results"][0]["status"] == "passed"
    audit = interactive.audit_interactive_evidence(
        fixture_dir,
        reports=[output_dir / "interactive-probe-report.json"],
    )
    assert audit["probes"]["embedded_control_openability"]["status"] == "clear"


def test_embedded_control_probe_runner_fails_when_control_part_missing(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    _write_embedded_control_workbook(fixture_dir / "control.xlsx")
    _write_manifest(fixture_dir, "control.xlsx")

    def remove_control_during_smoke(src: Path, _output_dir: Path, _timeout: int):
        _rewrite_without_prefixes(src, ("xl/ctrlProps/",))
        return SimpleNamespace(
            status="passed",
            output=str(src),
            message="opened",
        )

    monkeypatch.setattr(
        probe_runner.run_ooxml_app_smoke,
        "_smoke_excel",
        remove_control_during_smoke,
    )

    report = probe_runner.run_interactive_probes(
        fixture_dir,
        output_dir,
        probes=("embedded_control_openability",),
    )

    assert report["failure_count"] == 1
    assert report["results"][0]["status"] == "failed"
    assert "missing after Excel open" in report["results"][0]["message"]


def test_external_link_probe_runner_emits_passing_interactive_report(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    _write_external_link_workbook(fixture_dir / "external-link.xlsx")
    _write_manifest(fixture_dir, "external-link.xlsx")

    def fake_smoke_excel(src: Path, _output_dir: Path, _timeout: int):
        return SimpleNamespace(
            status="passed",
            output=str(src),
            message="opened",
        )

    monkeypatch.setattr(
        probe_runner.run_ooxml_app_smoke,
        "_smoke_excel",
        fake_smoke_excel,
    )

    report = probe_runner.run_interactive_probes(
        fixture_dir,
        output_dir,
        probes=("external_link_update_prompt",),
    )

    assert report["failure_count"] == 0
    assert report["results"][0]["fixture"] == "external-link.xlsx"
    assert report["results"][0]["probe"] == "external_link_update_prompt"
    assert report["results"][0]["status"] == "passed"
    audit = interactive.audit_interactive_evidence(
        fixture_dir,
        reports=[output_dir / "interactive-probe-report.json"],
    )
    assert audit["probes"]["external_link_update_prompt"]["status"] == "clear"


def test_external_link_probe_runner_fails_when_link_part_missing(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    _write_external_link_workbook(fixture_dir / "external-link.xlsx")
    _write_manifest(fixture_dir, "external-link.xlsx")

    def remove_external_link_during_smoke(
        src: Path, _output_dir: Path, _timeout: int
    ):
        _rewrite_without_prefixes(src, ("xl/externalLinks/",))
        return SimpleNamespace(
            status="passed",
            output=str(src),
            message="opened",
        )

    monkeypatch.setattr(
        probe_runner.run_ooxml_app_smoke,
        "_smoke_excel",
        remove_external_link_during_smoke,
    )

    report = probe_runner.run_interactive_probes(
        fixture_dir,
        output_dir,
        probes=("external_link_update_prompt",),
    )

    assert report["failure_count"] == 1
    assert report["results"][0]["status"] == "failed"
    assert "missing after Excel open" in report["results"][0]["message"]


def test_pivot_probe_runner_emits_passing_interactive_report(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    _write_pivot_workbook(fixture_dir / "pivot.xlsx")
    _write_manifest(fixture_dir, "pivot.xlsx")

    def fake_smoke_excel(src: Path, _output_dir: Path, _timeout: int):
        return SimpleNamespace(
            status="passed",
            output=str(src),
            message="opened",
        )

    monkeypatch.setattr(
        probe_runner.run_ooxml_app_smoke,
        "_smoke_excel",
        fake_smoke_excel,
    )

    report = probe_runner.run_interactive_probes(
        fixture_dir,
        output_dir,
        probes=("pivot_refresh_state",),
    )

    assert report["failure_count"] == 0
    assert report["results"][0]["fixture"] == "pivot.xlsx"
    assert report["results"][0]["probe"] == "pivot_refresh_state"
    assert report["results"][0]["status"] == "passed"
    audit = interactive.audit_interactive_evidence(
        fixture_dir,
        reports=[output_dir / "interactive-probe-report.json"],
    )
    assert audit["probes"]["pivot_refresh_state"]["status"] == "clear"


def test_pivot_probe_runner_fails_when_pivot_part_missing(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    _write_pivot_workbook(fixture_dir / "pivot.xlsx")
    _write_manifest(fixture_dir, "pivot.xlsx")

    def remove_pivot_during_smoke(src: Path, _output_dir: Path, _timeout: int):
        _rewrite_without_prefixes(src, ("xl/pivotCache/", "xl/pivotTables/"))
        return SimpleNamespace(
            status="passed",
            output=str(src),
            message="opened",
        )

    monkeypatch.setattr(
        probe_runner.run_ooxml_app_smoke,
        "_smoke_excel",
        remove_pivot_during_smoke,
    )

    report = probe_runner.run_interactive_probes(
        fixture_dir,
        output_dir,
        probes=("pivot_refresh_state",),
    )

    assert report["failure_count"] == 1
    assert report["results"][0]["status"] == "failed"
    assert "missing after Excel open" in report["results"][0]["message"]


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


def _write_embedded_control_workbook(path: Path) -> None:
    entries = _base_entries()
    entries["xl/ctrlProps/ctrlProp1.xml"] = """<?xml version="1.0" encoding="UTF-8"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Button"/>"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)


def _write_external_link_workbook(path: Path) -> None:
    entries = _base_entries()
    entries["xl/_rels/workbook.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>"""
    entries["xl/externalLinks/externalLink1.xml"] = """<?xml version="1.0" encoding="UTF-8"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <externalBook><sheetNames><sheetName val="Sheet1"/></sheetNames></externalBook>
</externalLink>"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)


def _write_pivot_workbook(path: Path) -> None:
    entries = _base_entries()
    entries["xl/pivotCache/pivotCacheDefinition1.xml"] = """<?xml version="1.0" encoding="UTF-8"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      refreshOnLoad="1">
  <cacheSource type="worksheet"><worksheetSource ref="A1:B4" sheet="Sheet1"/></cacheSource>
  <cacheFields count="2"><cacheField name="Account"/><cacheField name="Amount"/></cacheFields>
</pivotCacheDefinition>"""
    entries["xl/pivotTables/pivotTable1.xml"] = """<?xml version="1.0" encoding="UTF-8"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      name="PivotTable1" cacheId="1">
  <location ref="A3:B6" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <rowFields count="1"><field x="0"/></rowFields>
</pivotTableDefinition>"""
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


def _rewrite_without_vba(path: Path) -> None:
    _rewrite_without_prefixes(path, ("xl/vbaProject.bin",))


def _rewrite_without_prefixes(path: Path, prefixes: tuple[str, ...]) -> None:
    with zipfile.ZipFile(path) as archive:
        entries = {
            name: archive.read(name)
            for name in archive.namelist()
            if not any(name.startswith(prefix) for prefix in prefixes)
        }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)
