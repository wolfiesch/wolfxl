from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

import openpyxl


def _load_runner_module() -> ModuleType:
    script = (
        Path(__file__).resolve().parents[1]
        / "scripts"
        / "run_ooxml_fidelity_mutations.py"
    )
    spec = importlib.util.spec_from_file_location("run_ooxml_fidelity_mutations", script)
    assert spec is not None
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


runner_module = _load_runner_module()


def _make_fixture(path: Path) -> None:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"] = "revenue"
    worksheet["A2"] = 100
    workbook.save(path)


def test_runner_writes_report_for_safe_mutations(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    fixture = fixture_dir / "simple.xlsx"
    _make_fixture(fixture)

    report = runner_module.run_sweep(fixture_dir, output_dir)

    assert report["result_count"] == 5
    assert report["failure_count"] == 0
    assert (output_dir / "report.json").is_file()
    statuses = {result["mutation"]: result["status"] for result in report["results"]}
    assert statuses == {
        "no_op": "passed",
        "marker_cell": "passed",
        "style_cell": "passed",
        "insert_tail_row": "passed",
        "move_marker_range": "passed",
    }


def test_runner_reports_manifest_hash_mismatch(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    fixture = fixture_dir / "simple.xlsx"
    _make_fixture(fixture)
    (fixture_dir / "manifest.json").write_text(
        """
{
  "fixtures": [
    {"filename": "simple.xlsx", "sha256": "not-the-real-hash"}
  ]
}
""".strip()
    )

    report = runner_module.run_sweep(
        fixture_dir,
        output_dir,
        mutations=("no_op",),
    )

    assert report["failure_count"] == 1
    assert report["results"][0]["status"] == "hash_mismatch"


def test_runner_separates_expected_rename_drift(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    fixture = fixture_dir / "simple.xlsx"
    _make_fixture(fixture)

    def fake_audit(_before: Path, _after: Path) -> dict:
        return {
            "issues": [
                {
                    "kind": "charts_semantic_drift",
                    "severity": "error",
                    "part": "charts",
                    "message": "expected formula change after sheet rename",
                }
            ]
        }

    monkeypatch.setattr(runner_module.audit_ooxml_fidelity, "audit", fake_audit)

    report = runner_module.run_sweep(
        fixture_dir,
        output_dir,
        mutations=("rename_first_sheet",),
    )

    assert report["failure_count"] == 0
    result = report["results"][0]
    assert result["status"] == "passed_with_expected_drift"
    assert result["issue_count"] == 0
    assert result["expected_issue_count"] == 1
    assert result["expected_issues"][0]["kind"] == "charts_semantic_drift"
