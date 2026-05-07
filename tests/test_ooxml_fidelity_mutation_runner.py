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

    assert report["result_count"] == 9
    assert report["failure_count"] == 0
    assert (output_dir / "report.json").is_file()
    statuses = {result["mutation"]: result["status"] for result in report["results"]}
    assert statuses == {
        "no_op": "passed",
        "marker_cell": "passed",
        "style_cell": "passed",
        "insert_tail_row": "passed",
        "insert_tail_col": "passed",
        "delete_marker_tail_row": "passed",
        "delete_marker_tail_col": "passed",
        "copy_remove_sheet": "passed",
        "move_marker_range": "passed",
    }


def test_runner_can_discover_recursive_fixture_trees(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    nested_dir = fixture_dir / "nested" / "deep"
    nested_dir.mkdir(parents=True)
    _make_fixture(nested_dir / "simple.xlsx")

    report = runner_module.run_sweep(
        fixture_dir,
        output_dir,
        mutations=("no_op",),
        recursive=True,
    )

    assert report["recursive"] is True
    assert report["result_count"] == 1
    result = report["results"][0]
    assert result["fixture"] == "nested/deep/simple.xlsx"
    assert result["status"] == "passed"
    assert (
        output_dir
        / "nested_deep_simple"
        / "no_op"
        / "after-simple.xlsx"
    ).is_file()


def test_runner_supports_add_remove_chart_mutation(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    fixture = fixture_dir / "simple.xlsx"
    _make_fixture(fixture)

    report = runner_module.run_sweep(
        fixture_dir,
        output_dir,
        mutations=("add_remove_chart",),
    )

    assert report["failure_count"] == 0
    assert report["results"][0]["status"] == "passed"


def test_runner_requires_formula_move_translation_drift(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    fixture = fixture_dir / "simple.xlsx"
    _make_fixture(fixture)

    report = runner_module.run_sweep(
        fixture_dir,
        output_dir,
        mutations=("move_formula_range",),
    )

    assert report["failure_count"] == 0
    result = report["results"][0]
    assert result["status"] == "passed_with_expected_drift"
    assert result["issue_count"] == 0
    assert result["expected_issue_count"] == 1
    assert result["expected_issues"][0]["kind"] == "worksheet_formulas_semantic_drift"
    assert "Z2" in result["expected_issues"][0]["message"]


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


def test_runner_separates_expected_interior_delete_drift(
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
                    "kind": "conditional_formatting_semantic_drift",
                    "severity": "error",
                    "part": "conditional_formatting",
                    "message": "expected range change after row delete",
                },
                {
                    "kind": "data_validations_semantic_drift",
                    "severity": "error",
                    "part": "data_validations",
                    "message": "expected validation range change after row delete",
                }
            ]
        }

    monkeypatch.setattr(runner_module.audit_ooxml_fidelity, "audit", fake_audit)

    report = runner_module.run_sweep(
        fixture_dir,
        output_dir,
        mutations=("delete_first_row",),
    )

    assert report["failure_count"] == 0
    result = report["results"][0]
    assert result["status"] == "passed_with_expected_drift"
    assert result["issue_count"] == 0
    assert result["expected_issue_count"] == 2
    assert {issue["kind"] for issue in result["expected_issues"]} == {
        "conditional_formatting_semantic_drift",
        "data_validations_semantic_drift",
    }


def test_runner_separates_expected_sheet_copy_drift(
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
                    "message": "expected copied chart part",
                },
                {
                    "kind": "slicers_semantic_drift",
                    "severity": "error",
                    "part": "slicers",
                    "message": "expected copied slicer part",
                },
                {
                    "kind": "timelines_semantic_drift",
                    "severity": "error",
                    "part": "timelines",
                    "message": "expected copied timeline part",
                },
                {
                    "kind": "data_validations_semantic_drift",
                    "severity": "error",
                    "part": "data_validations",
                    "message": "expected copied data validation part",
                },
            ]
        }

    monkeypatch.setattr(runner_module.audit_ooxml_fidelity, "audit", fake_audit)

    report = runner_module.run_sweep(
        fixture_dir,
        output_dir,
        mutations=("copy_first_sheet",),
    )

    assert report["failure_count"] == 0
    result = report["results"][0]
    assert result["status"] == "passed_with_expected_drift"
    assert result["issue_count"] == 0
    assert result["expected_issue_count"] == 4
    assert {issue["kind"] for issue in result["expected_issues"]} == {
        "charts_semantic_drift",
        "data_validations_semantic_drift",
        "slicers_semantic_drift",
        "timelines_semantic_drift",
    }


def test_runner_accepts_structural_external_link_formula_drift_only(
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
                    "kind": "external_links_semantic_drift",
                    "severity": "error",
                    "part": "external_links",
                    "message": "expected worksheet_formulas drift after sheet copy",
                },
                {
                    "kind": "external_links_semantic_drift",
                    "severity": "error",
                    "part": "external_links",
                    "message": "unexpected target changed from a.xlsx to b.xlsx",
                },
            ]
        }

    monkeypatch.setattr(runner_module.audit_ooxml_fidelity, "audit", fake_audit)

    report = runner_module.run_sweep(
        fixture_dir,
        output_dir,
        mutations=("copy_first_sheet",),
    )

    assert report["failure_count"] == 1
    result = report["results"][0]
    assert result["status"] == "failed"
    assert [issue["message"] for issue in result["expected_issues"]] == [
        "expected worksheet_formulas drift after sheet copy",
    ]
    assert [issue["message"] for issue in result["issues"]] == [
        "unexpected target changed from a.xlsx to b.xlsx",
    ]


def test_runner_accepts_structural_delete_calc_chain_volatility_only(
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
                    "kind": "missing_part",
                    "severity": "error",
                    "part": "xl/calcChain.xml",
                    "message": "missing part after save: xl/calcChain.xml",
                },
                {
                    "kind": "missing_relationship",
                    "severity": "error",
                    "part": "xl/_rels/workbook.xml.rels",
                    "message": (
                        "missing relationship after save: "
                        "xl/_rels/workbook.xml.rels rId6 "
                        "http://schemas.openxmlformats.org/officeDocument/2006/"
                        "relationships/calcChain -> calcChain.xml"
                    ),
                },
                {
                    "kind": "missing_part",
                    "severity": "error",
                    "part": "xl/charts/chart1.xml",
                    "message": "missing part after save: xl/charts/chart1.xml",
                },
            ]
        }

    monkeypatch.setattr(runner_module.audit_ooxml_fidelity, "audit", fake_audit)

    report = runner_module.run_sweep(
        fixture_dir,
        output_dir,
        mutations=("delete_first_row",),
    )

    assert report["failure_count"] == 1
    result = report["results"][0]
    assert result["status"] == "failed"
    assert [issue["message"] for issue in result["expected_issues"]] == [
        "missing part after save: xl/calcChain.xml",
        (
            "missing relationship after save: xl/_rels/workbook.xml.rels rId6 "
            "http://schemas.openxmlformats.org/officeDocument/2006/"
            "relationships/calcChain -> calcChain.xml"
        ),
    ]
    assert [issue["message"] for issue in result["issues"]] == [
        "missing part after save: xl/charts/chart1.xml",
    ]


def test_runner_separates_expected_feature_add_drift(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    fixture = fixture_dir / "simple.xlsx"
    _make_fixture(fixture)

    def fake_audit(_before: Path, after: Path) -> dict:
        if "add_data_validation" in str(after):
            issue = {
                "kind": "data_validations_semantic_drift",
                "severity": "error",
                "part": "data_validations",
                "message": "expected added data validation at AB2:AB10",
            }
        else:
            issue = {
                "kind": "conditional_formatting_semantic_drift",
                "severity": "error",
                "part": "conditional_formatting",
                "message": "expected added conditional format at AC2:AC10",
            }
        return {"issues": [issue]}

    monkeypatch.setattr(runner_module.audit_ooxml_fidelity, "audit", fake_audit)

    report = runner_module.run_sweep(
        fixture_dir,
        output_dir,
        mutations=("add_data_validation", "add_conditional_formatting"),
    )

    assert report["failure_count"] == 0
    statuses = {result["mutation"]: result["status"] for result in report["results"]}
    assert statuses == {
        "add_data_validation": "passed_with_expected_drift",
        "add_conditional_formatting": "passed_with_expected_drift",
    }


def test_runner_does_not_hide_feature_add_loss_drift(
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
                    "kind": "conditional_formatting_semantic_drift",
                    "severity": "error",
                    "part": "conditional_formatting",
                    "message": "before had conditional formatting after={}",
                }
            ]
        }

    monkeypatch.setattr(runner_module.audit_ooxml_fidelity, "audit", fake_audit)

    report = runner_module.run_sweep(
        fixture_dir,
        output_dir,
        mutations=("add_conditional_formatting",),
    )

    assert report["failure_count"] == 1
    assert report["results"][0]["status"] == "failed"


def test_runner_does_not_hide_missing_formula_translation_drift(
    tmp_path: Path, monkeypatch
) -> None:
    fixture_dir = tmp_path / "fixtures"
    output_dir = tmp_path / "out"
    fixture_dir.mkdir()
    fixture = fixture_dir / "simple.xlsx"
    _make_fixture(fixture)

    def fake_audit(_before: Path, _after: Path) -> dict:
        return {"issues": []}

    monkeypatch.setattr(runner_module.audit_ooxml_fidelity, "audit", fake_audit)

    report = runner_module.run_sweep(
        fixture_dir,
        output_dir,
        mutations=("move_formula_range",),
    )

    assert report["failure_count"] == 1
    result = report["results"][0]
    assert result["status"] == "failed"
    assert result["issues"][0]["kind"] == "missing_required_expected_drift"
