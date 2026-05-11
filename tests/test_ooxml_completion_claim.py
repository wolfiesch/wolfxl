from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from types import ModuleType


def _load_completion_module() -> ModuleType:
    script = Path(__file__).resolve().parents[1] / "scripts" / "audit_ooxml_completion_claim.py"
    spec = importlib.util.spec_from_file_location("audit_ooxml_completion_claim", script)
    assert spec is not None
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


completion = _load_completion_module()


def test_completion_claim_audit_supports_current_claim_but_not_exhaustive_claim(
    tmp_path: Path,
) -> None:
    manifest = _write_bundle_manifest(tmp_path, ready=True, include_required_reports=True)

    report = completion.audit_completion_claim(manifest)

    assert report["objective"] == "no real-world Excel fidelity gaps"
    assert report["current_supported_claim_ready"] is True
    assert report["exhaustive_claim_ready"] is False
    assert report["bundle_audit"]["ready"] is True
    assert report["missing_requirement_count"] == 4
    assert report["missing_requirement_ids"] == [
        "broader_real_world_corpus_diversity",
        "feature_specific_intentional_render_equivalence",
        "broader_click_level_interaction_variants",
        "future_surface_exhaustiveness",
    ]
    assert {
        requirement["id"] for requirement in report["missing_requirements"]
    } == {
        "broader_real_world_corpus_diversity",
        "feature_specific_intentional_render_equivalence",
        "broader_click_level_interaction_variants",
        "future_surface_exhaustiveness",
    }
    corpus_requirement = next(
        requirement
        for requirement in report["missing_requirements"]
        if requirement["id"] == "broader_real_world_corpus_diversity"
    )
    assert "11 unique readable workbooks across 3 source reports" in corpus_requirement["reason"]
    required_reports = next(
        criterion
        for criterion in report["criteria"]
        if criterion["id"] == "current_evidence_required_reports_present"
    )
    assert required_reports["evidence"]["missing_reports"] == []
    assert required_reports["evidence"]["required_report_count"] == len(
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "pivot_slicer_structural_render_equivalence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "local_project_holdouts_small_neutral_render_equivalence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "current_excel_16_108_delete_first_row_broad_external_tool_slicer_boundary" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "current_excel_16_108_delete_first_col_broad_external_tool_slicer_boundary" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "irs_soi_public_quick_mutation_report" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "bea_gdp_public_quick_mutation_report" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "usda_ers_county_public_quick_mutation_report" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "eia_energy_public_neutral_render_equivalence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "census_sitc_renderable_neutral_render_equivalence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "random_corpus_holdout_50" in completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    assert "random_corpus_holdout_50_quick_mutation_report" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "random_corpus_holdout_20_render_boundary" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "random_corpus_holdout_20_renderable_18_neutral_render_smoke" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "random_corpus_holdout_20_renderable_18_neutral_render_equivalence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "random_corpus_holdout_10_smoke_mutation_report" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "random_corpus_holdout_10_add_data_validation_render_smoke" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "random_corpus_holdout_10_add_data_validation_render_equivalence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "random_corpus_holdout_10_add_conditional_formatting_render_smoke" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert (
        "random_corpus_holdout_10_add_conditional_formatting_render_equivalence"
        in completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "random_corpus_holdout_10_chart_copy_render_smoke" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "random_corpus_holdout_10_chart_copy_render_equivalence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "excel_ui_interaction_add_conditional_formatting_shared_slicer_evidence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "excel_ui_interaction_add_remove_chart_external_link_forced_prompt_evidence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "excel_ui_interaction_rename_external_oracle_pivot_slicer_timeline_evidence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "excel_ui_interaction_move_formula_range_external_oracle_evidence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "excel_ui_interaction_delete_first_row_external_oracle_core_evidence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )


def test_completion_claim_audit_requires_named_current_evidence_reports(
    tmp_path: Path,
) -> None:
    manifest = _write_bundle_manifest(tmp_path, ready=True, include_required_reports=False)

    report = completion.audit_completion_claim(manifest)

    assert report["current_supported_claim_ready"] is False
    assert report["exhaustive_claim_ready"] is False
    assert report["bundle_audit"]["ready"] is True
    required_reports = next(
        criterion
        for criterion in report["criteria"]
        if criterion["id"] == "current_evidence_required_reports_present"
    )
    assert required_reports["status"] == "missing"
    assert "combined_all_evidence_gate" in required_reports["evidence"]["missing_reports"]
    assert report["missing_requirement_count"] == 5


def test_completion_claim_audit_blocks_current_claim_when_bundle_is_stale(
    tmp_path: Path,
) -> None:
    manifest = _write_bundle_manifest(tmp_path, ready=False, include_required_reports=True)

    report = completion.audit_completion_claim(manifest)

    assert report["current_supported_claim_ready"] is False
    assert report["exhaustive_claim_ready"] is False
    assert report["bundle_audit"]["issue_count"] == (
        len(completion.REQUIRED_CURRENT_EVIDENCE_REPORTS) + 1
    )
    required_reports = next(
        criterion
        for criterion in report["criteria"]
        if criterion["id"] == "current_evidence_required_reports_present"
    )
    assert required_reports["status"] == "satisfied"
    assert required_reports["evidence"]["missing_reports"] == []
    assert report["missing_requirements"][0]["id"] == "current_evidence_bundle_ready"


def test_completion_claim_strict_current_evidence_only_checks_bundle_freshness(
    tmp_path: Path,
    capsys,
) -> None:
    ready_manifest = _write_bundle_manifest(
        tmp_path / "ready", ready=True, include_required_reports=True
    )
    stale_manifest = _write_bundle_manifest(
        tmp_path / "stale", ready=False, include_required_reports=True
    )

    ready_code = completion.main([str(ready_manifest), "--strict-current-evidence"])
    ready_payload = json.loads(capsys.readouterr().out)
    stale_code = completion.main([str(stale_manifest), "--strict-current-evidence"])
    stale_payload = json.loads(capsys.readouterr().out)

    assert ready_code == 0
    assert ready_payload["current_supported_claim_ready"] is True
    assert ready_payload["exhaustive_claim_ready"] is False
    assert stale_code == 1
    assert stale_payload["current_supported_claim_ready"] is False


def test_completion_claim_strict_claim_fails_until_open_requirements_close(
    tmp_path: Path,
    capsys,
) -> None:
    manifest = _write_bundle_manifest(tmp_path, ready=True, include_required_reports=True)

    code = completion.main([str(manifest), "--strict-claim"])

    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert code == 1
    assert payload["current_supported_claim_ready"] is True
    assert payload["exhaustive_claim_ready"] is False


def _write_bundle_manifest(
    tmp_path: Path,
    *,
    ready: bool,
    include_required_reports: bool,
) -> Path:
    tmp_path.mkdir(parents=True, exist_ok=True)
    reports = []
    names = ["current"]
    if include_required_reports:
        names.extend(completion.REQUIRED_CURRENT_EVIDENCE_REPORTS)
    for index, name in enumerate(names):
        report_path = tmp_path / f"report-{index}.json"
        payload = {"ready": ready}
        expect = [{"path": "ready", "equals": True}]
        if name == "corpus_portfolio_diversity":
            payload.update({"source_count": 3, "workbook_count": 11})
            expect.extend(
                [
                    {"path": "source_count", "equals": 3},
                    {"path": "workbook_count", "equals": 11},
                ]
            )
        report_path.write_text(json.dumps(payload))
        reports.append(
            {
                "name": name,
                "path": str(report_path),
                "producer": "uv run --no-sync python scripts/example.py --strict",
                "expect": expect,
            }
        )
    manifest = tmp_path / "bundle.json"
    manifest.write_text(json.dumps({"reports": reports}))
    return manifest
