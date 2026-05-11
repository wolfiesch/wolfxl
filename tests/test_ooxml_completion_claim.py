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
    assert "below the customer-scale target" in corpus_requirement["reason"]
    assert corpus_requirement["evidence"] == {
        "actual_workbook_count": 11,
        "actual_source_count": 3,
        "customer_scale_min_workbooks": completion.CUSTOMER_SCALE_MIN_WORKBOOKS,
        "customer_scale_min_sources": completion.CUSTOMER_SCALE_MIN_SOURCES,
        "workbook_deficit": completion.CUSTOMER_SCALE_MIN_WORKBOOKS - 11,
        "source_deficit": completion.CUSTOMER_SCALE_MIN_SOURCES - 3,
    }
    render_requirement = next(
        requirement
        for requirement in report["missing_requirements"]
        if requirement["id"] == "feature_specific_intentional_render_equivalence"
    )
    assert render_requirement["evidence"]["ready_report_count"] > 0
    assert render_requirement["evidence"]["excel_report_count"] > 0
    assert render_requirement["evidence"]["passed_count"] > 0
    assert render_requirement["evidence"]["failure_count"] == 0
    assert render_requirement["evidence"]["target_status"] == (
        "open_unbounded_high_risk_feature_edit_universe"
    )
    interaction_requirement = next(
        requirement
        for requirement in report["missing_requirements"]
        if requirement["id"] == "broader_click_level_interaction_variants"
    )
    assert interaction_requirement["evidence"]["probe_report_count"] > 0
    assert interaction_requirement["evidence"]["raw_result_count"] > 0
    assert interaction_requirement["evidence"]["known_boundary_failure_count"] == 2
    assert interaction_requirement["evidence"]["diagnostic_non_state_change_failure_count"] == 0
    assert interaction_requirement["evidence"]["unresolved_non_boundary_failure_count"] == 0
    assert interaction_requirement["evidence"]["non_boundary_failure_count"] == 0
    assert interaction_requirement["evidence"]["failed_raw_reports"] == sorted(
        completion.KNOWN_UI_INTERACTION_BOUNDARY_REPORTS
    )
    assert interaction_requirement["evidence"]["known_boundary_reports"] == sorted(
        completion.KNOWN_UI_INTERACTION_BOUNDARY_REPORTS
    )
    assert interaction_requirement["evidence"]["diagnostic_non_state_change_reports"] == []
    assert interaction_requirement["evidence"]["unresolved_failed_raw_reports"] == []
    assert interaction_requirement["evidence"]["target_status"] == (
        "open_unbounded_click_level_variant_universe"
    )
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
    assert "synthgl_real_world_ingestion_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "synthgl_ingest_confidence_sample_100_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "synthgl_ingest_confidence_sample_100_gap_radar" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "synthgl_ingest_confidence_sample_100_quick_mutation_report" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "synthgl_codex_spark_archive_71_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "synthgl_codex_spark_archive_71_gap_radar" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "sec_commission_votes_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "sec_commission_votes_gap_radar" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "sec_commission_votes_quick_mutation_report" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "sec_enforcement_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "sec_enforcement_gap_radar" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "sec_enforcement_quick_mutation_report" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "domain_consolidation_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "domain_consolidation_gap_radar" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "domain_consolidation_quick_mutation_report" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "domain_international_standards_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "domain_international_standards_gap_radar" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "domain_international_standards_quick_mutation_report" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "fed_asset_pricing_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "fed_asset_pricing_gap_radar" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "fed_asset_pricing_quick_mutation_report" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "fdic_qbp_timeseries_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "fdic_qbp_timeseries_gap_radar" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "fdic_qbp_timeseries_quick_mutation_report" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "umya_result_files_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "calamine_reader_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "domain_ground_truth_valid_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "codexaudit_qoe_workbooks_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "fed_aea_papers_corpus_buckets" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "synthgl_real_world_ingestion_gap_radar" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "powerpivot_contoso_sidecar_gap_radar" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "domain_ground_truth_valid_gap_radar" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "codexaudit_qoe_workbooks_gap_radar" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "fed_aea_papers_gap_radar" in (
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
    assert "fintech_hackathon_demo_neutral_render_equivalence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "fed_aea_papers_neutral_render_smoke" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "fed_aea_papers_copy_chart_render_equivalence" in (
        completion.REQUIRED_CURRENT_EVIDENCE_REPORTS
    )
    assert "codexaudit_qoe_sample_add_dv_render_equivalence" in (
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


def test_completion_claim_marks_corpus_requirement_satisfied_at_customer_scale(
    tmp_path: Path,
) -> None:
    manifest = _write_bundle_manifest(
        tmp_path,
        ready=True,
        include_required_reports=True,
        corpus_source_count=completion.CUSTOMER_SCALE_MIN_SOURCES,
        corpus_workbook_count=completion.CUSTOMER_SCALE_MIN_WORKBOOKS,
    )

    report = completion.audit_completion_claim(manifest)

    corpus_requirement = next(
        requirement
        for requirement in report["criteria"]
        if requirement["id"] == "broader_real_world_corpus_diversity"
    )
    assert corpus_requirement["status"] == "satisfied"
    assert "satisfies the customer-scale corpus target" in corpus_requirement["reason"]
    assert "remains below" not in corpus_requirement["reason"]
    assert corpus_requirement["evidence"]["workbook_deficit"] == 0
    assert corpus_requirement["evidence"]["source_deficit"] == 0
    assert "broader_real_world_corpus_diversity" not in report["missing_requirement_ids"]


def test_completion_claim_classifies_paired_ui_diagnostic_non_state_change_rechecks(
    tmp_path: Path,
) -> None:
    manifest = _write_bundle_manifest(
        tmp_path,
        ready=True,
        include_required_reports=True,
        include_diagnostic_non_state_change_reports=True,
    )

    report = completion.audit_completion_claim(manifest)

    interaction_requirement = next(
        requirement
        for requirement in report["missing_requirements"]
        if requirement["id"] == "broader_click_level_interaction_variants"
    )
    evidence = interaction_requirement["evidence"]
    expected_diagnostics = sorted(
        completion.KNOWN_UI_INTERACTION_DIAGNOSTIC_NON_STATE_CHANGE_REPORTS
    )
    assert evidence["raw_failure_count"] == 5
    assert evidence["known_boundary_failure_count"] == 2
    assert evidence["observed_non_boundary_failure_count"] == 3
    assert evidence["diagnostic_non_state_change_failure_count"] == 3
    assert evidence["diagnostic_non_state_change_reports"] == expected_diagnostics
    assert {
        pairing["diagnostic_report"]
        for pairing in evidence["diagnostic_non_state_change_pairings"]
    } == set(expected_diagnostics)
    assert evidence["unresolved_non_boundary_failure_count"] == 0
    assert evidence["non_boundary_failure_count"] == 0
    assert evidence["unresolved_failed_raw_reports"] == []
    assert evidence["failed_raw_reports"] == sorted(
        [
            *completion.KNOWN_UI_INTERACTION_BOUNDARY_REPORTS,
            *completion.KNOWN_UI_INTERACTION_DIAGNOSTIC_NON_STATE_CHANGE_REPORTS,
        ]
    )


def test_completion_claim_counts_unexpected_ui_failures_from_failed_expectations(
    tmp_path: Path,
) -> None:
    manifest = _write_bundle_manifest(
        tmp_path,
        ready=True,
        include_required_reports=True,
        include_unexpected_ui_failure_report=True,
    )

    report = completion.audit_completion_claim(manifest)

    interaction_requirement = next(
        requirement
        for requirement in report["missing_requirements"]
        if requirement["id"] == "broader_click_level_interaction_variants"
    )
    evidence = interaction_requirement["evidence"]
    assert report["bundle_audit"]["ready"] is False
    assert evidence["raw_failure_count"] == 3
    assert evidence["known_boundary_failure_count"] == 2
    assert evidence["unresolved_non_boundary_failure_count"] == 1
    assert evidence["unresolved_failed_raw_reports"] == [
        "excel_ui_interaction_unexpected_failure_probe"
    ]
    assert "excel_ui_interaction_unexpected_failure_probe" in evidence["failed_raw_reports"]


def _write_bundle_manifest(
    tmp_path: Path,
    *,
    ready: bool,
    include_required_reports: bool,
    include_diagnostic_non_state_change_reports: bool = False,
    include_unexpected_ui_failure_report: bool = False,
    corpus_source_count: int = 3,
    corpus_workbook_count: int = 11,
) -> Path:
    tmp_path.mkdir(parents=True, exist_ok=True)
    reports = []
    names = ["current"]
    if include_required_reports:
        names.extend(completion.REQUIRED_CURRENT_EVIDENCE_REPORTS)
    if include_diagnostic_non_state_change_reports:
        names.extend(completion.KNOWN_UI_INTERACTION_DIAGNOSTIC_NON_STATE_CHANGE_REPORTS)
        names.extend(
            completion.KNOWN_UI_INTERACTION_DIAGNOSTIC_NON_STATE_CHANGE_REPORTS.values()
        )
    if include_unexpected_ui_failure_report:
        names.append("excel_ui_interaction_unexpected_failure_probe")
    for index, name in enumerate(names):
        report_path = tmp_path / f"report-{index}.json"
        payload = {"ready": ready}
        expect = [{"path": "ready", "equals": True}]
        if name == "corpus_portfolio_diversity":
            payload.update(
                {
                    "source_count": corpus_source_count,
                    "workbook_count": corpus_workbook_count,
                }
            )
            expect.extend(
                [
                    {"path": "source_count", "equals": corpus_source_count},
                    {"path": "workbook_count", "equals": corpus_workbook_count},
                ]
            )
        if "render_equivalence" in name:
            payload.update(
                {
                    "render_engine": "excel",
                    "observed_mutations": ["add_data_validation"],
                    "result_count": 2,
                    "passed_count": 2,
                    "failure_count": 0,
                    "inconclusive_count": 0,
                    "skipped_count": 0,
                }
            )
            expect.extend(
                [
                    {"path": "render_engine", "equals": "excel"},
                    {"path": "observed_mutations", "equals": ["add_data_validation"]},
                    {"path": "result_count", "equals": 2},
                    {"path": "passed_count", "equals": 2},
                    {"path": "failure_count", "equals": 0},
                    {"path": "inconclusive_count", "equals": 0},
                    {"path": "skipped_count", "equals": 0},
                ]
            )
        if name.startswith("excel_ui_interaction_"):
            payload.update({"probe_kind": "excel_ui_interaction", "report_count": 1})
            expect.extend(
                [
                    {"path": "probe_kind", "equals": "excel_ui_interaction"},
                    {"path": "report_count", "equals": 1},
                ]
            )
        if name in completion.KNOWN_UI_INTERACTION_BOUNDARY_REPORTS:
            payload.update(
                {
                    "probe_kind": "excel_ui_interaction",
                    "completed": True,
                    "result_count": 11,
                    "failure_count": 1,
                }
            )
            expect.extend(
                [
                    {"path": "probe_kind", "equals": "excel_ui_interaction"},
                    {"path": "completed", "equals": True},
                    {"path": "result_count", "equals": 11},
                    {"path": "failure_count", "equals": 1},
                ]
            )
        if name in completion.KNOWN_UI_INTERACTION_DIAGNOSTIC_NON_STATE_CHANGE_REPORTS:
            payload.update(
                {
                    "probe_kind": "excel_ui_interaction",
                    "completed": True,
                    "result_count": 1,
                    "failure_count": 1,
                }
            )
            expect.extend(
                [
                    {"path": "probe_kind", "equals": "excel_ui_interaction"},
                    {"path": "completed", "equals": True},
                    {"path": "result_count", "equals": 1},
                    {"path": "failure_count", "equals": 1},
                ]
            )
        if name in completion.KNOWN_UI_INTERACTION_DIAGNOSTIC_NON_STATE_CHANGE_REPORTS.values():
            payload.update(
                {
                    "probe_kind": "excel_ui_interaction",
                    "completed": True,
                    "result_count": 1,
                    "failure_count": 0,
                }
            )
            expect.extend(
                [
                    {"path": "probe_kind", "equals": "excel_ui_interaction"},
                    {"path": "completed", "equals": True},
                    {"path": "result_count", "equals": 1},
                    {"path": "failure_count", "equals": 0},
                ]
            )
        if name == "excel_ui_interaction_unexpected_failure_probe":
            payload.update(
                {
                    "probe_kind": "excel_ui_interaction",
                    "completed": True,
                    "result_count": 1,
                    "failure_count": 1,
                }
            )
            expect.extend(
                [
                    {"path": "probe_kind", "equals": "excel_ui_interaction"},
                    {"path": "completed", "equals": True},
                    {"path": "result_count", "equals": 1},
                    {"path": "failure_count", "equals": 0},
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
