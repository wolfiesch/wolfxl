#!/usr/bin/env python3
"""Audit whether the OOXML fidelity evidence supports the broad completion claim."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Optional

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_ooxml_evidence_bundle  # noqa: E402

OBJECTIVE = "no real-world Excel fidelity gaps"
EXHAUSTIVE_CLAIM = "no_real_world_excel_fidelity_gaps"
CURRENT_SUPPORTED_CLAIM = (
    "no known fidelity gap in the currently pinned and classified real-world OOXML surface"
)

REQUIRED_CURRENT_EVIDENCE_REPORTS = (
    "combined_all_evidence_gate",
    "interactive_evidence_gate",
    "excel_ui_interaction_evidence_gate",
    "excel_render_full_pack_with_rename_sheet_intentional_coverage_gate",
    "excel_render_marker_cell_delta_full_pack_report",
    "excel_render_style_cell_delta_full_pack_report",
    "excel_render_insert_tail_col_delta_full_pack_report",
    "excel_render_insert_tail_row_delta_full_pack_report",
    "excel_render_delete_marker_tail_col_delta_full_pack_report",
    "excel_render_delete_marker_tail_row_delta_full_pack_report",
    "excel_render_copy_sheet_equivalence_full_pack_report",
    "excel_render_rename_sheet_equivalence_full_pack_report",
    "excel_render_retarget_external_links_equivalence_full_pack_report",
    "excel_render_add_data_validation_equivalence_full_pack_report",
    "excel_render_add_conditional_formatting_equivalence_full_pack_report",
    "excel_render_add_remove_chart_equivalence_full_pack_report",
    "excel_render_copy_remove_sheet_equivalence_full_pack_report",
    "excel_render_move_formula_range_delta_full_pack_report",
    "excel_render_move_marker_range_delta_full_pack_report",
    "excel_render_delete_first_row_delta_full_pack_report",
    "excel_render_delete_first_col_delta_full_pack_report",
    "excel_app_open_full_pack_with_cf_verified_coverage_gate",
    "excel_ui_interaction_marker_macro_evidence",
    "excel_ui_interaction_style_macro_evidence",
    "excel_ui_interaction_copy_remove_macro_evidence",
    "excel_ui_interaction_add_data_validation_macro_evidence",
    "excel_ui_interaction_mutated_pivot_refresh_evidence",
    "excel_ui_interaction_add_data_validation_pivot_refresh_evidence",
    "excel_ui_interaction_marker_table_slicer_evidence",
    "excel_ui_interaction_style_table_slicer_evidence",
    "excel_ui_interaction_copy_remove_table_slicer_evidence",
    "excel_ui_interaction_add_data_validation_table_slicer_evidence",
    "excel_ui_interaction_marker_pivot_chart_slicer_evidence",
    "excel_ui_interaction_style_pivot_chart_slicer_evidence",
    "excel_ui_interaction_copy_remove_pivot_chart_slicer_evidence",
    "excel_ui_interaction_marker_external_tool_pivot_slicer_evidence",
    "excel_ui_interaction_style_external_tool_pivot_slicer_evidence",
    "excel_ui_interaction_copy_remove_external_tool_pivot_slicer_evidence",
    "excel_ui_interaction_add_data_validation_pivot_slicer_evidence",
    "excel_ui_interaction_add_conditional_formatting_pivot_slicer_evidence",
    "excel_ui_interaction_add_data_validation_shared_slicer_evidence",
    "excel_ui_interaction_marker_timeline_evidence",
    "excel_ui_interaction_style_timeline_evidence",
    "excel_ui_interaction_copy_remove_timeline_evidence",
    "excel_ui_interaction_add_data_validation_timeline_evidence",
    "excel_ui_interaction_marker_external_link_current_prompt_evidence",
    "excel_ui_interaction_style_external_link_current_prompt_evidence",
    "excel_ui_interaction_copy_remove_external_link_current_prompt_evidence",
    "excel_ui_interaction_add_data_validation_external_link_current_prompt_evidence",
    "excel_ui_interaction_marker_umya_external_link_forced_prompt_evidence",
    "excel_ui_interaction_style_umya_external_link_forced_prompt_evidence",
    "excel_ui_interaction_copy_remove_umya_external_link_forced_prompt_evidence",
    "excel_ui_interaction_add_data_validation_umya_external_link_forced_prompt_evidence",
    "excel_ui_interaction_marker_umya_listbox_control_evidence",
    "excel_ui_interaction_style_umya_listbox_control_evidence",
    "excel_ui_interaction_copy_remove_umya_listbox_control_evidence",
    "excel_ui_interaction_add_data_validation_umya_listbox_control_evidence",
    "excel_ui_interaction_marker_control_evidence",
    "excel_ui_interaction_marker_button_control_evidence",
    "excel_ui_interaction_style_control_evidence",
    "excel_ui_interaction_style_button_control_evidence",
    "excel_ui_interaction_copy_remove_control_evidence",
    "excel_ui_interaction_copy_remove_button_control_evidence",
    "excel_ui_interaction_add_data_validation_control_evidence",
    "excel_ui_interaction_add_conditional_formatting_pivot_refresh_evidence",
    "excel_ui_interaction_add_conditional_formatting_shared_slicer_evidence",
    "excel_ui_interaction_add_conditional_formatting_control_evidence",
    "excel_ui_interaction_add_conditional_formatting_timeline_evidence",
    "excel_ui_interaction_add_remove_chart_pivot_refresh_evidence",
    "excel_ui_interaction_add_remove_chart_shared_slicer_evidence",
    "excel_ui_interaction_add_remove_chart_control_evidence",
    "excel_ui_interaction_add_remove_chart_timeline_evidence",
    "excel_ui_interaction_add_remove_chart_macro_evidence",
    "excel_ui_interaction_add_remove_chart_external_link_current_setting_evidence",
    "excel_ui_interaction_add_remove_chart_external_link_forced_prompt_evidence",
    "excel_ui_interaction_add_remove_chart_umya_external_link_forced_prompt_evidence",
    "excel_ui_interaction_rename_external_oracle_pivot_slicer_timeline_evidence",
    "excel_ui_interaction_rename_shared_slicer_evidence",
    "excel_ui_interaction_rename_external_oracle_prompt_control_evidence",
    "excel_ui_interaction_rename_powerview_evidence",
    "excel_ui_interaction_move_formula_range_external_oracle_evidence",
    "excel_ui_interaction_move_marker_range_external_oracle_evidence",
    "excel_ui_interaction_rename_first_sheet_external_oracle_evidence",
    "excel_ui_interaction_add_data_validation_external_oracle_evidence",
    "excel_ui_interaction_copy_remove_sheet_external_oracle_evidence",
    "excel_ui_interaction_add_conditional_formatting_external_oracle_evidence",
    "excel_ui_interaction_add_remove_chart_external_oracle_evidence",
    "excel_ui_interaction_insert_tail_row_external_oracle_evidence",
    "excel_ui_interaction_insert_tail_col_external_oracle_evidence",
    "excel_ui_interaction_style_cell_external_oracle_evidence",
    "excel_ui_interaction_marker_cell_external_oracle_evidence",
    "excel_ui_interaction_delete_marker_tail_row_external_oracle_evidence",
    "excel_ui_interaction_delete_marker_tail_col_external_oracle_evidence",
    "excel_ui_interaction_retarget_external_links_external_oracle_evidence",
    "excel_ui_interaction_delete_first_row_external_oracle_core_evidence",
    "excel_ui_interaction_delete_first_col_external_oracle_core_evidence",
    "external_oracle_corpus_diversity",
    "corpus_portfolio_diversity",
    "external_oracle_gap_radar",
    "sec_municipal_advisers_gap_radar",
    "sec_municipal_advisers_corpus_buckets",
    "sec_investment_mgmt_gap_radar",
    "sec_investment_mgmt_corpus_buckets",
    "irs_soi_public_corpus_buckets",
    "irs_soi_public_gap_radar",
    "irs_soi_public_quick_mutation_report",
    "bea_gdp_public_corpus_buckets",
    "bea_gdp_public_gap_radar",
    "bea_gdp_public_quick_mutation_report",
    "usda_ers_county_public_corpus_buckets",
    "usda_ers_county_public_gap_radar",
    "usda_ers_county_public_quick_mutation_report",
    "eia_energy_public_corpus_buckets",
    "eia_energy_public_gap_radar",
    "eia_energy_public_quick_mutation_report",
    "eia_energy_public_neutral_render_equivalence",
    "census_sitc_renderable_corpus_buckets",
    "census_sitc_renderable_gap_radar",
    "census_sitc_renderable_quick_mutation_report",
    "census_sitc_renderable_neutral_render_equivalence",
    "fintech_hackathon_demo_neutral_render_equivalence",
    "fed_aea_papers_neutral_render_smoke",
    "fed_aea_papers_copy_chart_render_equivalence",
    "codexaudit_qoe_sample_add_dv_render_equivalence",
    "synthgl_recursive_gap_radar",
    "umya_test_files_gap_radar",
    "umya_test_files_quick_plus_structural_mutation_coverage",
    "powerpivot_contoso_sidecar_coverage",
    "powerpivot_contoso_sidecar_excel_expected_unsupported",
    "powerpivot_contoso_sidecar_unsupported_content_prompt_evidence",
    "powerpivot_contoso_sidecar_add_data_validation_unsupported_content_prompt_evidence",
    "powerpivot_contoso_sidecar_marker_unsupported_content_prompt_evidence",
    "powerpivot_contoso_sidecar_style_unsupported_content_prompt_evidence",
    "powerpivot_contoso_sidecar_copy_remove_unsupported_content_prompt_evidence",
    "slicer_shared_two_pivots_sidecar_coverage",
    "slicer_shared_two_pivots_sidecar_copy_remove_sheet_render_equivalence",
    "slicer_shared_two_pivots_sidecar_add_conditional_formatting_render_equivalence",
    "slicer_shared_two_pivots_sidecar_neutral_feature_render_equivalence",
    "slicer_shared_two_pivots_sidecar_rename_sheet_render_equivalence",
    "timeline_slicer_neutral_feature_render_equivalence",
    "timeline_slicer_add_conditional_formatting_render_equivalence",
    "timeline_slicer_rename_sheet_render_equivalence",
    "timeline_slicer_copy_remove_sheet_render_equivalence",
    "timeline_slicer_delete_first_row_render_delta",
    "timeline_slicer_delete_first_col_render_delta",
    "slicer_shared_two_pivots_sidecar_delete_first_row_render_delta",
    "slicer_shared_two_pivots_sidecar_delete_first_col_render_delta",
    "slicer_shared_two_pivots_sidecar_interactive_evidence",
    "slicer_shared_two_pivots_sidecar_ui_interaction_evidence",
    "slicer_shared_two_pivots_sidecar_marker_ui_interaction_evidence",
    "slicer_shared_two_pivots_sidecar_style_ui_interaction_evidence",
    "slicer_shared_two_pivots_sidecar_copy_remove_ui_interaction_evidence",
    "external_link_retarget_excel_app_open",
    "ilpa_reporting_template_retarget_mutation_report",
    "ilpa_reporting_template_retarget_excel_app_open",
    "ilpa_reporting_template_retarget_render_equivalence",
    "ilpa_reporting_template_add_data_validation_render_equivalence",
    "ilpa_reporting_template_add_conditional_formatting_render_equivalence",
    "ilpa_reporting_template_add_remove_chart_render_equivalence",
    "ilpa_reporting_template_copy_remove_sheet_render_equivalence",
    "wbd_wdesk_add_data_validation_render_equivalence",
    "wbd_wdesk_add_conditional_formatting_render_equivalence",
    "wbd_wdesk_add_remove_chart_render_equivalence",
    "wbd_wdesk_copy_remove_sheet_render_equivalence",
    "bf30_remaining_feature_render_equivalence",
    "bf30_remaining_copy_remove_sheet_render_equivalence",
    "blind_holdout_feature_render_equivalence",
    "blind_holdout_copy_remove_sheet_render_equivalence",
    "rescue_downloads_feature_render_equivalence",
    "rescue_downloads_copy_remove_sheet_render_equivalence",
    "sec_edgar_add_data_validation_render_equivalence",
    "sec_edgar_add_conditional_formatting_render_equivalence",
    "sec_edgar_add_remove_chart_render_equivalence",
    "sec_edgar_copy_remove_sheet_render_equivalence",
    "iran_osint_recursive_add_data_validation_render_equivalence",
    "iran_osint_recursive_add_conditional_formatting_render_equivalence",
    "iran_osint_recursive_copy_remove_sheet_render_equivalence",
    "iran_osint_recursive_add_remove_chart_render_equivalence",
    "synthgl_docs_las_vegas_feature_render_equivalence",
    "synthgl_docs_las_vegas_copy_remove_sheet_render_equivalence",
    "excelbench_external_validated_neutral_render_equivalence",
    "spreadsheet_peek_neutral_feature_render_equivalence",
    "spreadsheet_peek_copy_remove_sheet_render_equivalence",
    "excelbench_real_world_existing_neutral_render_equivalence",
    "ticker_to_gl_strongbox_noop_marker_mutation_report",
    "ticker_to_gl_strongbox_style_copy_remove_mutation_report",
    "ticker_to_gl_strongbox_excel_app_smoke_source_marker",
    "ticker_to_gl_strongbox_excel_app_smoke_style_copy",
    "ticker_to_gl_strongbox_feature_render_equivalence",
    "ticker_to_gl_strongbox_add_conditional_formatting_render_equivalence",
    "ticker_to_gl_strongbox_copy_remove_sheet_render_equivalence",
    "sec_adviser_reports_feature_neutral_render_equivalence",
    "sec_adviser_reports_add_conditional_formatting_render_equivalence",
    "sec_investment_mgmt_feature_neutral_render_equivalence",
    "sec_investment_mgmt_add_conditional_formatting_render_equivalence",
    "sec_municipal_adviser_reports_feature_neutral_render_equivalence",
    "sec_municipal_adviser_reports_add_conditional_formatting_render_equivalence",
    "domain_ground_truth_neutral_render_equivalence",
    "pivot_slicer_structural_render_equivalence",
    "local_project_holdouts_small_neutral_render_equivalence",
    "current_excel_16_108_delete_first_row_broad_external_tool_slicer_boundary",
    "current_excel_16_108_delete_first_col_broad_external_tool_slicer_boundary",
    "random_corpus_holdout_50",
    "random_corpus_holdout_50_quick_mutation_report",
    "random_corpus_holdout_20_render_boundary",
    "random_corpus_holdout_20_renderable_18_neutral_render_smoke",
    "random_corpus_holdout_20_renderable_18_neutral_render_equivalence",
    "random_corpus_holdout_10_smoke_mutation_report",
    "random_corpus_holdout_10_add_data_validation_render_smoke",
    "random_corpus_holdout_10_add_data_validation_render_equivalence",
    "random_corpus_holdout_10_add_conditional_formatting_render_smoke",
    "random_corpus_holdout_10_add_conditional_formatting_render_equivalence",
    "random_corpus_holdout_10_chart_copy_render_smoke",
    "random_corpus_holdout_10_chart_copy_render_equivalence",
    "public_powerbi_expanded_mutations",
    "synthgl_recursive_mutation_coverage",
    "synthgl_recursive_excel_render_noop_byte_identical",
    "rename_sheet_defined_name_ref_audit_report",
)

OPEN_REQUIREMENTS = (
    {
        "id": "broader_real_world_corpus_diversity",
        "status": "open",
        "reason": (
            "The current corpus portfolio covers all required diversity buckets, "
            "and the deterministic 50-workbook random holdout now has richer "
            "mutation evidence, but it is still not customer-scale real-world "
            "Excel evidence."
        ),
    },
    {
        "id": "feature_specific_intentional_render_equivalence",
        "status": "open",
        "reason": (
            "Intentional Microsoft Excel render checks now include equivalence for "
            "copy-sheet, rename-sheet, external-link retargeting, and the visually "
            "neutral add-data-validation, add-conditional-formatting, and "
            "add-remove-chart and copy-remove-sheet edits; shared pivot-slicer "
            "sidecar evidence separately proves copy-remove-sheet and rename-sheet "
            "equivalence, add-conditional-formatting equivalence, plus "
            "add-data-validation and add-remove-chart equivalence on a high-risk "
            "shared-slicer workbook, plus add-data-validation, "
            "add-conditional-formatting, add-remove-chart, and "
            "copy-remove-sheet and rename-sheet equivalence on a high-risk "
            "timeline workbook, along with shared-slicer and timeline "
            "delete-first-row/column visual deltas; ILPA "
            "side evidence separately proves external-link "
            "retarget, add-data-validation, add-conditional-formatting, and "
            "add-remove-chart and copy-remove-sheet render "
            "equivalence on public "
            "finance templates; domain-ground-truth side evidence separately proves "
            "add-data-validation, add-conditional-formatting, add-remove-chart, and "
            "copy-remove-sheet render equivalence across 24 source-valid public "
            "domain workbooks; and WBD wDesk "
            "side evidence separately proves add-data-validation, "
            "add-conditional-formatting, add-remove-chart, and copy-remove-sheet "
            "render equivalence "
            "on an external-tool-authored "
            "public-company workbook; BF30 public-download side evidence "
            "separately proves add-data-validation, "
            "add-conditional-formatting, add-remove-chart, and "
            "copy-remove-sheet render equivalence on six readable real-world "
            "public downloads; blind-holdout side evidence "
            "separately proves add-data-validation, add-conditional-formatting, "
            "add-remove-chart, and copy-remove-sheet render equivalence on its "
            "one readable held-out workbook; rescue-downloads side evidence "
            "separately proves add-data-validation, add-conditional-formatting, "
            "add-remove-chart, and copy-remove-sheet render equivalence on two "
            "readable budget workbooks; SEC/EDGAR side evidence separately proves "
            "add-data-validation, add-conditional-formatting, and "
            "add-remove-chart and copy-remove-sheet render equivalence on six "
            "public-company "
            "workbooks; Iran OSINT side evidence separately proves "
            "add-data-validation and add-conditional-formatting render "
            "equivalence plus copy-remove-sheet and add-remove-chart render "
            "equivalence on the full 16-workbook recursive public-data corpus, "
            "and SynthGL docs QoE side evidence separately proves "
            "add-data-validation, add-conditional-formatting, and "
            "add-remove-chart plus copy-remove-sheet render equivalence on two "
            "renderable Las Vegas QoE databooks, "
            "and ExcelBench external-validated side evidence separately proves "
            "add-data-validation, add-conditional-formatting, add-remove-chart, "
            "and copy-remove-sheet render equivalence on five Excel-renderable "
            "external-tool fixtures, "
            "and Spreadsheet Peek side evidence separately proves "
            "add-data-validation, add-remove-chart, and copy-remove-sheet render "
            "equivalence on three generated finance/table examples while its "
            "add-conditional-formatting run intentionally remains outside the "
            "neutral-equivalence set because one generated wide-table example "
            "visibly changes under the new formatting rule; ExcelBench "
            "curated-manifest side evidence separately proves "
            "add-data-validation, add-conditional-formatting, add-remove-chart, "
            "and copy-remove-sheet render equivalence on seven existing "
            "Excel-renderable project-owned fixtures, while the table fixture "
            "is excluded from that neutral-render claim because Excel PDF export "
            "fails after neutral edits; ticker-to-GL "
            "Strongbox side evidence separately proves add-data-validation, "
            "add-conditional-formatting, add-remove-chart, and "
            "copy-remove-sheet render equivalence on five finance-cache "
            "workbooks under the recorded temporary print-area clamp, "
            "and SEC adviser side evidence separately proves "
            "add-data-validation, add-conditional-formatting, "
            "add-remove-chart, and copy-remove-sheet "
            "render equivalence on four public/regulatory adviser workbooks, "
            "and SEC investment-management sidecar evidence separately proves "
            "add-data-validation, add-conditional-formatting, add-remove-chart, "
            "and copy-remove-sheet render equivalence on five "
            "public/regulatory investment-management workbooks, "
            "and SEC municipal adviser sidecar evidence separately proves "
            "add-data-validation, add-conditional-formatting, add-remove-chart, "
            "and copy-remove-sheet render equivalence on all 29 "
            "public/regulatory municipal-adviser workbooks, "
            "pivot/slicer structural side evidence separately proves "
            "rename-sheet and copy-remove-sheet render equivalence on three "
            "high-risk pivot/slicer fixtures, and local-project holdout side "
            "evidence separately proves add-data-validation and "
            "copy-remove-sheet render equivalence on two representative "
            "local workbooks, and deterministic random-holdout side evidence "
            "separately proves add-data-validation, add-conditional-formatting, "
            "add-remove-chart, and copy-remove-sheet render equivalence on 10 "
            "sampled workbooks spanning eight source reports, "
            "and proves the same four neutral feature edits on an 18-workbook "
            "Excel-renderable subset of a 20-workbook random holdout spanning "
            "14 source reports, with the two Excel PDF-export boundary "
            "workbooks recorded separately, "
            "plus fintech-hackathon finance-demo neutral equivalence, Fed AEA "
            "research-data neutral render smoke and copy/chart equivalence, "
            "and CodexAudit QoE add-data-validation equivalence, "
            "plus expected "
            "visual-delta evidence for "
            "marker-cell, style-cell, insert-tail-row/column, move-marker-range, "
            "delete-marker-tail-row/column, move-formula-range, and first "
            "row/column deletion; they still do "
            "not prove semantic visual equivalence for every high-risk feature "
            "edit."
        ),
    },
    {
        "id": "broader_click_level_interaction_variants",
        "status": "open",
        "reason": (
            "Targeted UI-interaction evidence now covers source and selected "
            "marker-cell, style-cell, copy-remove-sheet, add-data-validation, "
            "add-conditional-formatting, add-remove-chart, rename-sheet, "
            "move-formula-range, move-marker-range, insert-tail-row/column, "
            "delete-marker-tail-row/column, retarget-external-links, and "
            "delete-first-row/column saves across pivot refresh, "
            "slicer/shared-slicer clicks, timeline clicks, embedded controls, "
            "macro prompts, external-link prompts, and PowerView read-only "
            "prompts. The detailed artifact list lives in "
            "Plans/real-world-excel-fidelity-gap-discovery.md. This still is "
            "not exhaustive: current Excel 16.108 rechecks show some source "
            "table-slicer, timeline, and list-box UI paths can complete actions "
            "without persisting a state change in this environment, and pinned "
            "delete-first-row/column broad UI boundary reports still expose one "
            "external-tool slicer click that is not observed after each "
            "destructive-axis edit. Broader slicer, timeline, embedded-control, "
            "and prompt variants remain unexhausted."
        ),
    },
    {
        "id": "future_surface_exhaustiveness",
        "status": "open",
        "reason": (
            "Gap radar can catch newly seen OOXML part families and extension URIs; it "
            "cannot prove that no unseen real-world Excel surface exists."
        ),
    },
)


def _check_actual(bundle_audit: dict, report_name: str, check_path: str) -> Optional[object]:
    for report in bundle_audit["reports"]:
        if report["name"] != report_name:
            continue
        for check in report["checks"]:
            if check["path"] == check_path and check["passed"]:
                return check["actual"]
    return None


def _open_requirements(bundle_audit: dict) -> list[dict]:
    requirements = [dict(requirement) for requirement in OPEN_REQUIREMENTS]
    source_count = _check_actual(bundle_audit, "corpus_portfolio_diversity", "source_count")
    workbook_count = _check_actual(bundle_audit, "corpus_portfolio_diversity", "workbook_count")
    if isinstance(source_count, int) and isinstance(workbook_count, int):
        for requirement in requirements:
            if requirement["id"] != "broader_real_world_corpus_diversity":
                continue
            requirement["reason"] = (
                f"The current corpus portfolio spans {workbook_count} unique readable "
                f"workbooks across {source_count} source reports, including the "
                "domain-ground-truth public workbook sidecar, standalone SEC "
                "municipal-adviser and SEC investment-management public/regulatory "
                "sidecars, IRS SOI, BEA GDP, USDA ERS county, EIA energy, and "
                "Census SITC public-statistics sidecars, and all required diversity "
                "buckets. The pinned deterministic random holdout now samples "
                "50 workbooks from that portfolio across 22 source reports and "
                "stages all selected files, with a smaller 10-workbook smoke "
                "mutation passing no-op and marker-cell saves; this improves "
                "curated-corpus pressure evidence, but it is still not "
                "customer-scale real-world Excel evidence."
            )
            break
    return requirements


def audit_completion_claim(bundle_path: Path) -> dict:
    bundle_audit = audit_ooxml_evidence_bundle.audit_bundle(bundle_path)
    bundle_ready = bool(bundle_audit["ready"])
    report_names = {str(report["name"]) for report in bundle_audit["reports"]}
    missing_report_names = sorted(set(REQUIRED_CURRENT_EVIDENCE_REPORTS) - report_names)
    required_reports_present = not missing_report_names
    current_supported_claim_ready = bundle_ready and required_reports_present
    criteria = [
        {
            "id": "current_evidence_bundle_ready",
            "status": "satisfied" if bundle_ready else "missing",
            "evidence": {
                "manifest": str(bundle_path),
                "ready": bundle_ready,
                "report_count": bundle_audit["report_count"],
                "producer_count": bundle_audit["producer_count"],
                "issue_count": bundle_audit["issue_count"],
            },
        },
        {
            "id": "current_evidence_required_reports_present",
            "status": "satisfied" if required_reports_present else "missing",
            "reason": (
                "The current supported claim requires the named coverage, render, "
                "app-open, interactive, corpus, gap-radar, sidecar, and broader "
                "corpus evidence reports pinned in the bundle."
            ),
            "evidence": {
                "required_report_count": len(REQUIRED_CURRENT_EVIDENCE_REPORTS),
                "present_report_count": len(
                    set(REQUIRED_CURRENT_EVIDENCE_REPORTS) & report_names
                ),
                "missing_reports": missing_report_names,
            },
        },
        *_open_requirements(bundle_audit),
    ]
    missing = [
        criterion
        for criterion in criteria
        if criterion["status"] in {"missing", "open", "partial"}
    ]
    return {
        "objective": OBJECTIVE,
        "exhaustive_claim": EXHAUSTIVE_CLAIM,
        "exhaustive_claim_ready": False,
        "current_supported_claim": CURRENT_SUPPORTED_CLAIM,
        "current_supported_claim_ready": current_supported_claim_ready,
        "criteria": criteria,
        "missing_requirement_count": len(missing),
        "missing_requirement_ids": [criterion["id"] for criterion in missing],
        "missing_requirements": missing,
        "bundle_audit": {
            "ready": bundle_ready,
            "report_count": bundle_audit["report_count"],
            "producer_count": bundle_audit["producer_count"],
            "issue_count": bundle_audit["issue_count"],
            "issues": bundle_audit["issues"],
        },
    }


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "bundle",
        type=Path,
        nargs="?",
        default=Path("Plans/ooxml-current-evidence-bundle.json"),
        help="Pinned evidence bundle manifest to audit.",
    )
    parser.add_argument(
        "--strict-current-evidence",
        action="store_true",
        help="Exit non-zero when the pinned current evidence bundle is stale.",
    )
    parser.add_argument(
        "--strict-claim",
        action="store_true",
        help="Exit non-zero unless the exhaustive no-gap claim is supported.",
    )
    args = parser.parse_args(argv)
    report = audit_completion_claim(args.bundle)
    print(json.dumps(report, indent=2, sort_keys=True))
    if args.strict_claim and not report["exhaustive_claim_ready"]:
        return 1
    if args.strict_current_evidence and not report["current_supported_claim_ready"]:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
