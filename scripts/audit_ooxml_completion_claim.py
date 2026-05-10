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
    "excel_ui_interaction_marker_table_slicer_evidence",
    "excel_ui_interaction_style_table_slicer_evidence",
    "excel_ui_interaction_copy_remove_table_slicer_evidence",
    "excel_ui_interaction_marker_pivot_chart_slicer_evidence",
    "excel_ui_interaction_style_pivot_chart_slicer_evidence",
    "excel_ui_interaction_copy_remove_pivot_chart_slicer_evidence",
    "excel_ui_interaction_marker_external_tool_pivot_slicer_evidence",
    "excel_ui_interaction_style_external_tool_pivot_slicer_evidence",
    "excel_ui_interaction_copy_remove_external_tool_pivot_slicer_evidence",
    "excel_ui_interaction_marker_timeline_evidence",
    "excel_ui_interaction_style_timeline_evidence",
    "excel_ui_interaction_copy_remove_timeline_evidence",
    "excel_ui_interaction_marker_external_link_current_prompt_evidence",
    "excel_ui_interaction_style_external_link_current_prompt_evidence",
    "excel_ui_interaction_copy_remove_external_link_current_prompt_evidence",
    "excel_ui_interaction_marker_umya_external_link_forced_prompt_evidence",
    "excel_ui_interaction_style_umya_external_link_forced_prompt_evidence",
    "excel_ui_interaction_copy_remove_umya_external_link_forced_prompt_evidence",
    "excel_ui_interaction_marker_umya_listbox_control_evidence",
    "excel_ui_interaction_style_umya_listbox_control_evidence",
    "excel_ui_interaction_copy_remove_umya_listbox_control_evidence",
    "excel_ui_interaction_marker_control_evidence",
    "excel_ui_interaction_marker_button_control_evidence",
    "excel_ui_interaction_style_control_evidence",
    "excel_ui_interaction_style_button_control_evidence",
    "excel_ui_interaction_copy_remove_control_evidence",
    "excel_ui_interaction_copy_remove_button_control_evidence",
    "external_oracle_corpus_diversity",
    "corpus_portfolio_diversity",
    "external_oracle_gap_radar",
    "synthgl_recursive_gap_radar",
    "umya_test_files_gap_radar",
    "umya_test_files_quick_plus_structural_mutation_coverage",
    "powerpivot_contoso_sidecar_coverage",
    "powerpivot_contoso_sidecar_excel_expected_unsupported",
    "powerpivot_contoso_sidecar_unsupported_content_prompt_evidence",
    "powerpivot_contoso_sidecar_marker_unsupported_content_prompt_evidence",
    "powerpivot_contoso_sidecar_style_unsupported_content_prompt_evidence",
    "powerpivot_contoso_sidecar_copy_remove_unsupported_content_prompt_evidence",
    "slicer_shared_two_pivots_sidecar_coverage",
    "slicer_shared_two_pivots_sidecar_copy_remove_sheet_render_equivalence",
    "slicer_shared_two_pivots_sidecar_rename_sheet_render_equivalence",
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
            "The current corpus portfolio spans 280 readable workbooks across 19 "
            "source reports and covers all required diversity buckets, but it is "
            "still not customer-scale or random real-world Excel evidence."
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
            "equivalence plus delete-first-row/column visual deltas on a high-risk "
            "workbook; ILPA side evidence separately proves external-link "
            "retarget, add-data-validation, add-conditional-formatting, and "
            "add-remove-chart and copy-remove-sheet render "
            "equivalence on public "
            "finance templates, and WBD wDesk "
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
            "add-remove-chart render equivalence on two renderable Las Vegas "
            "QoE databooks, "
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
            "Targeted UI-interaction evidence exists, including "
            "marker-cell-mutated and style-cell-mutated table-slicer, "
            "copy-remove-sheet-mutated table-slicer, "
            "marker-cell-mutated and style-cell-mutated pivot-chart "
            "slicer, copy-remove-sheet-mutated pivot-chart slicer, "
            "marker-cell-mutated and style-cell-mutated external-tool "
            "pivot-slicer item clicks, copy-remove-sheet-mutated external-tool "
            "pivot-slicer item clicks, marker-cell-mutated, style-cell-mutated, "
            "and copy-remove-sheet-mutated shared pivot-slicer cache clicks, plus "
            "marker-cell-mutated, style-cell-mutated, and copy-remove-sheet-mutated "
            "timeline clicks, plus marker-cell-mutated, style-cell-mutated, "
            "and copy-remove-sheet-mutated macro security prompt paths, plus "
            "marker-cell-mutated, style-cell-mutated, and copy-remove-sheet-mutated "
            "current-setting external-link prompt paths, marker-cell-mutated and style-cell-mutated adjacent "
            "issue-corpus forced external-link prompt paths, "
            "copy-remove-sheet-mutated adjacent issue-corpus forced external-link prompt path, "
            "marker-cell-mutated, style-cell-mutated, and "
            "copy-remove-sheet-mutated adjacent issue-corpus list-box "
            "clicks, and source, marker-cell-mutated, "
            "style-cell-mutated, and copy-remove-sheet-mutated PowerView "
            "read-only prompt paths, plus "
            "marker-cell-mutated and style-cell-mutated list-box and "
            "button-control click persistence, plus copy-remove-sheet-mutated "
            "list-box and button-control click persistence, but "
            "broader slicer, timeline, embedded-control, and prompt variants "
            "remain unexhausted."
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
        *OPEN_REQUIREMENTS,
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
