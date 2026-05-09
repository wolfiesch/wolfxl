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
    "excel_render_copy_sheet_equivalence_full_pack_report",
    "excel_render_rename_sheet_equivalence_full_pack_report",
    "excel_render_retarget_external_links_equivalence_full_pack_report",
    "excel_render_add_data_validation_equivalence_full_pack_report",
    "excel_render_add_conditional_formatting_equivalence_full_pack_report",
    "excel_render_move_formula_range_delta_full_pack_report",
    "excel_render_delete_first_row_delta_full_pack_report",
    "excel_render_delete_first_col_delta_full_pack_report",
    "excel_app_open_full_pack_with_cf_verified_coverage_gate",
    "external_oracle_corpus_diversity",
    "corpus_portfolio_diversity",
    "external_oracle_gap_radar",
    "synthgl_recursive_gap_radar",
    "umya_test_files_gap_radar",
    "umya_test_files_quick_plus_structural_mutation_coverage",
    "powerpivot_contoso_sidecar_coverage",
    "powerpivot_contoso_sidecar_excel_expected_unsupported",
    "slicer_shared_two_pivots_sidecar_coverage",
    "slicer_shared_two_pivots_sidecar_interactive_evidence",
    "slicer_shared_two_pivots_sidecar_ui_interaction_evidence",
    "external_link_retarget_excel_app_open",
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
            "The current corpus portfolio spans 236 readable workbooks across 15 "
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
            "neutral add-data-validation and add-conditional-formatting edits, plus "
            "expected visual-delta evidence for move-formula-range and first "
            "row/column deletion; they still do not prove semantic visual "
            "equivalence for every high-risk feature edit."
        ),
    },
    {
        "id": "broader_click_level_interaction_variants",
        "status": "open",
        "reason": (
            "Targeted UI-interaction evidence exists, but broader slicer, timeline, "
            "embedded-control, and prompt variants remain unexhausted."
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
