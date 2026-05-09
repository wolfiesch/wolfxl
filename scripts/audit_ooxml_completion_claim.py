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

OPEN_REQUIREMENTS = (
    {
        "id": "broader_real_world_corpus_diversity",
        "status": "open",
        "reason": (
            "The current evidence spans pinned and sidecar corpora, but not a broad "
            "customer-scale or random real-world Excel corpus."
        ),
    },
    {
        "id": "feature_specific_intentional_render_equivalence",
        "status": "open",
        "reason": (
            "Intentional Microsoft Excel render checks prove renderability for selected "
            "edits, not semantic visual equivalence for every high-risk feature edit."
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
        "current_supported_claim_ready": bundle_ready,
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
