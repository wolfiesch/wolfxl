#!/usr/bin/env python3
"""Audit that an intentional OOXML mutation produces a rendered visual delta.

Render-smoke proves Excel can export the mutated workbook. For mutations that
are expected to visibly change output, this audit compares the exported after
PDF against the matching before workbook and records whether sampled pages show
an intentional visual delta.
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from dataclasses import asdict, replace
from pathlib import Path
from typing import Optional

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_ooxml_rename_sheet_render_equivalence as base  # noqa: E402

MUTATION_LABELS = {
    "delete_first_col": "delete-first-col",
    "delete_first_row": "delete-first-row",
    "insert_tail_col": "insert-tail-col",
    "insert_tail_row": "insert-tail-row",
    "marker_cell": "marker-cell",
    "move_formula_range": "move-formula-range",
    "style_cell": "style-cell",
}
EXPECTED_PAGE_COUNT_DELTA_MUTATIONS = frozenset(
    {
        "delete_first_col",
        "delete_first_row",
        "insert_tail_col",
        "insert_tail_row",
        "marker_cell",
        "style_cell",
    }
)


def audit_intentional_render_delta(
    render_report_path: Path,
    *,
    mutation: str,
    min_changed_count: int = 1,
    max_unchanged_count: Optional[int] = None,
    max_normalized_rmse: float = 0.001,
    timeout: int = 30,
) -> dict:
    payload = json.loads(render_report_path.read_text())
    compare_cmd = base.run_ooxml_render_compare._find_imagemagick_compare()
    pdftoppm = shutil.which("pdftoppm")
    pdfinfo = shutil.which("pdfinfo")
    render_engine = str(payload.get("render_engine", "libreoffice"))
    density = int(payload.get("density", 96))
    soffice = None
    if render_engine == "libreoffice":
        soffice = base.run_ooxml_app_smoke._find_libreoffice()

    results: list[base.RenameSheetEquivalenceResult] = []
    observed = False
    label = MUTATION_LABELS.get(mutation, mutation.replace("_", "-"))
    for result in payload.get("results", []):
        if result.get("mutation") != mutation:
            continue
        observed = True
        audit_result = base._audit_result(
            result,
            compare_cmd=compare_cmd,
            pdftoppm=pdftoppm,
            pdfinfo=pdfinfo,
            render_engine=render_engine,
            soffice=soffice,
            density=density,
            max_normalized_rmse=max_normalized_rmse,
            timeout=timeout,
        )
        if audit_result.status == "failed" and audit_result.max_normalized_rmse is not None:
            audit_result = replace(
                audit_result,
                status="changed",
                message=(
                    f"{label} intentional render delta observed: "
                    f"max_normalized_rmse={audit_result.max_normalized_rmse:.8f}"
                ),
            )
        elif (
            mutation in EXPECTED_PAGE_COUNT_DELTA_MUTATIONS
            and audit_result.status == "failed"
            and audit_result.message.startswith("page-count mismatch:")
        ):
            audit_result = replace(
                audit_result,
                status="changed",
                message=f"{label} intentional page-count delta observed: {audit_result.message}",
            )
        elif audit_result.status == "passed":
            audit_result = replace(
                audit_result,
                status="unchanged",
                message=(
                    f"{label} sampled pages unchanged: "
                    f"max_normalized_rmse={audit_result.max_normalized_rmse:.8f}"
                ),
            )
        results.append(audit_result)

    changed_count = sum(1 for result in results if result.status == "changed")
    unchanged_count = sum(1 for result in results if result.status == "unchanged")
    failure_count = sum(1 for result in results if result.status == "failed")
    inconclusive_count = sum(1 for result in results if result.status == "inconclusive")
    skipped_count = sum(1 for result in results if result.status == "skipped")
    unchanged_count_allowed = (
        max_unchanged_count is None or unchanged_count <= max_unchanged_count
    )
    missing_mutation = not observed
    ready = (
        changed_count >= min_changed_count
        and unchanged_count_allowed
        and failure_count == 0
        and inconclusive_count == 0
        and skipped_count == 0
        and not missing_mutation
    )
    return {
        "render_report": str(render_report_path),
        "render_engine": render_engine,
        "mutation": mutation,
        "missing_mutation": missing_mutation,
        "max_normalized_rmse_threshold": max_normalized_rmse,
        "min_changed_count": min_changed_count,
        "max_unchanged_count": max_unchanged_count,
        "result_count": len(results),
        "changed_count": changed_count,
        "unchanged_count": unchanged_count,
        "failure_count": failure_count,
        "inconclusive_count": inconclusive_count,
        "skipped_count": skipped_count,
        "ready": ready,
        "results": [asdict(result) for result in results],
    }


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("render_report", type=Path)
    parser.add_argument("--mutation", required=True, choices=sorted(MUTATION_LABELS))
    parser.add_argument("--min-changed-count", type=int, default=1)
    parser.add_argument("--max-unchanged-count", type=int)
    parser.add_argument("--max-normalized-rmse", type=float, default=0.001)
    parser.add_argument("--timeout", type=int, default=30)
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when the expected visual delta is not proven.",
    )
    args = parser.parse_args(argv)
    report = audit_intentional_render_delta(
        args.render_report,
        mutation=args.mutation,
        min_changed_count=args.min_changed_count,
        max_unchanged_count=args.max_unchanged_count,
        max_normalized_rmse=args.max_normalized_rmse,
        timeout=args.timeout,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
