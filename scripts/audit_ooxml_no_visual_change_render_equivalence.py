#!/usr/bin/env python3
"""Audit rendered equivalence for visually neutral OOXML mutations.

The normal intentional render smoke proves Excel can export a mutated workbook.
For mutations that should not change printed output, this audit is stricter: it
exports the matching before workbook, rasterizes the same sampled pages, and
checks that rendered pages remain equivalent after the mutation.
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from dataclasses import asdict, replace
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_ooxml_rename_sheet_render_equivalence as base  # noqa: E402

DEFAULT_MUTATIONS = ("add_data_validation",)
MUTATION_LABELS = {
    "add_data_validation": "add-data-validation",
    "rename_first_sheet": "rename-sheet",
    "retarget_external_links": "external-link retarget",
}


def audit_no_visual_change_render_equivalence(
    render_report_path: Path,
    *,
    mutations: tuple[str, ...] = DEFAULT_MUTATIONS,
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

    wanted = set(mutations)
    results: list[base.RenameSheetEquivalenceResult] = []
    observed_mutations: set[str] = set()
    for result in payload.get("results", []):
        mutation = result.get("mutation")
        if mutation not in wanted:
            continue
        observed_mutations.add(str(mutation))
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
        if audit_result.status == "passed":
            label = MUTATION_LABELS.get(str(mutation), str(mutation).replace("_", "-"))
            audit_result = replace(
                audit_result,
                message=(
                    f"{label} render equivalent: "
                    f"max_normalized_rmse={audit_result.max_normalized_rmse:.8f}"
                ),
            )
        results.append(audit_result)

    passed_count = sum(1 for result in results if result.status == "passed")
    failure_count = sum(1 for result in results if result.status == "failed")
    inconclusive_count = sum(1 for result in results if result.status == "inconclusive")
    skipped_count = sum(1 for result in results if result.status == "skipped")
    missing_mutations = sorted(wanted - observed_mutations)
    return {
        "render_report": str(render_report_path),
        "render_engine": render_engine,
        "mutations": list(mutations),
        "observed_mutations": sorted(observed_mutations),
        "missing_mutations": missing_mutations,
        "max_normalized_rmse_threshold": max_normalized_rmse,
        "result_count": len(results),
        "passed_count": passed_count,
        "failure_count": failure_count,
        "inconclusive_count": inconclusive_count,
        "skipped_count": skipped_count,
        "ready": (
            passed_count > 0
            and failure_count == 0
            and inconclusive_count == 0
            and not missing_mutations
        ),
        "results": [asdict(result) for result in results],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("render_report", type=Path)
    parser.add_argument(
        "--mutation",
        action="append",
        choices=sorted(MUTATION_LABELS),
        help=(
            "Visually neutral mutation to audit. Defaults to add_data_validation. "
            "Repeat to audit multiple mutations in one render report."
        ),
    )
    parser.add_argument("--max-normalized-rmse", type=float, default=0.001)
    parser.add_argument("--timeout", type=int, default=30)
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when visual equivalence is not proven.",
    )
    args = parser.parse_args(argv)
    mutations = tuple(args.mutation) if args.mutation else DEFAULT_MUTATIONS
    report = audit_no_visual_change_render_equivalence(
        args.render_report,
        mutations=mutations,
        max_normalized_rmse=args.max_normalized_rmse,
        timeout=args.timeout,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
