#!/usr/bin/env python3
"""Audit rendered equivalence for retarget-external-links OOXML mutations.

The normal intentional render smoke proves Excel can export the retargeted
workbook. This audit is stricter for the visually neutral retarget operation:
it exports the matching before workbook, rasterizes the same sampled pages, and
checks that the rendered pages remain equivalent after external-link targets
are rewritten.
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

MUTATION = "retarget_external_links"


def audit_retarget_external_links_render_equivalence(
    render_report_path: Path,
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
    for result in payload.get("results", []):
        if result.get("mutation") != MUTATION:
            continue
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
            audit_result = replace(
                audit_result,
                message=(
                    "external-link retarget render equivalent: "
                    f"max_normalized_rmse={audit_result.max_normalized_rmse:.8f}"
                ),
            )
        results.append(audit_result)

    passed_count = sum(1 for result in results if result.status == "passed")
    failure_count = sum(1 for result in results if result.status == "failed")
    inconclusive_count = sum(1 for result in results if result.status == "inconclusive")
    skipped_count = sum(1 for result in results if result.status == "skipped")
    return {
        "render_report": str(render_report_path),
        "render_engine": render_engine,
        "mutation": MUTATION,
        "max_normalized_rmse_threshold": max_normalized_rmse,
        "result_count": len(results),
        "passed_count": passed_count,
        "failure_count": failure_count,
        "inconclusive_count": inconclusive_count,
        "skipped_count": skipped_count,
        "ready": passed_count > 0 and failure_count == 0 and inconclusive_count == 0,
        "results": [asdict(result) for result in results],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("render_report", type=Path)
    parser.add_argument("--max-normalized-rmse", type=float, default=0.001)
    parser.add_argument("--timeout", type=int, default=30)
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when external-link retarget render equivalence is not proven.",
    )
    args = parser.parse_args(argv)
    report = audit_retarget_external_links_render_equivalence(
        args.render_report,
        max_normalized_rmse=args.max_normalized_rmse,
        timeout=args.timeout,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
