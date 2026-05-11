#!/usr/bin/env python3
"""Summarize a render-compare boundary without treating failures as passes."""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Optional


def summarize_render_boundary(stage_report: Path, render_report: Path) -> dict:
    stage_payload = json.loads(stage_report.read_text())
    render_payload = json.loads(render_report.read_text())
    failures = [
        result
        for result in render_payload.get("results", [])
        if result.get("status") == "failed"
    ]
    failures_by_fixture: dict[str, list[str]] = defaultdict(list)
    failure_messages_by_fixture: dict[str, set[str]] = defaultdict(set)
    for failure in failures:
        fixture = str(failure.get("fixture", ""))
        mutation = str(failure.get("mutation", ""))
        message = str(failure.get("message", "")).strip()
        failures_by_fixture[fixture].append(mutation)
        if message:
            failure_messages_by_fixture[fixture].add(message)

    status_counts = Counter(
        str(result.get("status")) for result in render_payload.get("results", [])
    )
    selected_count = int(stage_payload.get("selected_count", 0))
    render_results = render_payload.get("results", [])
    observed_fixtures = {
        str(result.get("fixture", ""))
        for result in render_results
        if isinstance(result, dict) and result.get("fixture")
    }
    consistency_issues = []
    if selected_count < len(observed_fixtures):
        consistency_issues.append(
            "selected_count is smaller than the number of unique fixtures "
            "observed in the render report"
        )
    if selected_count < len(failures_by_fixture):
        consistency_issues.append(
            "selected_count is smaller than the number of unique failed fixtures"
        )
    renderable_subset_count = max(0, selected_count - len(failures_by_fixture))
    ready = (
        bool(stage_payload.get("ready"))
        and renderable_subset_count > 0
        and not consistency_issues
    )
    return {
        "ready": ready,
        "purpose": (
            "Boundary record for a deterministic random-holdout Excel render "
            "attempt before pinning a renderable subset."
        ),
        "holdout_report": str(stage_report),
        "render_report": str(render_report),
        "seed": stage_payload.get("seed"),
        "sample_size": stage_payload.get("sample_size"),
        "selected_count": selected_count,
        "selected_source_count": stage_payload.get("selected_source_count"),
        "selected_bucket_counts": stage_payload.get("selected_bucket_counts"),
        "render_engine": render_payload.get("render_engine"),
        "excel_print_area": render_payload.get("excel_print_area"),
        "mutations": render_payload.get("mutations"),
        "render_result_count": render_payload.get("result_count"),
        "render_failure_count": render_payload.get("failure_count"),
        "status_counts": dict(sorted(status_counts.items())),
        "observed_fixture_count": len(observed_fixtures),
        "consistency_issues": consistency_issues,
        "excluded_from_renderable_subset": [
            {
                "fixture": fixture,
                "mutations": mutations,
                "failure_messages": sorted(failure_messages_by_fixture[fixture]),
            }
            for fixture, mutations in sorted(failures_by_fixture.items())
        ],
        "renderable_subset_count": renderable_subset_count,
    }


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage_report", type=Path)
    parser.add_argument("render_report", type=Path)
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero unless a non-empty renderable subset is available.",
    )
    args = parser.parse_args(argv)
    report = summarize_render_boundary(args.stage_report, args.render_report)
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
