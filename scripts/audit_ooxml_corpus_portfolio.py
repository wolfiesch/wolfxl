#!/usr/bin/env python3
"""Audit aggregate workbook diversity across multiple corpus bucket reports."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_ooxml_corpus_buckets


def audit_corpus_portfolio(
    reports: list[Path],
    *,
    min_workbooks: int = 200,
    min_sources: int = 8,
) -> dict:
    source_reports: list[dict] = []
    bucket_fixtures: dict[str, set[str]] = {
        bucket: set() for bucket in audit_ooxml_corpus_buckets.REQUIRED_BUCKETS
    }
    all_workbooks: set[str] = set()
    skipped_workbook_count = 0

    for report_path in reports:
        payload = json.loads(report_path.read_text())
        workbook_paths = {
            str(workbook.get("path"))
            for workbook in payload.get("workbooks", [])
            if workbook.get("path")
        }
        all_workbooks.update(workbook_paths)
        skipped_workbook_count += int(payload.get("skipped_workbook_count", 0))
        for bucket, paths in payload.get("bucket_fixtures", {}).items():
            bucket_fixtures.setdefault(str(bucket), set()).update(str(path) for path in paths)
        source_reports.append(
            {
                "path": str(report_path),
                "ready": bool(payload.get("ready")),
                "workbook_count": int(payload.get("workbook_count", 0)),
                "skipped_workbook_count": int(payload.get("skipped_workbook_count", 0)),
                "missing_buckets": list(payload.get("missing_buckets", [])),
            }
        )

    combined_bucket_fixtures = {
        bucket: sorted(paths) for bucket, paths in sorted(bucket_fixtures.items())
    }
    missing_buckets = sorted(
        bucket
        for bucket in audit_ooxml_corpus_buckets.REQUIRED_BUCKETS
        if not combined_bucket_fixtures.get(bucket)
    )
    workbook_count = len(all_workbooks)
    source_count = len(source_reports)
    threshold_failures = []
    if workbook_count < min_workbooks:
        threshold_failures.append(
            {
                "id": "min_workbooks",
                "actual": workbook_count,
                "expected_at_least": min_workbooks,
            }
        )
    if source_count < min_sources:
        threshold_failures.append(
            {
                "id": "min_sources",
                "actual": source_count,
                "expected_at_least": min_sources,
            }
        )
    return {
        "ready": not missing_buckets and not threshold_failures,
        "source_count": source_count,
        "workbook_count": workbook_count,
        "skipped_workbook_count": skipped_workbook_count,
        "min_workbooks": min_workbooks,
        "min_sources": min_sources,
        "required_buckets": sorted(audit_ooxml_corpus_buckets.REQUIRED_BUCKETS),
        "missing_buckets": missing_buckets,
        "bucket_counts": {
            bucket: len(paths) for bucket, paths in combined_bucket_fixtures.items()
        },
        "bucket_fixtures": combined_bucket_fixtures,
        "threshold_failures": threshold_failures,
        "source_reports": source_reports,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("reports", nargs="+", type=Path)
    parser.add_argument("--min-workbooks", type=int, default=200)
    parser.add_argument("--min-sources", type=int, default=8)
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero unless aggregate corpus diversity thresholds are met.",
    )
    args = parser.parse_args(argv)
    report = audit_corpus_portfolio(
        args.reports,
        min_workbooks=args.min_workbooks,
        min_sources=args.min_sources,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
