#!/usr/bin/env python3
"""Select and audit a deterministic random holdout from a corpus portfolio."""

from __future__ import annotations

import argparse
import hashlib
import json
import random
import shutil
import sys
from pathlib import Path
from typing import Optional

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_ooxml_corpus_buckets


def audit_random_holdout(
    portfolio_path: Path,
    *,
    sample_size: int,
    seed: str,
    min_sample_size: int,
    min_sources: int,
    stage_dir: Optional[Path] = None,
) -> dict:
    portfolio = json.loads(portfolio_path.read_text())
    workbook_sources, workbook_buckets = _load_workbook_index(portfolio, portfolio_path)
    population = sorted(workbook_sources)
    selected = _select_population(population, sample_size=sample_size, seed=seed)

    selected_source_paths = sorted(
        {source for workbook in selected for source in workbook_sources[workbook]}
    )
    selected_bucket_fixtures = {
        bucket: [
            workbook
            for workbook in selected
            if bucket in workbook_buckets.get(workbook, set())
        ]
        for bucket in sorted(audit_ooxml_corpus_buckets.REQUIRED_BUCKETS)
    }
    missing_selected_files = [
        workbook for workbook in selected if not Path(workbook).is_file()
    ]
    staged_workbooks = _stage_workbooks(selected, stage_dir) if stage_dir else []
    threshold_failures = _threshold_failures(
        selected_count=len(selected),
        source_count=len(selected_source_paths),
        min_sample_size=min_sample_size,
        min_sources=min_sources,
    )
    ready = (
        len(selected) == min(sample_size, len(population))
        and not threshold_failures
        and not missing_selected_files
    )
    return {
        "ready": ready,
        "portfolio_path": str(portfolio_path),
        "seed": seed,
        "sample_method": "python_random_without_replacement_sorted_population",
        "sample_size": sample_size,
        "min_sample_size": min_sample_size,
        "min_sources": min_sources,
        "population_count": len(population),
        "selected_count": len(selected),
        "selected_source_count": len(selected_source_paths),
        "selected_source_paths": selected_source_paths,
        "required_buckets": sorted(audit_ooxml_corpus_buckets.REQUIRED_BUCKETS),
        "selected_bucket_counts": {
            bucket: len(paths) for bucket, paths in selected_bucket_fixtures.items()
        },
        "missing_selected_files": missing_selected_files,
        "threshold_failures": threshold_failures,
        "selected_workbooks": [
            {
                "path": workbook,
                "source_reports": workbook_sources[workbook],
                "buckets": sorted(workbook_buckets.get(workbook, set())),
            }
            for workbook in selected
        ],
        "stage_dir": str(stage_dir) if stage_dir else None,
        "staged_workbooks": staged_workbooks,
    }


def _load_workbook_index(
    portfolio: dict,
    portfolio_path: Path,
) -> tuple[dict[str, list[str]], dict[str, set[str]]]:
    workbook_sources: dict[str, set[str]] = {}
    workbook_buckets: dict[str, set[str]] = {}
    for source in portfolio.get("source_reports", []):
        if source.get("contributes_source") is False:
            continue
        raw_report_path = source.get("path")
        if not raw_report_path:
            continue
        report_path = _resolve_report_path(Path(str(raw_report_path)), portfolio_path)
        if not report_path.is_file():
            continue
        payload = json.loads(report_path.read_text())
        report_key = str(report_path.resolve())
        for workbook in payload.get("workbooks", []):
            workbook_path = workbook.get("path")
            if not workbook_path:
                continue
            workbook_key = str(_resolve_workbook_path(Path(str(workbook_path)), report_path))
            workbook_sources.setdefault(workbook_key, set()).add(report_key)
            workbook_buckets.setdefault(workbook_key, set()).update(
                str(bucket) for bucket in workbook.get("buckets", [])
            )
        for bucket, paths in payload.get("bucket_fixtures", {}).items():
            for workbook_path in paths:
                workbook_key = str(
                    _resolve_workbook_path(Path(str(workbook_path)), report_path)
                )
                workbook_buckets.setdefault(workbook_key, set()).add(str(bucket))
    return (
        {workbook: sorted(sources) for workbook, sources in workbook_sources.items()},
        workbook_buckets,
    )


def _resolve_report_path(report_path: Path, portfolio_path: Path) -> Path:
    if report_path.is_absolute():
        return report_path
    return portfolio_path.parent / report_path


def _resolve_workbook_path(workbook_path: Path, report_path: Path) -> Path:
    if workbook_path.is_absolute():
        return workbook_path
    return (report_path.parent / workbook_path).resolve()


def _select_population(population: list[str], *, sample_size: int, seed: str) -> list[str]:
    rng = random.Random(seed)
    if sample_size >= len(population):
        return list(population)
    return sorted(rng.sample(population, sample_size))


def _threshold_failures(
    *,
    selected_count: int,
    source_count: int,
    min_sample_size: int,
    min_sources: int,
) -> list[dict]:
    failures = []
    if selected_count < min_sample_size:
        failures.append(
            {
                "id": "min_sample_size",
                "actual": selected_count,
                "expected_at_least": min_sample_size,
            }
        )
    if source_count < min_sources:
        failures.append(
            {
                "id": "min_sources",
                "actual": source_count,
                "expected_at_least": min_sources,
            }
        )
    return failures


def _stage_workbooks(selected: list[str], stage_dir: Path) -> list[dict]:
    stage_dir.mkdir(parents=True, exist_ok=True)
    staged = []
    for index, workbook in enumerate(selected, start=1):
        source = Path(workbook)
        digest = hashlib.sha256(workbook.encode("utf-8")).hexdigest()[:10]
        target = stage_dir / f"{index:03d}-{digest}-{_safe_name(source.name)}"
        if source.is_file():
            shutil.copy2(source, target)
            status = "copied"
        else:
            status = "missing_source"
        staged.append({"source": workbook, "path": str(target), "status": status})
    return staged


def _safe_name(name: str) -> str:
    return "".join(char if char.isalnum() or char in ".-_" else "_" for char in name)


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("portfolio", type=Path)
    parser.add_argument("--sample-size", type=int, default=50)
    parser.add_argument("--seed", default="wolfxl-random-holdout-20260511-v1")
    parser.add_argument("--min-sample-size", type=int, default=50)
    parser.add_argument("--min-sources", type=int, default=8)
    parser.add_argument("--stage-dir", type=Path)
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero unless sample and source thresholds are met.",
    )
    args = parser.parse_args(argv)
    report = audit_random_holdout(
        args.portfolio,
        sample_size=args.sample_size,
        seed=args.seed,
        min_sample_size=args.min_sample_size,
        min_sources=args.min_sources,
        stage_dir=args.stage_dir,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
