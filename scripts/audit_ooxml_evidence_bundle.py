#!/usr/bin/env python3
"""Verify a pinned OOXML fidelity evidence bundle."""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class BundleIssue:
    report: str
    path: str
    check: str
    message: str


RENDER_COMPARE_PRODUCER_MARKER = "scripts/run_ooxml_render_compare.py"
RENDER_COMPARE_RENDERED_STATUSES = {
    "passed",
    "sampled_passed",
    "rendered",
    "sampled_rendered",
}


def audit_bundle(manifest_path: Path) -> dict:
    manifest = json.loads(manifest_path.read_text())
    base_dir = manifest_path.resolve().parent
    issues: list[BundleIssue] = []
    report_results = []
    seen_names: dict[str, str] = {}
    seen_paths: dict[str, str] = {}
    for report in manifest.get("reports", []):
        name = str(report["name"])
        path = _resolve_path(str(report["path"]), base_dir)
        producer = report.get("producer")
        path_text = str(path)
        if name in seen_names:
            issues.append(
                BundleIssue(
                    report=name,
                    path=path_text,
                    check="duplicate_name",
                    message=f"duplicate report name also used for {seen_names[name]}",
                )
            )
        else:
            seen_names[name] = path_text
        if path_text in seen_paths:
            issues.append(
                BundleIssue(
                    report=name,
                    path=path_text,
                    check="duplicate_path",
                    message=f"duplicate report path also used by {seen_paths[path_text]}",
                )
            )
        else:
            seen_paths[path_text] = name
        result = {
            "name": name,
            "path": path_text,
            "producer": producer,
            "exists": path.is_file(),
            "checks": [],
        }
        if not isinstance(producer, str) or not producer.strip():
            issues.append(
                BundleIssue(
                    report=name,
                    path=path_text,
                    check="producer",
                    message="producer command is missing",
                )
            )
        if not path.is_file():
            issues.append(
                BundleIssue(
                    report=name,
                    path=path_text,
                    check="exists",
                    message="report file is missing",
                )
            )
            report_results.append(result)
            continue
        payload = json.loads(path.read_text())
        for check in report.get("expect", []):
            check_result = _evaluate_check(payload, check)
            result["checks"].append(check_result)
            if not check_result["passed"]:
                issues.append(
                    BundleIssue(
                        report=name,
                        path=path_text,
                        check=str(check.get("path")),
                        message=check_result["message"],
                    )
                )
        for check_result in _evaluate_implicit_checks(payload, producer):
            result["checks"].append(check_result)
            if not check_result["passed"]:
                issues.append(
                    BundleIssue(
                        report=name,
                        path=path_text,
                        check=check_result["path"],
                        message=check_result["message"],
                    )
                )
        report_results.append(result)
    return {
        "manifest": str(manifest_path),
        "report_count": len(report_results),
        "producer_count": sum(
            1
            for report in report_results
            if isinstance(report.get("producer"), str) and report["producer"].strip()
        ),
        "issue_count": len(issues),
        "ready": not issues,
        "reports": report_results,
        "issues": [issue.__dict__ for issue in issues],
    }


def _resolve_path(path: str, base_dir: Path) -> Path:
    candidate = Path(path)
    if candidate.is_absolute():
        return candidate
    return (base_dir / candidate).resolve()


def _evaluate_check(payload: object, check: dict[str, object]) -> dict:
    path = str(check["path"])
    try:
        actual = _get_path(payload, path)
    except (KeyError, IndexError, ValueError, TypeError) as exc:
        return {
            "path": path,
            "actual": None,
            "passed": False,
            "message": f"missing path {path!r}: {exc}",
        }
    if "equals" in check:
        expected = check["equals"]
        passed = actual == expected
        message = "ok" if passed else f"expected {path} == {expected!r}, got {actual!r}"
    elif "at_least" in check:
        expected = check["at_least"]
        passed = isinstance(actual, (int, float)) and actual >= expected
        message = "ok" if passed else f"expected {path} >= {expected!r}, got {actual!r}"
    elif "length_at_least" in check:
        expected = check["length_at_least"]
        actual_len = len(actual) if isinstance(actual, (list, dict, str)) else None
        passed = actual_len is not None and actual_len >= expected
        message = "ok" if passed else f"expected len({path}) >= {expected!r}, got {actual_len!r}"
    elif "length" in check:
        expected = check["length"]
        actual_len = len(actual) if isinstance(actual, (list, dict, str)) else None
        passed = actual_len is not None and actual_len == expected
        message = "ok" if passed else f"expected len({path}) == {expected!r}, got {actual_len!r}"
    elif "contains" in check:
        expected = check["contains"]
        passed = _contains(actual, expected)
        message = "ok" if passed else f"expected {path} to contain {expected!r}, got {actual!r}"
    else:
        raise ValueError(f"unsupported evidence check: {check}")
    return {
        "path": path,
        "actual": actual,
        "passed": passed,
        "message": message,
    }


def _contains(actual: object, expected: object) -> bool:
    if isinstance(actual, str):
        return isinstance(expected, str) and expected in actual
    if isinstance(actual, (list, tuple)):
        return expected in actual
    if isinstance(actual, (set, dict)):
        try:
            return expected in actual
        except TypeError:
            return False
    return False


def _evaluate_implicit_checks(payload: object, producer: object) -> list[dict]:
    if (
        isinstance(producer, str)
        and RENDER_COMPARE_PRODUCER_MARKER in producer
        and _is_render_compare_report(payload)
    ):
        return [_evaluate_render_compare_statuses(payload)]
    return []


def _is_render_compare_report(payload: object) -> bool:
    return (
        isinstance(payload, dict)
        and "max_normalized_rmse_threshold" in payload
        and "density" in payload
        and "results" in payload
    )


def _evaluate_render_compare_statuses(payload: object) -> dict:
    path = "results.*.status"
    if not isinstance(payload, dict):
        return {
            "path": path,
            "actual": None,
            "passed": False,
            "message": "expected render compare report to be a JSON object",
        }
    results = payload.get("results")
    if not isinstance(results, list):
        return {
            "path": path,
            "actual": None,
            "passed": False,
            "message": "expected render compare report to include a results list",
        }
    statuses = []
    bad_results = []
    for index, result in enumerate(results):
        status = result.get("status") if isinstance(result, dict) else None
        statuses.append(status)
        if status not in RENDER_COMPARE_RENDERED_STATUSES:
            if isinstance(result, dict):
                fixture = result.get("fixture")
                mutation = result.get("mutation")
            else:
                fixture = None
                mutation = None
            bad_results.append(
                {
                    "index": index,
                    "fixture": fixture,
                    "mutation": mutation,
                    "status": status,
                }
            )
    if bad_results:
        sample = bad_results[:5]
        return {
            "path": path,
            "actual": statuses,
            "passed": False,
            "message": (
                "render compare report contains non-rendered or skipped "
                f"result statuses: {sample!r}"
            ),
        }
    return {
        "path": path,
        "actual": statuses,
        "passed": True,
        "message": "ok",
    }


def _get_path(payload: Any, path: str) -> Any:
    current = payload
    for part in path.split("."):
        if isinstance(current, dict):
            current = current[part]
        elif isinstance(current, list):
            current = current[int(part)]
        else:
            raise KeyError(f"cannot descend into {part!r} in {path!r}")
    return current


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", type=Path)
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when any evidence report/check is missing or stale.",
    )
    args = parser.parse_args(argv)
    report = audit_bundle(args.manifest)
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
