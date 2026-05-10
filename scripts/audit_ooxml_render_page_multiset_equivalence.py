#!/usr/bin/env python3
"""Audit exact rendered-page multiset equivalence between two render reports.

Some Microsoft Excel operations can assign workbook/window state to different
printed pages while still producing the exact same set of page images. This
audit is useful for native-baseline comparisons, for example comparing a WolfXL
``copy_first_sheet`` render against a workbook copied by desktop Excel itself.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from collections import Counter, defaultdict, deque
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

PASSING_RENDER_STATUSES = {"passed", "rendered", "sampled_rendered"}
PAGE_RE_TEMPLATE = r"^{prefix}-pages-(?P<page>\d+)(?:-\d+)?\.png$"


@dataclass(frozen=True)
class PageMultisetEquivalenceResult:
    left_fixture: str
    right_fixture: str
    status: str
    left_page_count: int
    right_page_count: int
    differing_hash_count: int
    remapped_pages: list[dict[str, int]]
    message: str


def audit_render_page_multiset_equivalence(
    left_report_path: Path,
    right_report_path: Path,
    *,
    left_mutation: str | None = None,
    right_mutation: str | None = None,
    left_prefix: str = "after",
    right_prefix: str = "before",
) -> dict:
    left_payload = json.loads(left_report_path.read_text())
    right_payload = json.loads(right_report_path.read_text())

    left_result = _select_result(left_payload, left_mutation)
    right_result = _select_result(right_payload, right_mutation)
    result = _audit_result(
        left_result,
        right_result,
        left_prefix=left_prefix,
        right_prefix=right_prefix,
    )
    passed_count = 1 if result.status == "passed" else 0
    failure_count = 1 if result.status == "failed" else 0
    inconclusive_count = 1 if result.status == "inconclusive" else 0
    return {
        "left_render_report": str(left_report_path),
        "right_render_report": str(right_report_path),
        "left_mutation": left_mutation,
        "right_mutation": right_mutation,
        "left_prefix": left_prefix,
        "right_prefix": right_prefix,
        "result_count": 1,
        "passed_count": passed_count,
        "failure_count": failure_count,
        "inconclusive_count": inconclusive_count,
        "ready": result.status == "passed",
        "results": [asdict(result)],
    }


def _select_result(payload: dict, mutation: str | None) -> dict | None:
    for result in payload.get("results", []):
        if mutation is None or result.get("mutation") == mutation:
            return result
    return None


def _audit_result(
    left_result: dict | None,
    right_result: dict | None,
    *,
    left_prefix: str,
    right_prefix: str,
) -> PageMultisetEquivalenceResult:
    if left_result is None or right_result is None:
        return PageMultisetEquivalenceResult(
            left_fixture=_fixture(left_result),
            right_fixture=_fixture(right_result),
            status="inconclusive",
            left_page_count=0,
            right_page_count=0,
            differing_hash_count=0,
            remapped_pages=[],
            message="matching render result was not found in one or both reports",
        )
    left_fixture = _fixture(left_result)
    right_fixture = _fixture(right_result)
    for side, result in (("left", left_result), ("right", right_result)):
        status = str(result.get("status", ""))
        if status not in PASSING_RENDER_STATUSES:
            return PageMultisetEquivalenceResult(
                left_fixture=left_fixture,
                right_fixture=right_fixture,
                status="failed",
                left_page_count=0,
                right_page_count=0,
                differing_hash_count=0,
                remapped_pages=[],
                message=f"{side} render result status is not passing: {status or '<missing>'}",
            )

    left_dir = _result_dir(left_result, left_prefix)
    right_dir = _result_dir(right_result, right_prefix)
    if left_dir is None or right_dir is None:
        return PageMultisetEquivalenceResult(
            left_fixture=left_fixture,
            right_fixture=right_fixture,
            status="inconclusive",
            left_page_count=0,
            right_page_count=0,
            differing_hash_count=0,
            remapped_pages=[],
            message="one or both render results do not include the requested PDF field",
        )

    left_pages = _page_hashes(left_dir, left_prefix)
    right_pages = _page_hashes(right_dir, right_prefix)
    left_counter = Counter(left_pages.values())
    right_counter = Counter(right_pages.values())
    differing_hash_count = sum((left_counter - right_counter).values()) + sum(
        (right_counter - left_counter).values()
    )
    remapped_pages = _remapped_pages(left_pages, right_pages)
    if left_counter == right_counter:
        message = (
            "rendered page multisets are exactly equivalent"
            if not remapped_pages
            else "rendered page multisets are exactly equivalent with page remapping"
        )
        return PageMultisetEquivalenceResult(
            left_fixture=left_fixture,
            right_fixture=right_fixture,
            status="passed",
            left_page_count=len(left_pages),
            right_page_count=len(right_pages),
            differing_hash_count=0,
            remapped_pages=remapped_pages,
            message=message,
        )

    return PageMultisetEquivalenceResult(
        left_fixture=left_fixture,
        right_fixture=right_fixture,
        status="failed",
        left_page_count=len(left_pages),
        right_page_count=len(right_pages),
        differing_hash_count=differing_hash_count,
        remapped_pages=remapped_pages,
        message="rendered page multisets differ",
    )


def _fixture(result: dict | None) -> str:
    if result is None:
        return ""
    return str(result.get("fixture", ""))


def _result_dir(result: dict, prefix: str) -> Path | None:
    pdf = result.get(f"{prefix}_pdf")
    if not isinstance(pdf, str) or not pdf:
        return None
    return Path(pdf).parent.parent


def _page_hashes(result_dir: Path, prefix: str) -> dict[int, str]:
    pattern = re.compile(PAGE_RE_TEMPLATE.format(prefix=re.escape(prefix)))
    pages: dict[int, str] = {}
    for path in result_dir.glob(f"{prefix}-pages-*.png"):
        match = pattern.match(path.name)
        if match is None:
            continue
        pages[int(match.group("page"))] = hashlib.sha256(path.read_bytes()).hexdigest()
    return pages


def _remapped_pages(
    left_pages: dict[int, str],
    right_pages: dict[int, str],
) -> list[dict[str, int]]:
    right_by_hash: dict[str, deque[int]] = defaultdict(deque)
    for page, digest in sorted(right_pages.items()):
        right_by_hash[digest].append(page)

    remapped: list[dict[str, int]] = []
    for left_page, digest in sorted(left_pages.items()):
        candidates = right_by_hash.get(digest)
        if not candidates:
            continue
        if left_page in candidates:
            right_page = left_page
        else:
            right_page = candidates[0]
        if right_page != left_page:
            remapped.append({"left_page": left_page, "right_page": right_page})
    return remapped


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("left_render_report", type=Path)
    parser.add_argument("right_render_report", type=Path)
    parser.add_argument("--left-mutation")
    parser.add_argument("--right-mutation")
    parser.add_argument("--left-prefix", choices=("before", "after"), default="after")
    parser.add_argument("--right-prefix", choices=("before", "after"), default="before")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when page multiset equivalence is not proven.",
    )
    args = parser.parse_args(list(argv) if argv is not None else None)

    report = audit_render_page_multiset_equivalence(
        args.left_render_report,
        args.right_render_report,
        left_mutation=args.left_mutation,
        right_mutation=args.right_mutation,
        left_prefix=args.left_prefix,
        right_prefix=args.right_prefix,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
