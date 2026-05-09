#!/usr/bin/env python3
"""Audit rendered equivalence for copy-first-sheet OOXML mutations.

The intentional render smoke proves the copied workbook can be exported by
Excel, but it does not prove the copied sheet still looks like the original
sheet. This audit reuses the already-rasterized after-workbook pages and checks
whether page 1 has an exact or near-exact later-page match.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from dataclasses import asdict, dataclass
from xml.etree import ElementTree
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import run_ooxml_render_compare  # noqa: E402

MUTATION = "copy_first_sheet"
PASSING_RENDER_STATUSES = {"rendered", "sampled_rendered"}
PAGE_RE = re.compile(r"^after-pages-(?P<page>\d+)(?:-\d+)?\.png$")


@dataclass(frozen=True)
class CopySheetEquivalenceResult:
    fixture: str
    status: str
    source_page: int | None
    matched_page: int | None
    compared_pages: list[int]
    normalized_rmse: float | None
    sampled: bool
    message: str


def audit_copy_sheet_render_equivalence(
    render_report_path: Path,
    max_normalized_rmse: float = 0.001,
    timeout: int = 30,
) -> dict:
    payload = json.loads(render_report_path.read_text())
    compare_cmd = run_ooxml_render_compare._find_imagemagick_compare()
    results: list[CopySheetEquivalenceResult] = []
    for result in payload.get("results", []):
        if result.get("mutation") != MUTATION:
            continue
        results.append(
            _audit_result(
                result,
                compare_cmd=compare_cmd,
                max_normalized_rmse=max_normalized_rmse,
                timeout=timeout,
            )
        )

    passed_count = sum(1 for result in results if result.status == "passed")
    failure_count = sum(1 for result in results if result.status == "failed")
    inconclusive_count = sum(1 for result in results if result.status == "inconclusive")
    skipped_count = sum(1 for result in results if result.status == "skipped")
    return {
        "render_report": str(render_report_path),
        "mutation": MUTATION,
        "max_normalized_rmse_threshold": max_normalized_rmse,
        "result_count": len(results),
        "passed_count": passed_count,
        "failure_count": failure_count,
        "inconclusive_count": inconclusive_count,
        "skipped_count": skipped_count,
        "ready": passed_count > 0 and failure_count == 0,
        "results": [asdict(result) for result in results],
    }


def _audit_result(
    result: dict,
    compare_cmd: tuple[str, ...] | None,
    max_normalized_rmse: float,
    timeout: int,
) -> CopySheetEquivalenceResult:
    fixture = str(result.get("fixture", ""))
    status = str(result.get("status", ""))
    if status not in PASSING_RENDER_STATUSES:
        return CopySheetEquivalenceResult(
            fixture=fixture,
            status="skipped",
            source_page=None,
            matched_page=None,
            compared_pages=[],
            normalized_rmse=None,
            sampled=False,
            message=f"render result status is not passing: {status}",
        )

    after_pdf = result.get("after_pdf")
    if not isinstance(after_pdf, str):
        return _inconclusive(fixture, "render result has no after_pdf")
    result_dir = Path(after_pdf).parent.parent
    if _source_first_sheet_hidden(result_dir):
        return _inconclusive(
            fixture,
            "source first sheet is hidden; no rendered source page to compare",
        )
    pages = _after_pages(result_dir)
    if 1 not in pages:
        return _inconclusive(fixture, "page 1 image is missing")
    later_pages = [page for page in sorted(pages) if page > 1]
    if not later_pages:
        return _inconclusive(fixture, "no later rendered page to compare against page 1")

    best_page: int | None = None
    best_rmse: float | None = None
    source = pages[1]
    for page_number in later_pages:
        rmse = _page_rmse(source, pages[page_number], compare_cmd, timeout)
        if best_rmse is None or rmse < best_rmse:
            best_rmse = rmse
            best_page = page_number
        if rmse <= max_normalized_rmse:
            return CopySheetEquivalenceResult(
                fixture=fixture,
                status="passed",
                source_page=1,
                matched_page=page_number,
                compared_pages=later_pages,
                normalized_rmse=rmse,
                sampled=_is_sampled(result),
                message=(
                    f"page 1 matched copied-sheet candidate page {page_number}: "
                    f"normalized_rmse={rmse:.8f}"
                ),
            )

    sampled = _is_sampled(result)
    message = (
        f"no sampled later page matched page 1; best_page={best_page} "
        f"best_normalized_rmse={best_rmse:.8f}"
        if best_rmse is not None
        else "no comparable later page was available"
    )
    return CopySheetEquivalenceResult(
        fixture=fixture,
        status="inconclusive" if sampled else "failed",
        source_page=1,
        matched_page=None,
        compared_pages=later_pages,
        normalized_rmse=best_rmse,
        sampled=sampled,
        message=message,
    )


def _after_pages(result_dir: Path) -> dict[int, Path]:
    pages: dict[int, Path] = {}
    for path in result_dir.glob("after-pages-*.png"):
        match = PAGE_RE.match(path.name)
        if match is None:
            continue
        page = int(match.group("page"))
        pages[page] = path
    return pages


def _page_rmse(
    source: Path,
    candidate: Path,
    compare_cmd: tuple[str, ...] | None,
    timeout: int,
) -> float:
    if _sha256(source) == _sha256(candidate):
        return 0.0
    if compare_cmd is None:
        raise RuntimeError(
            "ImageMagick compare is required when page PNGs are not byte-identical"
        )
    return run_ooxml_render_compare._normalized_rmse(
        compare_cmd,
        source,
        candidate,
        timeout,
    )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _is_sampled(result: dict) -> bool:
    status = str(result.get("status", ""))
    page_count = result.get("page_count")
    compared_page_count = result.get("compared_page_count")
    return status.startswith("sampled_") or (
        isinstance(page_count, int)
        and isinstance(compared_page_count, int)
        and page_count > compared_page_count
    )


def _source_first_sheet_hidden(result_dir: Path) -> bool:
    candidates = sorted(result_dir.glob("after-*.xlsx"))
    if not candidates:
        return False
    try:
        with zipfile.ZipFile(candidates[0]) as zf:
            workbook_xml = zf.read("xl/workbook.xml")
    except (KeyError, OSError, zipfile.BadZipFile):
        return False
    try:
        root = ElementTree.fromstring(workbook_xml)
    except ElementTree.ParseError:
        return False
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    first_sheet = root.find("x:sheets/x:sheet", ns)
    if first_sheet is None:
        return False
    return first_sheet.attrib.get("state") in {"hidden", "veryHidden"}


def _inconclusive(fixture: str, message: str) -> CopySheetEquivalenceResult:
    return CopySheetEquivalenceResult(
        fixture=fixture,
        status="inconclusive",
        source_page=None,
        matched_page=None,
        compared_pages=[],
        normalized_rmse=None,
        sampled=False,
        message=message,
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("render_report", type=Path)
    parser.add_argument("--max-normalized-rmse", type=float, default=0.001)
    parser.add_argument("--timeout", type=int, default=30)
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when no duplicate render evidence is found or failures exist.",
    )
    args = parser.parse_args(argv)
    report = audit_copy_sheet_render_equivalence(
        args.render_report,
        max_normalized_rmse=args.max_normalized_rmse,
        timeout=args.timeout,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
