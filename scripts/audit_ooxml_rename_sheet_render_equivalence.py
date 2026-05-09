#!/usr/bin/env python3
"""Audit rendered equivalence for rename-first-sheet OOXML mutations.

The normal intentional render smoke proves Excel can export the renamed
workbook. This audit is stricter for the visually neutral rename operation: it
exports the matching before workbook, rasterizes the same sampled pages, and
checks that the rendered pages remain equivalent after the rename.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
from dataclasses import asdict, dataclass
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import run_ooxml_app_smoke  # noqa: E402
import run_ooxml_render_compare  # noqa: E402

MUTATION = "rename_first_sheet"
PASSING_RENDER_STATUSES = {"rendered", "sampled_rendered"}
SPREADSHEET_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}


@dataclass(frozen=True)
class RenameSheetEquivalenceResult:
    fixture: str
    status: str
    compared_pages: list[int]
    max_normalized_rmse: float | None
    sampled: bool
    message: str


def audit_rename_sheet_render_equivalence(
    render_report_path: Path,
    max_normalized_rmse: float = 0.001,
    timeout: int = 30,
) -> dict:
    payload = json.loads(render_report_path.read_text())
    compare_cmd = run_ooxml_render_compare._find_imagemagick_compare()
    pdftoppm = shutil.which("pdftoppm")
    pdfinfo = shutil.which("pdfinfo")
    render_engine = str(payload.get("render_engine", "libreoffice"))
    density = int(payload.get("density", 96))
    soffice = None
    if render_engine == "libreoffice":
        soffice = run_ooxml_app_smoke._find_libreoffice()
    results: list[RenameSheetEquivalenceResult] = []
    for result in payload.get("results", []):
        if result.get("mutation") != MUTATION:
            continue
        results.append(
            _audit_result(
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
        )

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


def _audit_result(
    result: dict,
    *,
    compare_cmd: tuple[str, ...] | None,
    pdftoppm: str | None,
    pdfinfo: str | None,
    render_engine: str,
    soffice: str | None,
    density: int,
    max_normalized_rmse: float,
    timeout: int,
) -> RenameSheetEquivalenceResult:
    fixture = str(result.get("fixture", ""))
    status = str(result.get("status", ""))
    if status not in PASSING_RENDER_STATUSES:
        return RenameSheetEquivalenceResult(
            fixture=fixture,
            status="failed",
            compared_pages=[],
            max_normalized_rmse=None,
            sampled=False,
            message=f"render result status is not passing: {status or '<missing>'}",
        )
    if compare_cmd is None:
        return _inconclusive(fixture, "ImageMagick compare is required")
    if pdftoppm is None:
        return _inconclusive(fixture, "pdftoppm is required")
    if pdfinfo is None:
        return _inconclusive(fixture, "pdfinfo is required")
    if render_engine == "libreoffice" and soffice is None:
        return _inconclusive(fixture, "soffice is required for LibreOffice rendering")
    if render_engine == "excel" and not Path(run_ooxml_app_smoke.EXCEL_APP).is_dir():
        return _inconclusive(fixture, "Microsoft Excel is required for Excel rendering")

    after_pdf = result.get("after_pdf")
    if not isinstance(after_pdf, str):
        return _inconclusive(fixture, "render result has no after_pdf")
    result_dir = Path(after_pdf).parent.parent
    before_xlsx = _before_workbook(result_dir)
    if before_xlsx is None:
        return _inconclusive(fixture, "matching before workbook is missing")
    after_pages = _after_pages(result_dir)
    compared_pages = [int(page) for page in result.get("compared_pages", [])]
    if not compared_pages:
        return _inconclusive(fixture, "render result has no compared_pages")
    missing_after_pages = [page for page in compared_pages if page not in after_pages]
    if missing_after_pages:
        return _inconclusive(
            fixture,
            f"after page image(s) missing: {missing_after_pages}",
        )

    try:
        before_pdf = run_ooxml_render_compare._export_pdf(
            render_engine,
            soffice,
            before_xlsx,
            result_dir / "before-equivalence-pdf",
            timeout,
        )
        before_page_count = run_ooxml_render_compare._pdf_page_count(before_pdf)
        after_page_count = result.get("page_count")
        if isinstance(after_page_count, int) and before_page_count != after_page_count:
            return RenameSheetEquivalenceResult(
                fixture=fixture,
                status="failed",
                compared_pages=compared_pages,
                max_normalized_rmse=None,
                sampled=_is_sampled(result),
                message=(
                    f"page-count mismatch: before={before_page_count} "
                    f"after={after_page_count}"
                ),
            )
        before_page_paths = run_ooxml_render_compare._rasterize_pdf_pages(
            pdftoppm,
            before_pdf,
            result_dir / "before-equivalence-pages",
            compared_pages,
            density,
            timeout,
        )
        max_rmse = 0.0
        for page, before_page in zip(compared_pages, before_page_paths):
            max_rmse = max(
                max_rmse,
                _page_rmse(before_page, after_pages[page], compare_cmd, timeout),
            )
    except Exception as exc:
        return RenameSheetEquivalenceResult(
            fixture=fixture,
            status="failed",
            compared_pages=compared_pages,
            max_normalized_rmse=None,
            sampled=_is_sampled(result),
            message=str(exc)[:1000],
        )

    if max_rmse > max_normalized_rmse:
        return RenameSheetEquivalenceResult(
            fixture=fixture,
            status="failed",
            compared_pages=compared_pages,
            max_normalized_rmse=max_rmse,
            sampled=_is_sampled(result),
            message=(
                f"render drift above threshold: max_normalized_rmse={max_rmse:.8f} "
                f"threshold={max_normalized_rmse:.8f}"
            ),
        )
    return RenameSheetEquivalenceResult(
        fixture=fixture,
        status="passed",
        compared_pages=compared_pages,
        max_normalized_rmse=max_rmse,
        sampled=_is_sampled(result),
        message=f"rename-sheet render equivalent: max_normalized_rmse={max_rmse:.8f}",
    )


def _before_workbook(result_dir: Path) -> Path | None:
    candidates = [
        path
        for path in sorted(result_dir.glob("before-*"))
        if path.suffix.lower() in SPREADSHEET_SUFFIXES
    ]
    return candidates[0] if candidates else None


def _after_pages(result_dir: Path) -> dict[int, Path]:
    pages: dict[int, Path] = {}
    for path in result_dir.glob("after-pages-*.png"):
        parts = path.stem.split("-")
        if len(parts) < 3:
            continue
        try:
            page = int(parts[2])
        except ValueError:
            continue
        pages[page] = path
    return pages


def _page_rmse(
    before_page: Path,
    after_page: Path,
    compare_cmd: tuple[str, ...],
    timeout: int,
) -> float:
    if _sha256(before_page) == _sha256(after_page):
        return 0.0
    return run_ooxml_render_compare._normalized_rmse(
        compare_cmd,
        before_page,
        after_page,
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


def _inconclusive(fixture: str, message: str) -> RenameSheetEquivalenceResult:
    return RenameSheetEquivalenceResult(
        fixture=fixture,
        status="inconclusive",
        compared_pages=[],
        max_normalized_rmse=None,
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
        help="Exit non-zero when rename-sheet render equivalence is not proven.",
    )
    args = parser.parse_args(argv)
    report = audit_rename_sheet_render_equivalence(
        args.render_report,
        max_normalized_rmse=args.max_normalized_rmse,
        timeout=args.timeout,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
