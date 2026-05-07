#!/usr/bin/env python3
"""Rendered-output comparison for OOXML fidelity fixtures.

For each fixture, this script performs a WolfXL no-op modify-save, exports the
original and saved workbook to PDF through LibreOffice, rasterizes the PDFs,
and compares corresponding page images with ImageMagick's RMSE metric.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from dataclasses import asdict, dataclass
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import run_ooxml_app_smoke  # noqa: E402
import run_ooxml_fidelity_mutations  # noqa: E402
import wolfxl  # noqa: E402

PASSING_STATUSES = {"passed", "sampled_passed", "skipped"}
RENDER_KEYWORDS = ("corrupt", "repaired", "repair", "error")
RMSE_RE = re.compile(r"\((?P<normalized>[0-9.]+(?:e[+-]?[0-9]+)?)\)", re.I)


@dataclass
class RenderCompareResult:
    fixture: str
    status: str
    before_pdf: str | None
    after_pdf: str | None
    page_count: int | None
    compared_page_count: int | None
    compared_pages: list[int]
    max_normalized_rmse: float | None
    message: str


def run_render_compare(
    fixture_dir: Path,
    output_dir: Path,
    timeout: int = 90,
    density: int = 96,
    max_normalized_rmse: float = 0.001,
    recursive: bool = False,
    max_pages_per_fixture: int | None = None,
    pass_byte_identical_xlsx: bool = False,
) -> dict:
    fixture_dir = fixture_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    results: list[RenderCompareResult] = []
    for entry in run_ooxml_fidelity_mutations.discover_fixtures(
        fixture_dir, recursive=recursive
    ):
        fixture_path = fixture_dir / entry.filename
        results.append(
            _compare_fixture(
                fixture_path,
                entry.filename,
                entry.sha256,
                output_dir,
                timeout=timeout,
                density=density,
                max_normalized_rmse=max_normalized_rmse,
                max_pages_per_fixture=max_pages_per_fixture,
                pass_byte_identical_xlsx=pass_byte_identical_xlsx,
            )
        )

    report = {
        "fixture_dir": str(fixture_dir),
        "output_dir": str(output_dir.resolve()),
        "density": density,
        "max_normalized_rmse_threshold": max_normalized_rmse,
        "max_pages_per_fixture": max_pages_per_fixture,
        "pass_byte_identical_xlsx": pass_byte_identical_xlsx,
        "recursive": recursive,
        "result_count": len(results),
        "failure_count": sum(
            1 for result in results if result.status not in PASSING_STATUSES
        ),
        "results": [asdict(result) for result in results],
    }
    (output_dir / "render-compare-report.json").write_text(
        json.dumps(report, indent=2, sort_keys=True)
    )
    return report


def _compare_fixture(
    fixture_path: Path,
    fixture_label: str,
    expected_sha256: str | None,
    output_dir: Path,
    timeout: int,
    density: int,
    max_normalized_rmse: float,
    max_pages_per_fixture: int | None,
    pass_byte_identical_xlsx: bool,
) -> RenderCompareResult:
    if not fixture_path.is_file():
        return RenderCompareResult(
            fixture=fixture_label,
            status="failed",
            before_pdf=None,
            after_pdf=None,
            page_count=None,
            compared_page_count=None,
            compared_pages=[],
            max_normalized_rmse=None,
            message=f"fixture missing: {fixture_path}",
        )
    if expected_sha256:
        actual_sha256 = hashlib.sha256(fixture_path.read_bytes()).hexdigest()
        if actual_sha256 != expected_sha256:
            return RenderCompareResult(
                fixture=fixture_label,
                status="failed",
                before_pdf=None,
                after_pdf=None,
                page_count=None,
                compared_page_count=None,
                compared_pages=[],
                max_normalized_rmse=None,
                message=(
                    f"sha256 mismatch: expected {expected_sha256}, "
                    f"got {actual_sha256}"
                ),
            )

    soffice = run_ooxml_app_smoke._find_libreoffice()
    pdftoppm = shutil.which("pdftoppm")
    compare_cmd = _find_imagemagick_compare()
    if soffice is None:
        return _skipped(fixture_label, "soffice not found")
    if pdftoppm is None:
        return _skipped(fixture_label, "pdftoppm not found")
    if compare_cmd is None:
        return _skipped(fixture_label, "ImageMagick compare not found")

    work = output_dir / run_ooxml_fidelity_mutations._safe_stem(
        Path(fixture_label).with_suffix("").as_posix()
    )
    work.mkdir(parents=True, exist_ok=True)
    before_xlsx = work / f"before-{fixture_path.name}"
    after_xlsx = work / f"after-{fixture_path.name}"
    shutil.copy2(fixture_path, before_xlsx)
    shutil.copy2(fixture_path, after_xlsx)

    try:
        workbook = wolfxl.load_workbook(after_xlsx, modify=True)
        try:
            workbook.save(after_xlsx)
        finally:
            close = getattr(workbook, "close", None)
            if close is not None:
                close()

        if pass_byte_identical_xlsx and _files_identical(before_xlsx, after_xlsx):
            return RenderCompareResult(
                fixture=fixture_label,
                status="passed",
                before_pdf=None,
                after_pdf=None,
                page_count=None,
                compared_page_count=None,
                compared_pages=[],
                max_normalized_rmse=0.0,
                message="byte-identical xlsx after no-op save; render equivalence inferred",
            )

        before_pdf = _export_pdf(soffice, before_xlsx, work / "before-pdf", timeout)
        after_pdf = _export_pdf(soffice, after_xlsx, work / "after-pdf", timeout)
        before_page_count = _pdf_page_count(before_pdf)
        after_page_count = _pdf_page_count(after_pdf)
        if before_page_count != after_page_count:
            return RenderCompareResult(
                fixture=fixture_label,
                status="failed",
                before_pdf=str(before_pdf),
                after_pdf=str(after_pdf),
                page_count=None,
                compared_page_count=None,
                compared_pages=[],
                max_normalized_rmse=None,
                message=(
                    f"page-count mismatch: before={before_page_count} "
                    f"after={after_page_count}"
                ),
            )
        sampled = (
            max_pages_per_fixture is not None
            and before_page_count > max_pages_per_fixture
        )
        compared_pages = (
            _sample_page_numbers(before_page_count, max_pages_per_fixture)
            if sampled
            else list(range(1, before_page_count + 1))
        )
        before_pages = _rasterize_pdf_pages(
            pdftoppm,
            before_pdf,
            work / "before-pages",
            compared_pages,
            density,
            timeout,
        )
        after_pages = _rasterize_pdf_pages(
            pdftoppm,
            after_pdf,
            work / "after-pages",
            compared_pages,
            density,
            timeout,
        )
        max_rmse = 0.0
        for before_page, after_page in zip(before_pages, after_pages, strict=True):
            max_rmse = max(
                max_rmse,
                _normalized_rmse(compare_cmd, before_page, after_page, timeout),
            )
    except Exception as exc:
        return RenderCompareResult(
            fixture=fixture_label,
            status="failed",
            before_pdf=None,
            after_pdf=None,
            page_count=None,
            compared_page_count=None,
            compared_pages=[],
            max_normalized_rmse=None,
            message=str(exc)[:1000],
        )

    if max_rmse > max_normalized_rmse:
        status = "failed"
        message = (
            f"render drift above threshold: max_normalized_rmse={max_rmse:.8f} "
            f"threshold={max_normalized_rmse:.8f}"
        )
    elif sampled:
        status = "sampled_passed"
        message = (
            f"sampled ok: compared {len(compared_pages)} of {before_page_count} pages; "
            f"max_normalized_rmse={max_rmse:.8f}"
        )
    else:
        status = "passed"
        message = f"ok: max_normalized_rmse={max_rmse:.8f}"
    return RenderCompareResult(
        fixture=fixture_label,
        status=status,
        before_pdf=str(before_pdf),
        after_pdf=str(after_pdf),
        page_count=before_page_count,
        compared_page_count=len(compared_pages),
        compared_pages=compared_pages,
        max_normalized_rmse=max_rmse,
        message=message,
    )


def _skipped(fixture: str, message: str) -> RenderCompareResult:
    return RenderCompareResult(
        fixture=fixture,
        status="skipped",
        before_pdf=None,
        after_pdf=None,
        page_count=None,
        compared_page_count=None,
        compared_pages=[],
        max_normalized_rmse=None,
        message=message,
    )


def _export_pdf(soffice: str, src: Path, outdir: Path, timeout: int) -> Path:
    outdir.mkdir(parents=True, exist_ok=True)
    proc = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(outdir),
            str(src),
        ],
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            f"LibreOffice PDF export failed for {src.name}: "
            f"exit {proc.returncode}: {proc.stderr[:500]}"
        )
    stderr_lc = proc.stderr.lower()
    for keyword in RENDER_KEYWORDS:
        if keyword in stderr_lc:
            raise RuntimeError(
                f"LibreOffice PDF export stderr contained {keyword!r}: "
                f"{proc.stderr[:500]}"
            )
    pdf = outdir / f"{src.stem}.pdf"
    if not pdf.is_file() or pdf.stat().st_size == 0:
        raise RuntimeError(f"LibreOffice did not produce a non-empty PDF at {pdf}")
    return pdf


def _files_identical(left: Path, right: Path) -> bool:
    if left.stat().st_size != right.stat().st_size:
        return False
    chunk_size = 1024 * 1024
    with left.open("rb") as left_file, right.open("rb") as right_file:
        while True:
            left_chunk = left_file.read(chunk_size)
            right_chunk = right_file.read(chunk_size)
            if left_chunk != right_chunk:
                return False
            if not left_chunk:
                return True


def _rasterize_pdf(
    pdftoppm: str,
    pdf: Path,
    prefix: Path,
    density: int,
    timeout: int,
) -> list[Path]:
    prefix.parent.mkdir(parents=True, exist_ok=True)
    for stale in prefix.parent.glob(f"{prefix.name}-*.png"):
        stale.unlink()
    proc = subprocess.run(
        [pdftoppm, "-png", "-r", str(density), str(pdf), str(prefix)],
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            f"pdftoppm failed for {pdf.name}: exit {proc.returncode}: {proc.stderr[:500]}"
        )
    pages = sorted(prefix.parent.glob(f"{prefix.name}-*.png"))
    if not pages:
        raise RuntimeError(f"pdftoppm produced no page images for {pdf}")
    return pages


def _rasterize_pdf_pages(
    pdftoppm: str,
    pdf: Path,
    prefix: Path,
    pages: list[int],
    density: int,
    timeout: int,
) -> list[Path]:
    if not pages:
        raise RuntimeError(f"no pages selected for {pdf}")
    if pages == list(range(1, len(pages) + 1)):
        return _rasterize_pdf(pdftoppm, pdf, prefix, density, timeout)

    prefix.parent.mkdir(parents=True, exist_ok=True)
    for stale in prefix.parent.glob(f"{prefix.name}-*.png"):
        stale.unlink()
    out: list[Path] = []
    for page in pages:
        page_prefix = prefix.parent / f"{prefix.name}-{page}"
        proc = subprocess.run(
            [
                pdftoppm,
                "-png",
                "-r",
                str(density),
                "-f",
                str(page),
                "-l",
                str(page),
                str(pdf),
                str(page_prefix),
            ],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        if proc.returncode != 0:
            raise RuntimeError(
                f"pdftoppm failed for {pdf.name} page {page}: "
                f"exit {proc.returncode}: {proc.stderr[:500]}"
            )
        rendered = sorted(page_prefix.parent.glob(f"{page_prefix.name}-*.png"))
        if not rendered:
            raise RuntimeError(f"pdftoppm produced no page image for {pdf} page {page}")
        out.append(rendered[-1])
    return out


def _pdf_page_count(pdf: Path) -> int:
    pdfinfo = shutil.which("pdfinfo")
    if pdfinfo is None:
        raise RuntimeError("pdfinfo not found")
    proc = subprocess.run(
        [pdfinfo, str(pdf)],
        capture_output=True,
        text=True,
        timeout=30,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            f"pdfinfo failed for {pdf.name}: exit {proc.returncode}: {proc.stderr[:500]}"
        )
    match = re.search(r"^Pages:\s+(\d+)\s*$", proc.stdout, re.MULTILINE)
    if match is None:
        raise RuntimeError(f"could not parse pdfinfo page count for {pdf}")
    return int(match.group(1))


def _sample_page_numbers(page_count: int, max_pages: int | None) -> list[int]:
    if max_pages is None or page_count <= max_pages:
        return list(range(1, page_count + 1))
    if max_pages <= 0:
        raise ValueError("max_pages_per_fixture must be positive")
    if max_pages == 1:
        return [1]
    if max_pages == 2:
        return [1, page_count]
    step = (page_count - 1) / (max_pages - 1)
    pages = {1, page_count}
    for idx in range(max_pages):
        pages.add(round(1 + idx * step))
    return sorted(pages)


def _normalized_rmse(
    compare_cmd: tuple[str, ...],
    before_page: Path,
    after_page: Path,
    timeout: int,
) -> float:
    cmd = [
        *compare_cmd,
        "-metric",
        "RMSE",
        str(before_page),
        str(after_page),
        "null:",
    ]
    proc = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    if proc.returncode not in (0, 1):
        raise RuntimeError(
            f"ImageMagick compare failed for {before_page.name}: "
            f"exit {proc.returncode}: {proc.stderr[:500]}"
        )
    metric = proc.stderr.strip()
    match = RMSE_RE.search(metric)
    if match is None:
        raise RuntimeError(f"could not parse ImageMagick RMSE metric: {metric!r}")
    return float(match.group("normalized"))


def _find_imagemagick_compare() -> tuple[str, ...] | None:
    if path := shutil.which("compare"):
        return (path,)
    if path := shutil.which("magick"):
        return (path, "compare")
    return None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("fixture_dir", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--timeout", type=int, default=90)
    parser.add_argument("--density", type=int, default=96)
    parser.add_argument("--max-normalized-rmse", type=float, default=0.001)
    parser.add_argument(
        "--max-pages-per-fixture",
        type=int,
        default=None,
        help=(
            "When a rendered PDF has more pages than this, compare a deterministic "
            "sample instead of rasterizing every page."
        ),
    )
    parser.add_argument(
        "--pass-byte-identical-xlsx",
        action="store_true",
        help=(
            "Return pass before rendering when the no-op saved workbook is "
            "byte-identical to the original copy."
        ),
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Discover .xlsx fixtures recursively when no manifest.json is present.",
    )
    args = parser.parse_args(argv)

    report = run_render_compare(
        args.fixture_dir,
        args.output_dir,
        timeout=args.timeout,
        density=args.density,
        max_normalized_rmse=args.max_normalized_rmse,
        recursive=args.recursive,
        max_pages_per_fixture=args.max_pages_per_fixture,
        pass_byte_identical_xlsx=args.pass_byte_identical_xlsx,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if report["failure_count"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
