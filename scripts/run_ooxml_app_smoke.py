#!/usr/bin/env python3
"""Open/save smoke tests for OOXML fidelity fixtures in real spreadsheet apps."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import run_ooxml_fidelity_mutations  # noqa: E402

LIBREOFFICE_CANDIDATES = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/opt/homebrew/bin/soffice",
    "/usr/local/bin/soffice",
    "/usr/bin/soffice",
)
EXCEL_APP = "/Applications/Microsoft Excel.app"
MANIFEST_NAME = "manifest.json"
SMOKE_KEYWORDS = ("corrupt", "repaired", "repair", "error")
PASSING_STATUSES = {"passed", "skipped"}
SOURCE_MUTATION = "source"


@dataclass
class AppSmokeResult:
    fixture: str
    mutation: str
    app: str
    status: str
    output: str | None
    message: str


def run_smoke(
    fixture_dir: Path,
    output_dir: Path,
    apps: Iterable[str],
    timeout: int = 90,
    mutations: Iterable[str] = (SOURCE_MUTATION,),
) -> dict:
    fixture_dir = fixture_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    results: list[AppSmokeResult] = []
    for entry in run_ooxml_fidelity_mutations.discover_fixtures(fixture_dir):
        fixture_path = fixture_dir / entry.filename
        if not fixture_path.is_file():
            results.append(
                AppSmokeResult(
                    fixture=entry.filename,
                    mutation=SOURCE_MUTATION,
                    app="discover",
                    status="failed",
                    output=None,
                    message=f"fixture missing: {fixture_path}",
                )
            )
            continue
        for mutation in mutations:
            smoke_path, mutation_error = _fixture_for_mutation(
                fixture_path,
                entry.filename,
                output_dir,
                mutation,
            )
            if mutation_error is not None:
                for app in apps:
                    results.append(
                        AppSmokeResult(
                            fixture=entry.filename,
                            mutation=mutation,
                            app=app,
                            status="failed",
                            output=None,
                            message=mutation_error,
                        )
                    )
                continue
            for app in apps:
                app_output_dir = output_dir / _safe_stem(mutation)
                if app == "libreoffice":
                    result = _smoke_libreoffice(smoke_path, app_output_dir, timeout)
                elif app == "excel":
                    result = _smoke_excel(smoke_path, app_output_dir, timeout)
                else:
                    raise ValueError(f"unknown app: {app}")
                result.fixture = entry.filename
                result.mutation = mutation
                results.append(result)

    report = {
        "fixture_dir": str(fixture_dir),
        "output_dir": str(output_dir.resolve()),
        "apps": list(apps),
        "mutations": list(mutations),
        "result_count": len(results),
        "failure_count": sum(
            1 for result in results if result.status not in PASSING_STATUSES
        ),
        "results": [asdict(result) for result in results],
    }
    (output_dir / "app-smoke-report.json").write_text(
        json.dumps(report, indent=2, sort_keys=True)
    )
    return report


def _smoke_libreoffice(src: Path, output_dir: Path, timeout: int) -> AppSmokeResult:
    soffice = _find_libreoffice()
    if soffice is None:
        return AppSmokeResult(
            src.name,
            SOURCE_MUTATION,
            "libreoffice",
            "skipped",
            None,
            "soffice not found",
        )
    work = output_dir / _safe_stem(src.stem) / "libreoffice"
    work.mkdir(parents=True, exist_ok=True)
    proc = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(work),
            str(src),
        ],
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    if proc.returncode != 0:
        return AppSmokeResult(
            src.name,
            SOURCE_MUTATION,
            "libreoffice",
            "failed",
            None,
            f"exit {proc.returncode}: {proc.stderr[:500]}",
        )
    stderr_lc = proc.stderr.lower()
    for keyword in SMOKE_KEYWORDS:
        if keyword in stderr_lc:
            return AppSmokeResult(
                src.name,
                SOURCE_MUTATION,
                "libreoffice",
                "failed",
                None,
                f"stderr contained {keyword!r}: {proc.stderr[:500]}",
            )
    converted = work / src.name
    ok, message = _validate_xlsx(converted)
    return AppSmokeResult(
        src.name,
        SOURCE_MUTATION,
        "libreoffice",
        "passed" if ok else "failed",
        str(converted) if converted.exists() else None,
        message,
    )


def _smoke_excel(src: Path, output_dir: Path, timeout: int) -> AppSmokeResult:
    if not Path(EXCEL_APP).is_dir():
        return AppSmokeResult(
            src.name,
            SOURCE_MUTATION,
            "excel",
            "skipped",
            None,
            "Microsoft Excel not found",
        )
    work = output_dir / _safe_stem(src.stem) / "excel"
    work.mkdir(parents=True, exist_ok=True)
    out = work / src.name
    out.unlink(missing_ok=True)
    script = f"""
tell application "Microsoft Excel"
  if (count of workbooks) > 0 then close every workbook saving no
  open workbook workbook file name (POSIX file "{_applescript_escape(str(src.resolve()))}") update links do not update links
  save workbook as active workbook filename (POSIX file "{_applescript_escape(str(out.resolve()))}") file format Excel XML file format
  close active workbook saving no
end tell
"""
    try:
        proc = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as exc:
        _close_excel_best_effort()
        return AppSmokeResult(
            src.name,
            SOURCE_MUTATION,
            "excel",
            "failed",
            str(out) if out.exists() else None,
            f"timeout after {timeout}s: {str(exc)[:500]}",
        )
    if proc.returncode != 0:
        return AppSmokeResult(
            src.name,
            SOURCE_MUTATION,
            "excel",
            "failed",
            None,
            f"exit {proc.returncode}: {proc.stderr[:500]}",
        )
    ok, message = _validate_xlsx(out)
    return AppSmokeResult(
        src.name,
        SOURCE_MUTATION,
        "excel",
        "passed" if ok else "failed",
        str(out) if out.exists() else None,
        message,
    )


def _validate_xlsx(path: Path) -> tuple[bool, str]:
    if not path.exists():
        return False, f"no output file at {path}"
    if not zipfile.is_zipfile(path):
        return False, "output is not a valid zip"
    try:
        with zipfile.ZipFile(path) as archive:
            bad = archive.testzip()
            if bad is not None:
                return False, f"zip integrity failure: {bad}"
            names = set(archive.namelist())
    except zipfile.BadZipFile as exc:
        return False, f"bad central directory: {exc}"
    required = {"[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"}
    missing = sorted(required - names)
    if missing:
        return False, f"missing required OOXML parts: {missing}"
    return True, "ok"


def _fixture_for_mutation(
    fixture_path: Path,
    fixture_label: str,
    output_dir: Path,
    mutation: str,
) -> tuple[Path, str | None]:
    if mutation == SOURCE_MUTATION:
        return fixture_path, None
    work = (
        output_dir
        / "_mutations"
        / run_ooxml_fidelity_mutations._safe_stem(
            Path(fixture_label).with_suffix("").as_posix()
        )
        / mutation
    )
    work.mkdir(parents=True, exist_ok=True)
    before_path = work / f"before-{fixture_path.name}"
    after_path = work / f"after-{fixture_path.name}"
    shutil.copy2(fixture_path, before_path)
    shutil.copy2(fixture_path, after_path)
    try:
        run_ooxml_fidelity_mutations._prepare_mutation_baseline(
            before_path,
            after_path,
            mutation,
        )
        run_ooxml_fidelity_mutations._apply_mutation(after_path, mutation)
    except Exception as exc:
        return after_path, f"mutation {mutation!r} failed: {str(exc)[:500]}"
    return after_path, None


def _find_libreoffice() -> str | None:
    for candidate in LIBREOFFICE_CANDIDATES:
        path = Path(candidate)
        if path.is_file() and os.access(path, os.X_OK):
            return str(path)
    return shutil.which("soffice")


def _safe_stem(stem: str) -> str:
    return "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in stem)


def _applescript_escape(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def _close_excel_best_effort() -> None:
    subprocess.run(
        [
            "osascript",
            "-e",
            (
                'tell application "Microsoft Excel"\n'
                "  if (count of workbooks) > 0 then close every workbook saving no\n"
                "end tell"
            ),
        ],
        capture_output=True,
        text=True,
        timeout=10,
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("fixture_dir", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument(
        "--app",
        action="append",
        choices=("libreoffice", "excel"),
        dest="apps",
        help="App smoke to run. May be passed multiple times.",
    )
    parser.add_argument("--timeout", type=int, default=90)
    parser.add_argument(
        "--mutation",
        action="append",
        choices=(SOURCE_MUTATION, *run_ooxml_fidelity_mutations.SUPPORTED_MUTATIONS),
        dest="mutations",
        help=(
            "Workbook mutation to smoke. May be passed multiple times. "
            "Defaults to source fixtures without mutation."
        ),
    )
    args = parser.parse_args(argv)

    apps = tuple(args.apps or ("libreoffice",))
    mutations = tuple(args.mutations or (SOURCE_MUTATION,))
    report = run_smoke(
        args.fixture_dir,
        args.output_dir,
        apps,
        timeout=args.timeout,
        mutations=mutations,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if report["failure_count"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
