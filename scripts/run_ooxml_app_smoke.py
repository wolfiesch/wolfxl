#!/usr/bin/env python3
"""Open/save smoke tests for OOXML fidelity fixtures in real spreadsheet apps."""

from __future__ import annotations

import argparse
import json
import os
import signal
import shutil
import subprocess
import sys
import time
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
REPAIR_DISMISS_BUTTONS = (
    "No",
    "Cancel",
    "Don't Recover",
    "Delete",
    "OK",
    "Close",
)
PASSING_STATUSES = {"passed", "skipped"}
SOURCE_MUTATION = "source"
EXCEL_REPAIR_MARKER = "Excel repair/error dialog while opening:"


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
    stop_on_excel_repair: bool = True,
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
                if (
                    stop_on_excel_repair
                    and app == "excel"
                    and result.status not in PASSING_STATUSES
                    and EXCEL_REPAIR_MARKER in result.message
                ):
                    return _write_report(
                        fixture_dir,
                        output_dir,
                        apps,
                        mutations,
                        results,
                        aborted=True,
                        abort_reason=(
                            "stopped after first Microsoft Excel repair dialog; "
                            "rerun with --continue-after-excel-repair to collect "
                            "additional failures"
                        ),
                    )

    return _write_report(
        fixture_dir,
        output_dir,
        apps,
        mutations,
        results,
        aborted=False,
        abort_reason=None,
    )


def _write_report(
    fixture_dir: Path,
    output_dir: Path,
    apps: Iterable[str],
    mutations: Iterable[str],
    results: list[AppSmokeResult],
    *,
    aborted: bool,
    abort_reason: str | None,
) -> dict:
    report = {
        "fixture_dir": str(fixture_dir),
        "output_dir": str(output_dir.resolve()),
        "apps": list(apps),
        "mutations": list(mutations),
        "aborted": aborted,
        "abort_reason": abort_reason,
        "result_count": len(results),
        "failure_count": sum(1 for result in results if result.status not in PASSING_STATUSES),
        "results": [asdict(result) for result in results],
    }
    (output_dir / "app-smoke-report.json").write_text(json.dumps(report, indent=2, sort_keys=True))
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
    converted = _libreoffice_converted_path(work, src)
    ok, message = _validate_xlsx(converted)
    return AppSmokeResult(
        src.name,
        SOURCE_MUTATION,
        "libreoffice",
        "passed" if ok else "failed",
        str(converted) if converted.exists() else None,
        message,
    )


def _libreoffice_converted_path(work: Path, src: Path) -> Path:
    expected = work / src.name
    if expected.exists():
        return expected
    xlsx = work / f"{src.stem}.xlsx"
    if xlsx.exists():
        return xlsx
    matches = sorted(work.glob(f"{src.stem}.*"))
    return matches[0] if matches else expected


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
    try:
        opened = _open_excel_with_finder_and_close(
            src,
            timeout,
        )
    except subprocess.TimeoutExpired as exc:
        _close_excel_best_effort()
        return AppSmokeResult(
            src.name,
            SOURCE_MUTATION,
            "excel",
            "failed",
            str(src),
            f"timeout after {timeout}s: {str(exc)[:500]}",
        )
    except RuntimeError as exc:
        _dismiss_excel_repair_dialogs()
        _close_excel_best_effort()
        _quit_excel_best_effort()
        return AppSmokeResult(
            src.name,
            SOURCE_MUTATION,
            "excel",
            "failed",
            None,
            str(exc)[:500],
        )
    if opened != src.name:
        return AppSmokeResult(
            src.name,
            SOURCE_MUTATION,
            "excel",
            "failed",
            str(src),
            f"Microsoft Excel opened {opened!r}, expected {src.name!r}",
        )
    ok, message = _validate_xlsx(src)
    if ok:
        message = f"opened and closed in Microsoft Excel: {opened or src.name}"
    return AppSmokeResult(
        src.name,
        SOURCE_MUTATION,
        "excel",
        "passed" if ok else "failed",
        str(src),
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
        / run_ooxml_fidelity_mutations._safe_stem(Path(fixture_label).with_suffix("").as_posix())
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


def _open_excel_with_finder_and_close(src: Path, timeout: int) -> str:
    # Finder-style open avoids Office's AppleScript sandbox prompt for generated
    # files; AppleScript is still used for the observable close/no-save step.
    launched = subprocess.Popen(
        ["open", "-a", "Microsoft Excel", str(src.resolve())],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        start_new_session=True,
    )

    deadline = time.monotonic() + timeout
    last_error = ""
    while launched.poll() is None and time.monotonic() < deadline:
        _dismiss_excel_safe_dialogs()
        dialog = _excel_dialog_text()
        if _is_excel_repair_dialog(dialog):
            _dismiss_excel_repair_dialogs()
            _close_excel_best_effort()
            _quit_excel_best_effort()
            _kill_process_group_best_effort(launched)
            raise RuntimeError(f"Excel repair/error dialog while opening: {dialog[:500]}")
        last_error = dialog
        time.sleep(0.25)
    if launched.poll() is None:
        _dismiss_excel_repair_dialogs()
        _close_excel_best_effort()
        _quit_excel_best_effort()
        _kill_process_group_best_effort(launched)
        raise subprocess.TimeoutExpired(
            ["open", "-a", "Microsoft Excel", str(src.resolve())],
            timeout,
            output=last_error,
        )
    stdout, stderr = launched.communicate()
    if launched.returncode != 0:
        raise RuntimeError(f"open -a Microsoft Excel failed: {stderr[:500]}")

    while time.monotonic() < deadline:
        _dismiss_excel_safe_dialogs()
        dialog = _excel_dialog_text()
        if _is_excel_repair_dialog(dialog):
            _dismiss_excel_repair_dialogs()
            _close_excel_best_effort()
            _quit_excel_best_effort()
            raise RuntimeError(f"Excel repair/error dialog while opening: {dialog[:500]}")
        name = _excel_active_workbook_name()
        if name:
            _close_excel_best_effort()
            _quit_excel_best_effort()
            return name
        last_error = dialog
        time.sleep(0.5)
    _dismiss_excel_repair_dialogs()
    _close_excel_best_effort()
    _quit_excel_best_effort()
    raise subprocess.TimeoutExpired(
        ["open", "-a", "Microsoft Excel", str(src.resolve())],
        timeout,
        output=last_error,
    )


def _is_excel_repair_dialog(dialog: str) -> bool:
    dialog_lc = dialog.lower()
    return any(keyword in dialog_lc for keyword in SMOKE_KEYWORDS)


def _kill_process_group_best_effort(proc: subprocess.Popen[str]) -> None:
    killpg = getattr(os, "killpg", None)
    kill_signal = getattr(signal, "SIGKILL", None)
    try:
        if killpg is None or kill_signal is None:
            raise PermissionError("process groups are not available")
        killpg(proc.pid, kill_signal)
    except (PermissionError, ProcessLookupError):
        try:
            proc.kill()
        except (PermissionError, ProcessLookupError):
            pass


def _excel_active_workbook_name() -> str | None:
    script = """
tell application "Microsoft Excel"
  try
    return name of active workbook as text
  on error
    return ""
  end try
end tell
"""
    proc = _run_osascript(script, timeout=2)
    name = proc.stdout.strip()
    if name == "missing value":
        return None
    return name or None


def _excel_dialog_text() -> str:
    script = """
tell application "System Events"
  if not (exists process "Microsoft Excel") then return ""
  tell process "Microsoft Excel"
    try
      set win_names to name of windows
      set button_names to name of buttons of window 1
      set static_values to value of static texts of window 1
      return "windows=" & (win_names as text) & "\nbuttons=" & (button_names as text) & "\ntext=" & (static_values as text)
    on error
      return ""
    end try
  end tell
end tell
"""
    try:
        proc = _run_osascript(script, timeout=1)
    except subprocess.TimeoutExpired:
        return ""
    return proc.stdout.strip()


def _dismiss_excel_safe_dialogs() -> None:
    # Macro-enabled fixtures can block AppleScript until the security prompt is
    # answered. Disable macros for smoke validation: we only need to prove Excel
    # can open the workbook without repair/corruption prompts.
    script = """
tell application "System Events"
  if not (exists process "Microsoft Excel") then return
  tell process "Microsoft Excel"
    try
      if exists button "Disable Macros" of window 1 then click button "Disable Macros" of window 1
    end try
    try
      if exists button "Don't Update" of window 1 then click button "Don't Update" of window 1
    end try
  end tell
end tell
"""
    try:
        _run_osascript(script, timeout=1)
    except subprocess.TimeoutExpired:
        # Excel can briefly stop responding to UI automation while opening large
        # pivot/slicer workbooks. Keep the primary open/close attempt alive.
        pass


def _dismiss_excel_repair_dialogs() -> None:
    # Treat repair/recovery prompts as a failure signal and choose the
    # non-repairing/default-dismiss path. Do not click "Yes" or "Recover":
    # accepting repair can mutate the workbook and hide the OOXML defect.
    buttons = "\n".join(
        f'        if exists button "{button}" of w then click button "{button}" of w'
        for button in REPAIR_DISMISS_BUTTONS
    )
    script = f"""
tell application "System Events"
  if not (exists process "Microsoft Excel") then return
  tell process "Microsoft Excel"
    set frontmost to true
    try
      repeat with w in windows
{buttons}
      end repeat
    end try
  end tell
end tell
"""
    try:
        _run_osascript(script, timeout=1)
    except subprocess.TimeoutExpired:
        pass


def _close_excel_best_effort() -> None:
    try:
        _run_osascript(
            (
                'tell application "Microsoft Excel"\n'
                "  try\n"
                "    close active workbook saving no\n"
                "  end try\n"
                "end tell"
            ),
            timeout=3,
        )
    except subprocess.TimeoutExpired:
        _quit_excel_best_effort()


def _quit_excel_best_effort() -> None:
    try:
        _run_osascript(
            ('tell application "Microsoft Excel"\n  try\n    quit saving no\n  end try\nend tell'),
            timeout=3,
        )
    except subprocess.TimeoutExpired:
        subprocess.run(
            ["pkill", "-x", "Microsoft Excel"],
            capture_output=True,
            text=True,
            timeout=5,
        )


def _run_osascript(script: str, timeout: int) -> subprocess.CompletedProcess[str]:
    proc = subprocess.Popen(
        ["osascript", "-e", script],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        start_new_session=True,
    )
    try:
        stdout, stderr = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        _kill_process_group_best_effort(proc)
        try:
            stdout, stderr = proc.communicate(timeout=1)
        except subprocess.TimeoutExpired:
            stdout = exc.output or ""
            stderr = exc.stderr or ""
        raise subprocess.TimeoutExpired(
            proc.args,
            timeout,
            output=stdout,
            stderr=stderr,
        ) from exc
    return subprocess.CompletedProcess(proc.args, proc.returncode, stdout, stderr)


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
    parser.add_argument(
        "--continue-after-excel-repair",
        action="store_true",
        help=(
            "For Microsoft Excel GUI smoke only, keep running after a repair "
            "dialog failure. By default the run aborts on the first repair "
            "dialog to avoid repeated desktop popups."
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
        stop_on_excel_repair=not args.continue_after_excel_repair,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if report["failure_count"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
