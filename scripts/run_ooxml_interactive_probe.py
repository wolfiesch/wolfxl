#!/usr/bin/env python3
"""Run narrow interactive Excel probes for OOXML fidelity fixtures."""

from __future__ import annotations

import argparse
import fnmatch
import json
import subprocess
import shutil
import sys
import time
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_ooxml_fidelity  # noqa: E402
import audit_ooxml_fidelity_coverage  # noqa: E402
import run_ooxml_app_smoke  # noqa: E402
import run_ooxml_fidelity_mutations  # noqa: E402

SOURCE_MUTATION = run_ooxml_app_smoke.SOURCE_MUTATION
PASSING_STATUSES = {"passed"}
PRESENCE_PROBE_KIND = "ooxml_state_presence"
UI_INTERACTION_PROBE_KIND = "excel_ui_interaction"
PROBE_KIND = PRESENCE_PROBE_KIND
SUPPORTED_PROBE_KINDS = (PRESENCE_PROBE_KIND, UI_INTERACTION_PROBE_KIND)
SUPPORTED_PROBES = (
    "macro_project_presence",
    "embedded_control_openability",
    "external_link_update_prompt",
    "pivot_refresh_state",
    "slicer_selection_state",
    "timeline_selection_state",
)
PROBE_FEATURE_KEYS = {
    "macro_project_presence": "vba",
    "embedded_control_openability": "embedded_object",
    "external_link_update_prompt": "external_link",
    "pivot_refresh_state": "pivot",
    "slicer_selection_state": "slicer",
    "timeline_selection_state": "timeline",
}
SUPPORTED_UI_INTERACTION_PROBES = (
    "macro_project_presence",
    "external_link_update_prompt",
    "pivot_refresh_state",
    "slicer_selection_state",
    "timeline_selection_state",
)
REQUIRED_UI_ACTIONS = {
    "macro_project_presence": "clicked button: Disable Macros",
    "external_link_update_prompt": "clicked button: Don't Update",
    "pivot_refresh_state": "executed Excel command: refresh all",
    "slicer_selection_state": "selected Excel slicer shape",
    "timeline_selection_state": "selected Excel timeline shape",
}


@dataclass
class InteractiveProbeResult:
    fixture: str
    probe: str
    probe_kind: str
    mutation: str
    app: str
    status: str
    output: str | None
    message: str
    ui_actions: list[str] | None = None


def run_interactive_probes(
    fixture_dir: Path,
    output_dir: Path,
    probes: tuple[str, ...] = SUPPORTED_PROBES,
    mutation: str = SOURCE_MUTATION,
    timeout: int = 90,
    include_fixture_patterns: tuple[str, ...] = (),
    probe_kind: str = PROBE_KIND,
) -> dict:
    if probe_kind not in SUPPORTED_PROBE_KINDS:
        raise ValueError(f"unsupported probe kind: {probe_kind}")
    fixture_dir = fixture_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    results: list[InteractiveProbeResult] = []
    if probe_kind == UI_INTERACTION_PROBE_KIND:
        probes = tuple(probe for probe in probes if probe in SUPPORTED_UI_INTERACTION_PROBES)
    for entry in run_ooxml_fidelity_mutations.discover_fixtures(fixture_dir):
        if include_fixture_patterns and not _fixture_matches(
            entry.filename, include_fixture_patterns
        ):
            continue
        fixture_path = fixture_dir / entry.filename
        if not fixture_path.is_file():
            continue
        feature_keys = _feature_keys(fixture_path)
        for probe in probes:
            feature_key = PROBE_FEATURE_KEYS[probe]
            if feature_key not in feature_keys:
                continue
            results.append(
                _run_probe(
                    fixture_path,
                    entry.filename,
                    output_dir,
                    probe=probe,
                    mutation=mutation,
                    timeout=timeout,
                    probe_kind=probe_kind,
                )
            )
            _write_report(
                fixture_dir,
                output_dir,
                probes,
                mutation,
                results,
                include_fixture_patterns,
                probe_kind,
                completed=False,
            )

    return _write_report(
        fixture_dir,
        output_dir,
        probes,
        mutation,
        results,
        include_fixture_patterns,
        probe_kind,
        completed=True,
    )


def _fixture_matches(filename: str, patterns: tuple[str, ...]) -> bool:
    return any(fnmatch.fnmatch(filename, pattern) for pattern in patterns)


def _write_report(
    fixture_dir: Path,
    output_dir: Path,
    probes: tuple[str, ...],
    mutation: str,
    results: list[InteractiveProbeResult],
    include_fixture_patterns: tuple[str, ...],
    probe_kind: str,
    completed: bool,
) -> dict:
    report = {
        "completed": completed,
        "fixture_dir": str(fixture_dir),
        "output_dir": str(output_dir.resolve()),
        "probes": list(probes),
        "include_fixture_patterns": list(include_fixture_patterns),
        "probe_kind": probe_kind,
        "mutation": mutation,
        "result_count": len(results),
        "failure_count": sum(1 for result in results if result.status not in PASSING_STATUSES),
        "results": [asdict(result) for result in results],
    }
    (output_dir / "interactive-probe-report.json").write_text(
        json.dumps(report, indent=2, sort_keys=True)
    )
    return report


def _feature_keys(path: Path) -> set[str]:
    snapshot = audit_ooxml_fidelity.snapshot(path)
    return set(audit_ooxml_fidelity_coverage._feature_keys_for_snapshot(snapshot))


def _run_probe(
    fixture_path: Path,
    fixture_label: str,
    output_dir: Path,
    *,
    probe: str,
    mutation: str,
    timeout: int,
    probe_kind: str,
) -> InteractiveProbeResult:
    if probe not in SUPPORTED_PROBES:
        return InteractiveProbeResult(
            fixture=fixture_label,
            probe=probe,
            probe_kind=probe_kind,
            mutation=mutation,
            app="excel",
            status="failed",
            output=None,
            message=f"unsupported probe: {probe}",
        )
    if probe_kind == UI_INTERACTION_PROBE_KIND:
        return _run_ui_interaction_probe(
            fixture_path,
            fixture_label,
            output_dir,
            probe=probe,
            mutation=mutation,
            timeout=timeout,
        )

    work = (
        output_dir
        / run_ooxml_fidelity_mutations._safe_stem(Path(fixture_label).with_suffix("").as_posix())
        / mutation
    )
    work.mkdir(parents=True, exist_ok=True)
    probe_path = work / fixture_path.name
    shutil.copy2(fixture_path, probe_path)
    if mutation != SOURCE_MUTATION:
        prepared, mutation_error = run_ooxml_app_smoke._fixture_for_mutation(
            probe_path,
            fixture_label,
            work,
            mutation,
        )
        if mutation_error is not None:
            return InteractiveProbeResult(
                fixture=fixture_label,
                probe=probe,
                probe_kind=probe_kind,
                mutation=mutation,
                app="excel",
                status="failed",
                output=None,
                message=mutation_error,
            )
        probe_path = prepared

    if not _probe_part_present(probe_path, probe):
        return InteractiveProbeResult(
            fixture=fixture_label,
            probe=probe,
            probe_kind=probe_kind,
            mutation=mutation,
            app="excel",
            status="failed",
            output=str(probe_path),
            message=f"{_probe_part_label(probe)} missing before Excel open",
        )
    smoke = run_ooxml_app_smoke._smoke_excel(probe_path, work / "excel", timeout)
    if smoke.status != "passed":
        return InteractiveProbeResult(
            fixture=fixture_label,
            probe=probe,
            probe_kind=probe_kind,
            mutation=mutation,
            app="excel",
            status=smoke.status,
            output=smoke.output,
            message=smoke.message,
        )
    if not _probe_part_present(probe_path, probe):
        return InteractiveProbeResult(
            fixture=fixture_label,
            probe=probe,
            probe_kind=probe_kind,
            mutation=mutation,
            app="excel",
            status="failed",
            output=str(probe_path),
            message=f"{_probe_part_label(probe)} missing after Excel open",
        )
    part_label = _probe_part_label(probe)
    return InteractiveProbeResult(
        fixture=fixture_label,
        probe=probe,
        probe_kind=probe_kind,
        mutation=mutation,
        app="excel",
        status="passed",
        output=str(probe_path),
        message=f"Microsoft Excel opened workbook and {part_label} is present",
    )


def _run_ui_interaction_probe(
    fixture_path: Path,
    fixture_label: str,
    output_dir: Path,
    *,
    probe: str,
    mutation: str,
    timeout: int,
) -> InteractiveProbeResult:
    if probe not in SUPPORTED_UI_INTERACTION_PROBES:
        return InteractiveProbeResult(
            fixture=fixture_label,
            probe=probe,
            probe_kind=UI_INTERACTION_PROBE_KIND,
            mutation=mutation,
            app="excel",
            status="failed",
            output=None,
            message=f"probe {probe!r} does not have a click-level UI implementation",
            ui_actions=[],
        )

    work = (
        output_dir
        / run_ooxml_fidelity_mutations._safe_stem(Path(fixture_label).with_suffix("").as_posix())
        / mutation
        / UI_INTERACTION_PROBE_KIND
    )
    work.mkdir(parents=True, exist_ok=True)
    probe_path = work / fixture_path.name
    shutil.copy2(fixture_path, probe_path)
    if mutation != SOURCE_MUTATION:
        prepared, mutation_error = run_ooxml_app_smoke._fixture_for_mutation(
            probe_path,
            fixture_label,
            work,
            mutation,
        )
        if mutation_error is not None:
            return InteractiveProbeResult(
                fixture=fixture_label,
                probe=probe,
                probe_kind=UI_INTERACTION_PROBE_KIND,
                mutation=mutation,
                app="excel",
                status="failed",
                output=None,
                message=mutation_error,
                ui_actions=[],
            )
        probe_path = prepared

    if not _probe_part_present(probe_path, probe):
        return InteractiveProbeResult(
            fixture=fixture_label,
            probe=probe,
            probe_kind=UI_INTERACTION_PROBE_KIND,
            mutation=mutation,
            app="excel",
            status="failed",
            output=str(probe_path),
            message=f"{_probe_part_label(probe)} missing before Excel UI interaction",
            ui_actions=[],
        )

    try:
        active_name, ui_actions = _open_excel_with_ui_interaction(probe_path, probe, timeout)
    except Exception as exc:
        return InteractiveProbeResult(
            fixture=fixture_label,
            probe=probe,
            probe_kind=UI_INTERACTION_PROBE_KIND,
            mutation=mutation,
            app="excel",
            status="failed",
            output=str(probe_path),
            message=f"Excel UI interaction failed: {str(exc)[:500]}",
            ui_actions=[],
        )

    required_action = REQUIRED_UI_ACTIONS[probe]
    if required_action not in ui_actions:
        return InteractiveProbeResult(
            fixture=fixture_label,
            probe=probe,
            probe_kind=UI_INTERACTION_PROBE_KIND,
            mutation=mutation,
            app="excel",
            status="failed",
            output=str(probe_path),
            message=f"required UI action was not observed: {required_action}",
            ui_actions=ui_actions,
        )
    if not _probe_part_present(probe_path, probe):
        return InteractiveProbeResult(
            fixture=fixture_label,
            probe=probe,
            probe_kind=UI_INTERACTION_PROBE_KIND,
            mutation=mutation,
            app="excel",
            status="failed",
            output=str(probe_path),
            message=f"{_probe_part_label(probe)} missing after Excel UI interaction",
            ui_actions=ui_actions,
        )

    return InteractiveProbeResult(
        fixture=fixture_label,
        probe=probe,
        probe_kind=UI_INTERACTION_PROBE_KIND,
        mutation=mutation,
        app="excel",
        status="passed",
        output=str(probe_path),
        message=(
            "Microsoft Excel opened "
            f"{active_name}, completed required UI action, and "
            f"{_probe_part_label(probe)} is present"
        ),
        ui_actions=ui_actions,
    )


def _open_excel_with_ui_interaction(src: Path, probe: str, timeout: int) -> tuple[str, list[str]]:
    launched = subprocess.Popen(
        ["open", "-a", "Microsoft Excel", str(src.resolve())],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        start_new_session=True,
    )
    deadline = time.monotonic() + timeout
    ui_actions: list[str] = []
    last_error = ""
    while launched.poll() is None and time.monotonic() < deadline:
        ui_actions.extend(_perform_probe_ui_actions(probe))
        dialog = run_ooxml_app_smoke._excel_dialog_text()
        if run_ooxml_app_smoke._is_excel_unsupported_content_dialog(dialog):
            run_ooxml_app_smoke._dismiss_excel_unsupported_content_dialogs()
            run_ooxml_app_smoke._close_excel_best_effort()
            run_ooxml_app_smoke._quit_excel_best_effort()
            run_ooxml_app_smoke._kill_process_group_best_effort(launched)
            raise RuntimeError(f"Excel unsupported-content dialog: {dialog[:500]}")
        if run_ooxml_app_smoke._is_excel_repair_dialog(dialog):
            run_ooxml_app_smoke._dismiss_excel_repair_dialogs()
            run_ooxml_app_smoke._close_excel_best_effort()
            run_ooxml_app_smoke._quit_excel_best_effort()
            run_ooxml_app_smoke._kill_process_group_best_effort(launched)
            raise RuntimeError(f"Excel repair/error dialog while opening: {dialog[:500]}")
        last_error = dialog
        time.sleep(0.25)
    if launched.poll() is None:
        run_ooxml_app_smoke._close_excel_best_effort()
        run_ooxml_app_smoke._quit_excel_best_effort()
        run_ooxml_app_smoke._kill_process_group_best_effort(launched)
        raise subprocess.TimeoutExpired(
            ["open", "-a", "Microsoft Excel", str(src.resolve())],
            timeout,
            output=last_error,
        )
    stdout, stderr = launched.communicate()
    if launched.returncode != 0:
        raise RuntimeError(f"open -a Microsoft Excel failed: {stderr[:500]}")

    while time.monotonic() < deadline:
        ui_actions.extend(_perform_probe_ui_actions(probe))
        dialog = run_ooxml_app_smoke._excel_dialog_text()
        if run_ooxml_app_smoke._is_excel_unsupported_content_dialog(dialog):
            run_ooxml_app_smoke._dismiss_excel_unsupported_content_dialogs()
            run_ooxml_app_smoke._close_excel_best_effort()
            run_ooxml_app_smoke._quit_excel_best_effort()
            raise RuntimeError(f"Excel unsupported-content dialog: {dialog[:500]}")
        if run_ooxml_app_smoke._is_excel_repair_dialog(dialog):
            run_ooxml_app_smoke._dismiss_excel_repair_dialogs()
            run_ooxml_app_smoke._close_excel_best_effort()
            run_ooxml_app_smoke._quit_excel_best_effort()
            raise RuntimeError(f"Excel repair/error dialog while opening: {dialog[:500]}")
        name = run_ooxml_app_smoke._excel_active_workbook_name()
        if name:
            if probe == "pivot_refresh_state":
                ui_actions.extend(_refresh_all_active_workbook())
            if probe in {"slicer_selection_state", "timeline_selection_state"}:
                ui_actions.extend(_select_first_probe_shape(src, probe))
            run_ooxml_app_smoke._close_excel_best_effort()
            run_ooxml_app_smoke._quit_excel_best_effort()
            return name, _dedupe_actions(ui_actions)
        last_error = dialog
        time.sleep(0.5)
    run_ooxml_app_smoke._close_excel_best_effort()
    run_ooxml_app_smoke._quit_excel_best_effort()
    raise subprocess.TimeoutExpired(
        ["open", "-a", "Microsoft Excel", str(src.resolve())],
        timeout,
        output=last_error,
    )


def _perform_probe_ui_actions(probe: str) -> list[str]:
    if probe == "macro_project_presence":
        return _click_excel_button("Disable Macros")
    if probe == "external_link_update_prompt":
        return _click_excel_button("Don't Update")
    return []


def _click_excel_button(button: str) -> list[str]:
    script = f'''
tell application "System Events"
  if not (exists process "Microsoft Excel") then return ""
  tell process "Microsoft Excel"
    try
      if exists button "{button}" of window 1 then
        click button "{button}" of window 1
        return "clicked button: {button}"
      end if
    end try
  end tell
end tell
return ""
'''
    try:
        proc = run_ooxml_app_smoke._run_osascript(script, timeout=1)
    except subprocess.TimeoutExpired:
        return []
    action = proc.stdout.strip()
    return [action] if action else []


def _refresh_all_active_workbook() -> list[str]:
    script = """
tell application "Microsoft Excel"
  try
    refresh all active workbook
    return "executed Excel command: refresh all"
  on error errText
    return "failed Excel command: refresh all: " & errText
  end try
end tell
"""
    try:
        proc = run_ooxml_app_smoke._run_osascript(script, timeout=10)
    except subprocess.TimeoutExpired:
        return ["failed Excel command: refresh all: timeout"]
    action = proc.stdout.strip()
    return [action] if action else []


def _select_first_probe_shape(src: Path, probe: str) -> list[str]:
    names = _probe_shape_names(src, probe)
    if not names:
        return [f"failed Excel shape selection: no {probe} shape names in package"]
    quoted_names = ", ".join(f'"{_escape_applescript_text(name)}"' for name in names)
    label = "slicer" if probe == "slicer_selection_state" else "timeline"
    script = f"""
tell application "Microsoft Excel"
  try
    set expectedNames to {{{quoted_names}}}
    repeat with i from 1 to (count of shapes of active sheet)
      set candidate to shape i of active sheet
      set candidateName to name of candidate as text
      if expectedNames contains candidateName then
        select candidate
        return "selected Excel {label} shape" & linefeed & "selected Excel {label} shape: " & candidateName
      end if
    end repeat
    return "failed Excel {label} shape selection: no matching shape"
  on error errText
    return "failed Excel {label} shape selection: " & errText
  end try
end tell
"""
    try:
        proc = run_ooxml_app_smoke._run_osascript(script, timeout=10)
    except subprocess.TimeoutExpired:
        return [f"failed Excel {label} shape selection: timeout"]
    return [line for line in proc.stdout.splitlines() if line.strip()]


def _probe_shape_names(path: Path, probe: str) -> list[str]:
    if probe == "slicer_selection_state":
        prefix = "xl/slicers/"
        tag = "slicer"
    elif probe == "timeline_selection_state":
        prefix = "xl/timelines/"
        tag = "timeline"
    else:
        return []
    names: list[str] = []
    try:
        with zipfile.ZipFile(path) as archive:
            for part_name in archive.namelist():
                if not part_name.startswith(prefix) or not part_name.endswith(".xml"):
                    continue
                root = ET.fromstring(archive.read(part_name))
                for element in root.iter():
                    if _local_name(element.tag) == tag:
                        name = element.attrib.get("name")
                        if name and name not in names:
                            names.append(name)
    except (zipfile.BadZipFile, ET.ParseError, OSError):
        return []
    return names


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _escape_applescript_text(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def _dedupe_actions(actions: list[str]) -> list[str]:
    out: list[str] = []
    for action in actions:
        if action and action not in out:
            out.append(action)
    return out


def _probe_part_present(path: Path, probe: str) -> bool:
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
    except zipfile.BadZipFile:
        return False
    if probe == "macro_project_presence":
        return "xl/vbaProject.bin" in names
    if probe == "embedded_control_openability":
        return any(
            name.startswith(("xl/embeddings/", "xl/ctrlProps/", "xl/activeX/")) for name in names
        )
    if probe == "external_link_update_prompt":
        return any(name.startswith("xl/externalLinks/") for name in names)
    if probe == "pivot_refresh_state":
        return any(
            name.startswith(("xl/pivotCache/", "xl/pivotTables/", "pivotCache/")) for name in names
        )
    if probe == "slicer_selection_state":
        return any(name.startswith(("xl/slicers/", "xl/slicerCaches/")) for name in names)
    if probe == "timeline_selection_state":
        return any(name.startswith(("xl/timelines/", "xl/timelineCaches/")) for name in names)
    return False


def _probe_part_label(probe: str) -> str:
    if probe == "macro_project_presence":
        return "xl/vbaProject.bin"
    if probe == "embedded_control_openability":
        return "embedded/control OOXML parts"
    if probe == "external_link_update_prompt":
        return "external-link OOXML parts"
    if probe == "pivot_refresh_state":
        return "pivot cache/table OOXML parts"
    if probe == "slicer_selection_state":
        return "slicer/slicer-cache OOXML parts"
    if probe == "timeline_selection_state":
        return "timeline/timeline-cache OOXML parts"
    return "required OOXML parts"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("fixture_dir", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument(
        "--probe",
        action="append",
        choices=SUPPORTED_PROBES,
        dest="probes",
        help="Interactive probe to run. May be passed multiple times.",
    )
    parser.add_argument(
        "--mutation",
        choices=(SOURCE_MUTATION, *run_ooxml_fidelity_mutations.SUPPORTED_MUTATIONS),
        default=SOURCE_MUTATION,
    )
    parser.add_argument(
        "--probe-kind",
        choices=SUPPORTED_PROBE_KINDS,
        default=PROBE_KIND,
        help=(
            "Probe implementation to run. ooxml_state_presence opens/saves and checks "
            "OOXML parts; excel_ui_interaction records targeted Excel UI actions."
        ),
    )
    parser.add_argument("--timeout", type=int, default=90)
    parser.add_argument(
        "--fixture",
        action="append",
        default=[],
        dest="fixtures",
        help="Fixture filename or shell-style pattern to include. May be passed multiple times.",
    )
    args = parser.parse_args(argv)
    if args.probes:
        probes = tuple(args.probes)
    elif args.probe_kind == UI_INTERACTION_PROBE_KIND:
        probes = SUPPORTED_UI_INTERACTION_PROBES
    else:
        probes = SUPPORTED_PROBES

    report = run_interactive_probes(
        args.fixture_dir,
        args.output_dir,
        probes=probes,
        mutation=args.mutation,
        timeout=args.timeout,
        include_fixture_patterns=tuple(args.fixtures),
        probe_kind=args.probe_kind,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if report["failure_count"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
