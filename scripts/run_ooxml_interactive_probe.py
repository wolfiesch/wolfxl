#!/usr/bin/env python3
"""Run narrow interactive Excel probes for OOXML fidelity fixtures."""

from __future__ import annotations

import argparse
import datetime as dt
import fnmatch
import json
import os
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
EXTERNAL_LINK_PROMPT_MODE_FORCE = "force"
EXTERNAL_LINK_PROMPT_MODE_CURRENT = "current"
SUPPORTED_EXTERNAL_LINK_PROMPT_MODES = (
    EXTERNAL_LINK_PROMPT_MODE_FORCE,
    EXTERNAL_LINK_PROMPT_MODE_CURRENT,
)
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
    "embedded_control_openability",
    "external_link_update_prompt",
    "pivot_refresh_state",
    "slicer_selection_state",
    "timeline_selection_state",
)
REQUIRED_UI_ACTIONS = {
    "macro_project_presence": "clicked button: Disable Macros",
    "embedded_control_openability": "clicked Excel embedded/control object",
    "external_link_update_prompt": "clicked button: Don't Update",
    "pivot_refresh_state": "executed Excel command: refresh all",
    "slicer_selection_state": "clicked Excel slicer item",
    "timeline_selection_state": "clicked Excel timeline month",
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
    external_link_prompt_mode: str = EXTERNAL_LINK_PROMPT_MODE_FORCE,
) -> dict:
    if probe_kind not in SUPPORTED_PROBE_KINDS:
        raise ValueError(f"unsupported probe kind: {probe_kind}")
    if external_link_prompt_mode not in SUPPORTED_EXTERNAL_LINK_PROMPT_MODES:
        raise ValueError(f"unsupported external-link prompt mode: {external_link_prompt_mode}")
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
                    external_link_prompt_mode=external_link_prompt_mode,
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
                external_link_prompt_mode,
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
        external_link_prompt_mode,
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
    external_link_prompt_mode: str,
    completed: bool,
) -> dict:
    report = {
        "completed": completed,
        "fixture_dir": str(fixture_dir),
        "output_dir": str(output_dir.resolve()),
        "probes": list(probes),
        "include_fixture_patterns": list(include_fixture_patterns),
        "probe_kind": probe_kind,
        "external_link_prompt_mode": external_link_prompt_mode,
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
    external_link_prompt_mode: str,
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
            external_link_prompt_mode=external_link_prompt_mode,
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
    external_link_prompt_mode: str,
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

    timeline_before = (
        _timeline_selection_range(probe_path)
        if probe == "timeline_selection_state"
        else None
    )
    slicer_filter_before = (
        _slicer_table_filter_state(probe_path)
        if probe == "slicer_selection_state"
        else None
    )
    slicer_cache_before = (
        _slicer_cache_item_state(probe_path)
        if probe == "slicer_selection_state"
        else None
    )
    control_state_before = (
        _control_property_state(probe_path)
        if probe == "embedded_control_openability"
        else None
    )
    try:
        if probe == "external_link_update_prompt":
            active_name, ui_actions = _open_excel_with_ui_interaction(
                probe_path,
                probe,
                timeout,
                external_link_prompt_mode=external_link_prompt_mode,
            )
        else:
            active_name, ui_actions = _open_excel_with_ui_interaction(
                probe_path,
                probe,
                timeout,
            )
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
    if probe == "timeline_selection_state" and timeline_before is not None:
        timeline_after = _timeline_selection_range(probe_path)
        if timeline_after == timeline_before:
            return InteractiveProbeResult(
                fixture=fixture_label,
                probe=probe,
                probe_kind=UI_INTERACTION_PROBE_KIND,
                mutation=mutation,
                app="excel",
                status="failed",
                output=str(probe_path),
                message=(
                    "Excel timeline UI click did not change persisted "
                    f"timeline selection: {timeline_before}"
                ),
                ui_actions=ui_actions,
            )
    if probe == "slicer_selection_state":
        slicer_filter_after = _slicer_table_filter_state(probe_path)
        slicer_cache_after = _slicer_cache_item_state(probe_path)
        table_filter_changed = (
            bool(slicer_filter_after) and slicer_filter_after != slicer_filter_before
        )
        cache_item_changed = bool(slicer_cache_after) and slicer_cache_after != slicer_cache_before
        if not table_filter_changed and not cache_item_changed:
            return InteractiveProbeResult(
                fixture=fixture_label,
                probe=probe,
                probe_kind=UI_INTERACTION_PROBE_KIND,
                mutation=mutation,
                app="excel",
                status="failed",
                output=str(probe_path),
                message=(
                    "Excel slicer UI click did not change persisted "
                    "table filter or slicer-cache item state: "
                    f"table={slicer_filter_before}, cache={slicer_cache_before}"
                ),
                ui_actions=ui_actions,
            )
    if probe == "embedded_control_openability":
        control_state_after = _control_property_state(probe_path)
        if not control_state_after:
            return InteractiveProbeResult(
                fixture=fixture_label,
                probe=probe,
                probe_kind=UI_INTERACTION_PROBE_KIND,
                mutation=mutation,
                app="excel",
                status="failed",
                output=str(probe_path),
                message=(
                    "Excel embedded/control UI click did not change persisted "
                    f"control-property state: {control_state_before}"
                ),
                ui_actions=ui_actions,
            )
        if control_state_after == control_state_before and not _stateless_button_control_state(
            control_state_after
        ):
            return InteractiveProbeResult(
                fixture=fixture_label,
                probe=probe,
                probe_kind=UI_INTERACTION_PROBE_KIND,
                mutation=mutation,
                app="excel",
                status="failed",
                output=str(probe_path),
                message=(
                    "Excel embedded/control UI click did not change persisted "
                    f"control-property state: {control_state_before}"
                ),
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


def _open_excel_with_ui_interaction(
    src: Path,
    probe: str,
    timeout: int,
    *,
    external_link_prompt_mode: str = EXTERNAL_LINK_PROMPT_MODE_FORCE,
) -> tuple[str, list[str]]:
    ask_to_update_links = None
    if (
        probe == "external_link_update_prompt"
        and external_link_prompt_mode == EXTERNAL_LINK_PROMPT_MODE_FORCE
    ):
        ask_to_update_links = _excel_ask_to_update_links()
        if ask_to_update_links is None:
            raise RuntimeError(
                "could not read Excel ask to update links setting before forcing "
                "external-link prompt"
            )
        _set_excel_ask_to_update_links(True)
    try:
        return _open_excel_with_ui_interaction_impl(src, probe, timeout)
    finally:
        if ask_to_update_links is not None:
            _set_excel_ask_to_update_links(ask_to_update_links)


def _open_excel_with_ui_interaction_impl(
    src: Path, probe: str, timeout: int
) -> tuple[str, list[str]]:
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
            if probe == "embedded_control_openability":
                ui_actions.extend(_click_embedded_control(src))
                ui_actions.extend(_save_active_workbook())
            if probe == "pivot_refresh_state":
                ui_actions.extend(_refresh_all_active_workbook())
            if probe in {"slicer_selection_state", "timeline_selection_state"}:
                ui_actions.extend(_select_first_probe_shape(src, probe))
            if probe == "slicer_selection_state":
                ui_actions.extend(_click_slicer_items(src))
                ui_actions.extend(_save_active_workbook())
            if probe == "timeline_selection_state":
                ui_actions.extend(_click_timeline_month(src))
                ui_actions.extend(_save_active_workbook())
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


def _excel_ask_to_update_links() -> bool | None:
    script = 'tell application "Microsoft Excel" to get ask to update links'
    try:
        proc = run_ooxml_app_smoke._run_osascript(script, timeout=3)
    except subprocess.TimeoutExpired:
        return None
    value = proc.stdout.strip().lower()
    if value == "true":
        return True
    if value == "false":
        return False
    return None


def _set_excel_ask_to_update_links(enabled: bool) -> None:
    value = "true" if enabled else "false"
    script = f'tell application "Microsoft Excel" to set ask to update links to {value}'
    try:
        run_ooxml_app_smoke._run_osascript(script, timeout=3)
    except subprocess.TimeoutExpired:
        return


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


def _save_active_workbook() -> list[str]:
    script = """
tell application "Microsoft Excel"
  try
    save active workbook
    return "saved active workbook"
  on error errText
    return "failed Excel command: save active workbook: " & errText
  end try
end tell
"""
    try:
        proc = run_ooxml_app_smoke._run_osascript(script, timeout=10)
    except subprocess.TimeoutExpired:
        return ["failed Excel command: save active workbook: timeout"]
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


def _click_slicer_items(src: Path) -> list[str]:
    """Click the first visible item in every authored table slicer."""
    names = _probe_shape_names(src, "slicer_selection_state")
    if not names:
        return ["failed Excel slicer item click: no slicer shape names in package"]
    quoted_names = ", ".join(f'"{_escape_applescript_text(name)}"' for name in names)
    script = f"""
tell application "Microsoft Excel"
  activate
  try
    set bounds of active window to {{0, 40, 1600, 1050}}
  end try
  try
    set zoom of active window to 100
  end try
  try
    set scroll row of active window to 1
    set scroll column of active window to 1
  end try
  try
    set expectedNames to {{{quoted_names}}}
    set out to ""
    repeat with i from 1 to (count of shapes of active sheet)
      set candidate to shape i of active sheet
      set candidateName to name of candidate as text
      if expectedNames contains candidateName then
        if out is not "" then set out to out & linefeed
        set out to out & candidateName & "||" & ¬
          (left position of candidate as text) & "||" & ¬
          (top of candidate as text) & "||" & ¬
          (width of candidate as text) & "||" & ¬
          (height of candidate as text)
      end if
    end repeat
    return out
  on error errText
    return "ERROR: " & errText
  end try
end tell
"""
    try:
        proc = run_ooxml_app_smoke._run_osascript(script, timeout=10)
    except subprocess.TimeoutExpired:
        return ["failed Excel slicer item click: layout timeout"]
    geometry = proc.stdout.strip()
    if not geometry or geometry.startswith("ERROR:"):
        return [f"failed Excel slicer item click: {geometry or 'no matching shape'}"]

    # The probe normalizes the Excel window, zoom, and scroll origin above.
    # These offsets target the first visible button in the MyExcelOnline slicer
    # fixture; the persisted table filter change is the actual pass condition.
    sheet_origin_x = 24.0
    sheet_origin_y = 259.0
    actions = ["clicked Excel slicer item"]
    for line in geometry.splitlines():
        parts = line.split("||")
        if len(parts) != 5:
            return [f"failed Excel slicer item click: invalid shape geometry {line!r}"]
        name = parts[0]
        try:
            left, top, width, _height = [float(part) for part in parts[1:]]
        except ValueError:
            return [f"failed Excel slicer item click: invalid shape geometry {line!r}"]
        click_x = sheet_origin_x + left + (width / 2.0)
        click_y = sheet_origin_y + top + 33.0
        try:
            _post_mouse_click(click_x, click_y)
        except Exception as exc:
            return [f"failed Excel slicer item click: {str(exc)[:250]}"]
        actions.append(f"clicked Excel slicer item: {name}")
        time.sleep(0.2)
    return actions


def _click_embedded_control(src: Path) -> list[str]:
    """Click a visible worksheet form control and rely on persisted state."""
    names = _control_shape_names(src)
    if not names:
        return ["failed Excel embedded/control click: no control shape names in package"]
    quoted_names = ", ".join(f'"{_escape_applescript_text(name)}"' for name in names)
    script = f"""
tell application "Microsoft Excel"
  activate
  try
    set bounds of active window to {{0, 40, 1200, 850}}
  end try
  try
    set zoom of active window to 175
  end try
  try
    set scroll row of active window to 1
    set scroll column of active window to 1
  end try
  try
    set expectedNames to {{{quoted_names}}}
    repeat with i from 1 to (count of shapes of active sheet)
      set candidate to shape i of active sheet
      set candidateName to name of candidate as text
      if expectedNames contains candidateName then
        return candidateName & "||" & ¬
          (left position of candidate as text) & "||" & ¬
          (top of candidate as text) & "||" & ¬
          (width of candidate as text) & "||" & ¬
          (height of candidate as text)
      end if
    end repeat
    return ""
  on error errText
    return "ERROR: " & errText
  end try
end tell
"""
    try:
        proc = run_ooxml_app_smoke._run_osascript(script, timeout=10)
    except subprocess.TimeoutExpired:
        return ["failed Excel embedded/control click: layout timeout"]
    geometry = proc.stdout.strip()
    if not geometry or geometry.startswith("ERROR:"):
        return [f"failed Excel embedded/control click: {geometry or 'no matching shape'}"]
    parts = geometry.split("||")
    if len(parts) != 5:
        return [f"failed Excel embedded/control click: invalid shape geometry {geometry!r}"]
    name = parts[0]
    try:
        left, top, width, height = [float(part) for part in parts[1:]]
    except ValueError:
        return [f"failed Excel embedded/control click: invalid shape geometry {geometry!r}"]

    zoom = 1.75
    sheet_origin_x = 39.0
    sheet_origin_y = 276.0
    click_x = sheet_origin_x + (left * zoom) + (width * zoom * 0.2)
    click_y = sheet_origin_y + (top * zoom) + min(height * zoom * 0.45, 49.0)
    try:
        _post_mouse_click(click_x, click_y)
    except Exception as exc:
        return [f"failed Excel embedded/control click: {str(exc)[:250]}"]
    time.sleep(0.5)
    return [
        "clicked Excel embedded/control object",
        f"clicked Excel embedded/control object: {name}",
    ]


def _click_timeline_month(src: Path) -> list[str]:
    """Click a visible month in the first authored timeline on local Mac Excel."""
    selection = _timeline_selection_range(src)
    if selection is None:
        return ["failed Excel timeline month click: no persisted selection in package"]
    _, end_date = selection
    target_month_index = _timeline_target_month_index(end_date)
    names = _probe_shape_names(src, "timeline_selection_state")
    if not names:
        return ["failed Excel timeline month click: no timeline shape names in package"]
    quoted_names = ", ".join(f'"{_escape_applescript_text(name)}"' for name in names)
    script = f"""
tell application "Microsoft Excel"
  activate
  try
    set bounds of active window to {{0, 40, 1600, 1050}}
  end try
  try
    set zoom of active window to 80
  end try
  try
    set scroll row of active window to 1
    set scroll column of active window to 1
  end try
  try
    set expectedNames to {{{quoted_names}}}
    repeat with i from 1 to (count of shapes of active sheet)
      set candidate to shape i of active sheet
      set candidateName to name of candidate as text
      if expectedNames contains candidateName then
        return (left position of candidate as text) & "," & ¬
          (top of candidate as text) & "," & ¬
          (width of candidate as text) & "," & ¬
          (height of candidate as text)
      end if
    end repeat
    return ""
  on error errText
    return "ERROR: " & errText
  end try
end tell
"""
    try:
        proc = run_ooxml_app_smoke._run_osascript(script, timeout=10)
    except subprocess.TimeoutExpired:
        return ["failed Excel timeline month click: layout timeout"]
    geometry = proc.stdout.strip()
    if not geometry or geometry.startswith("ERROR:"):
        return [f"failed Excel timeline month click: {geometry or 'no matching shape'}"]
    try:
        left, top, width, _height = [float(part) for part in geometry.split(",")]
    except ValueError:
        return [f"failed Excel timeline month click: invalid shape geometry {geometry!r}"]

    # The probe normalizes Excel to a known top-left window, 80% worksheet zoom,
    # and scroll origin. These offsets target the month band inside the authored
    # MyExcelOnline timeline fixture and are verified by the persisted XML change.
    zoom = 0.8
    sheet_origin_x = 24.0
    sheet_origin_y = 253.0
    month_width = (width * zoom) / 12.0
    click_x = sheet_origin_x + (left * zoom) + ((target_month_index + 0.5) * month_width)
    click_y = sheet_origin_y + (top * zoom) + 62.0
    try:
        _post_mouse_click(click_x, click_y)
    except Exception as exc:
        return [f"failed Excel timeline month click: {str(exc)[:250]}"]
    target_label = _month_label(target_month_index)
    return ["clicked Excel timeline month", f"clicked Excel timeline month: {target_label}"]


def _timeline_target_month_index(end_date: str) -> int:
    parsed = dt.datetime.fromisoformat(end_date.replace("Z", "+00:00"))
    # Skip the immediately adjacent month to avoid timeline resize handles.
    return max(0, min(11, parsed.month + 1))


def _month_label(month_index: int) -> str:
    return dt.date(2000, month_index + 1, 1).strftime("%b")


def _post_mouse_click(x: float, y: float) -> None:
    try:
        import Quartz  # type: ignore[import-not-found]
    except ImportError:
        _post_mouse_click_with_external_python(x, y)
        return

    for event_type in (Quartz.kCGEventLeftMouseDown, Quartz.kCGEventLeftMouseUp):
        event = Quartz.CGEventCreateMouseEvent(
            None,
            event_type,
            (x, y),
            Quartz.kCGMouseButtonLeft,
        )
        Quartz.CGEventPost(Quartz.kCGHIDEventTap, event)
        time.sleep(0.08)


def _post_mouse_click_with_external_python(x: float, y: float) -> None:
    script = f"""
import time
import Quartz
for event_type in (Quartz.kCGEventLeftMouseDown, Quartz.kCGEventLeftMouseUp):
    event = Quartz.CGEventCreateMouseEvent(
        None,
        event_type,
        ({x!r}, {y!r}),
        Quartz.kCGMouseButtonLeft,
    )
    Quartz.CGEventPost(Quartz.kCGHIDEventTap, event)
    time.sleep(0.08)
"""
    candidates = [
        Path("/Library/Frameworks/Python.framework/Versions/3.13/bin/python3"),
        Path("/opt/homebrew/bin/python3"),
        Path("/usr/local/bin/python3"),
        Path("/usr/bin/python3"),
    ]
    for entry in os.environ.get("PATH", "").split(os.pathsep):
        if not entry:
            continue
        candidate = Path(entry) / "python3"
        if candidate not in candidates:
            candidates.append(candidate)
    errors: list[str] = []
    for candidate in candidates:
        if not candidate.exists() or str(candidate) == sys.executable:
            continue
        proc = subprocess.run(
            [str(candidate), "-c", script],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=5,
            check=False,
        )
        if proc.returncode == 0:
            return
        errors.append(f"{candidate}: {proc.stderr.strip()[:120]}")
    raise RuntimeError(
        "PyObjC Quartz is not available for mouse events"
        + (f" ({'; '.join(errors[:3])})" if errors else "")
    )


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


def _control_shape_names(path: Path) -> list[str]:
    names: list[str] = []
    try:
        with zipfile.ZipFile(path) as archive:
            for part_name in archive.namelist():
                if not part_name.startswith("xl/worksheets/") or not part_name.endswith(".xml"):
                    continue
                root = ET.fromstring(archive.read(part_name))
                for element in root.iter():
                    if _local_name(element.tag) != "control":
                        continue
                    name = element.attrib.get("name")
                    if name and name not in names:
                        names.append(name)
    except (zipfile.BadZipFile, ET.ParseError, OSError):
        return []
    return names


def _timeline_selection_range(path: Path) -> tuple[str, str] | None:
    try:
        with zipfile.ZipFile(path) as archive:
            timeline_cache_parts = [
                name
                for name in archive.namelist()
                if name.startswith("xl/timelineCaches/") and name.endswith(".xml")
            ]
            for part_name in sorted(timeline_cache_parts):
                root = ET.fromstring(archive.read(part_name))
                for element in root.iter():
                    if _local_name(element.tag) != "selection":
                        continue
                    start = element.attrib.get("startDate")
                    end = element.attrib.get("endDate")
                    if start and end:
                        return (start, end)
    except (zipfile.BadZipFile, ET.ParseError, OSError):
        return None
    return None


def _slicer_table_filter_state(path: Path) -> tuple[tuple[str, str, tuple[str, ...]], ...]:
    state: list[tuple[str, str, tuple[str, ...]]] = []
    try:
        with zipfile.ZipFile(path) as archive:
            table_parts = [
                name
                for name in archive.namelist()
                if name.startswith("xl/tables/") and name.endswith(".xml")
            ]
            for part_name in sorted(table_parts):
                root = ET.fromstring(archive.read(part_name))
                for filter_column in root.iter():
                    if _local_name(filter_column.tag) != "filterColumn":
                        continue
                    col_id = filter_column.attrib.get("colId", "")
                    values = tuple(
                        filter_element.attrib.get("val", "")
                        for filter_element in filter_column.iter()
                        if _local_name(filter_element.tag) == "filter"
                        and filter_element.attrib.get("val")
                    )
                    if values:
                        state.append((part_name, col_id, values))
    except (zipfile.BadZipFile, ET.ParseError, OSError):
        return ()
    return tuple(state)


def _slicer_cache_item_state(
    path: Path,
) -> tuple[tuple[str, tuple[tuple[int, tuple[tuple[str, str], ...]], ...]], ...]:
    state: list[tuple[str, tuple[tuple[int, tuple[tuple[str, str], ...]], ...]]] = []
    try:
        with zipfile.ZipFile(path) as archive:
            cache_parts = [
                name
                for name in archive.namelist()
                if name.startswith("xl/slicerCaches/") and name.endswith(".xml")
            ]
            for part_name in sorted(cache_parts):
                items: list[tuple[int, tuple[tuple[str, str], ...]]] = []
                root = ET.fromstring(archive.read(part_name))
                for index, element in enumerate(root.iter()):
                    if _local_name(element.tag) not in {"i", "item", "slicerCacheItem"}:
                        continue
                    if element.attrib:
                        items.append((index, tuple(sorted(element.attrib.items()))))
                if items:
                    state.append((part_name, tuple(items)))
    except (zipfile.BadZipFile, ET.ParseError, OSError):
        return ()
    return tuple(state)


def _control_property_state(path: Path) -> tuple[tuple[str, tuple[tuple[str, str], ...]], ...]:
    state: list[tuple[str, tuple[tuple[str, str], ...]]] = []
    try:
        with zipfile.ZipFile(path) as archive:
            control_parts = [
                name
                for name in archive.namelist()
                if name.startswith("xl/ctrlProps/") and name.endswith(".xml")
            ]
            for part_name in sorted(control_parts):
                root = ET.fromstring(archive.read(part_name))
                state.append((part_name, tuple(sorted(root.attrib.items()))))
    except (zipfile.BadZipFile, ET.ParseError, OSError):
        return ()
    return tuple(state)


def _stateless_button_control_state(
    state: tuple[tuple[str, tuple[tuple[str, str], ...]], ...],
) -> bool:
    if not state:
        return False
    selection_attrs = {"checked", "sel", "val"}
    for _part_name, attrs in state:
        attr_dict = dict(attrs)
        if attr_dict.get("objectType") != "Button":
            return False
        if selection_attrs & set(attr_dict):
            return False
    return True


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
        "--external-link-prompt-mode",
        choices=SUPPORTED_EXTERNAL_LINK_PROMPT_MODES,
        default=EXTERNAL_LINK_PROMPT_MODE_FORCE,
        help=(
            "For external-link UI interaction probes, either force Excel's "
            "ask-to-update-links setting temporarily or use the current "
            "setting without changing it."
        ),
    )
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
        external_link_prompt_mode=args.external_link_prompt_mode,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if report["failure_count"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
