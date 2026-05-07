#!/usr/bin/env python3
"""Audit interactive Excel evidence for high-risk OOXML surfaces."""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_ooxml_fidelity  # noqa: E402
import audit_ooxml_fidelity_coverage  # noqa: E402
import run_ooxml_fidelity_mutations  # noqa: E402

PASSING_STATUSES = {"passed"}
SPREADSHEET_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}

INTERACTIVE_PROBES = {
    "slicer_selection_state": {
        "label": "Slicer selection state survives open/save",
        "feature_keys": ("slicer",),
    },
    "timeline_selection_state": {
        "label": "Timeline selection state survives open/save",
        "feature_keys": ("timeline",),
    },
    "pivot_refresh_state": {
        "label": "Pivot refresh/filter state remains usable",
        "feature_keys": ("pivot",),
    },
    "external_link_update_prompt": {
        "label": "External-link update prompt and cached values remain sane",
        "feature_keys": ("external_link",),
    },
    "macro_project_presence": {
        "label": "Macro project remains present and openable",
        "feature_keys": ("vba",),
    },
    "embedded_control_openability": {
        "label": "Embedded controls/objects remain openable",
        "feature_keys": ("embedded_object",),
    },
}


@dataclass(frozen=True)
class InteractiveFixture:
    filename: str
    feature_keys: list[str]


def audit_interactive_evidence(
    fixture_dir: Path,
    reports: Iterable[Path] = (),
    recursive: bool = False,
) -> dict:
    fixture_dir = fixture_dir.resolve()
    report_paths = list(reports)
    fixtures = _discover_interactive_fixtures(fixture_dir, recursive=recursive)
    passed = _passed_probes_by_fixture(report_paths)
    fixture_dicts = [asdict(fixture) for fixture in fixtures]
    probe_results = {
        name: _probe_result(name, fixture_dicts, passed)
        for name in INTERACTIVE_PROBES
    }
    return {
        "fixture_dir": str(fixture_dir),
        "recursive": recursive,
        "report_count": len(report_paths),
        "fixture_count": len(fixtures),
        "fixtures": fixture_dicts,
        "probes": probe_results,
        "ready": all(result["clear"] for result in probe_results.values()),
    }


def _discover_interactive_fixtures(
    fixture_dir: Path, recursive: bool
) -> list[InteractiveFixture]:
    fixtures: list[InteractiveFixture] = []
    for entry in run_ooxml_fidelity_mutations.discover_fixtures(
        fixture_dir, recursive=recursive
    ):
        path = fixture_dir / entry.filename
        if not path.is_file() or path.suffix.lower() not in SPREADSHEET_SUFFIXES:
            continue
        try:
            snapshot = audit_ooxml_fidelity.snapshot(path)
        except zipfile.BadZipFile:
            continue
        feature_keys = audit_ooxml_fidelity_coverage._feature_keys_for_snapshot(
            snapshot
        )
        fixtures.append(
            InteractiveFixture(
                filename=entry.filename,
                feature_keys=feature_keys,
            )
        )
    return fixtures


def _passed_probes_by_fixture(reports: Iterable[Path]) -> dict[str, set[str]]:
    out: dict[str, set[str]] = {}
    for report_path in reports:
        payload = json.loads(Path(report_path).read_text())
        for result in payload.get("results", []):
            if result.get("status") not in PASSING_STATUSES:
                continue
            fixture = result.get("fixture")
            probe = result.get("probe")
            if fixture and probe:
                out.setdefault(str(fixture), set()).add(str(probe))
    return out


def _probe_result(
    probe: str,
    fixtures: list[dict],
    passed: dict[str, set[str]],
) -> dict:
    config = INTERACTIVE_PROBES[probe]
    candidates = [
        fixture["filename"]
        for fixture in fixtures
        if any(key in fixture["feature_keys"] for key in config["feature_keys"])
    ]
    passed_fixtures = [
        fixture for fixture in candidates if probe in passed.get(fixture, set())
    ]
    if not candidates:
        status = "not_applicable"
        missing: list[str] = []
    elif passed_fixtures:
        status = "clear"
        missing = []
    else:
        status = "missing"
        missing = ["interactive_probe_pass"]
    return {
        "label": config["label"],
        "feature_keys": list(config["feature_keys"]),
        "candidate_fixtures": candidates,
        "passed_fixtures": passed_fixtures,
        "missing": missing,
        "status": status,
        "clear": not missing,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("fixture_dir", type=Path)
    parser.add_argument(
        "--report",
        action="append",
        type=Path,
        default=[],
        help="Interactive probe report JSON. May be passed multiple times.",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Discover workbooks recursively for non-manifest fixture dirs.",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when any applicable interactive probe lacks evidence.",
    )
    args = parser.parse_args(argv)

    report = audit_interactive_evidence(
        args.fixture_dir,
        reports=args.report,
        recursive=args.recursive,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
