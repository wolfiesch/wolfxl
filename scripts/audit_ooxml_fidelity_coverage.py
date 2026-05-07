#!/usr/bin/env python3
"""Summarize OOXML fidelity evidence coverage for real-world gap discovery."""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_ooxml_fidelity  # noqa: E402
import run_ooxml_fidelity_mutations  # noqa: E402

PASSING_STATUSES = {"passed", "passed_with_expected_drift"}
REAL_EXCEL_TOOLS = {"excel", "microsoft-excel", "excel-365", "excel-2021"}

SURFACES = {
    "pivot_slicer_preservation": {
        "label": "Pivot/slicer preservation across modify saves",
        "feature_keys": ("pivot", "slicer", "timeline"),
        "semantic_keys": ("pivots", "slicers", "timelines"),
        "structural_mutations": (
            "delete_first_row",
            "delete_first_col",
            "copy_first_sheet",
            "rename_first_sheet",
        ),
    },
    "chart_style_color_preservation": {
        "label": "Chart/style/color dependency preservation",
        "feature_keys": ("chart", "chart_sheet", "chart_style"),
        "semantic_keys": ("charts", "chart_sheets", "chart_styles"),
        "structural_mutations": (
            "add_remove_chart",
            "copy_first_sheet",
            "rename_first_sheet",
        ),
    },
    "conditional_formatting_extension_preservation": {
        "label": "Conditional formatting extension preservation",
        "feature_keys": ("conditional_formatting",),
        "semantic_keys": ("conditional_formatting",),
        "structural_mutations": (
            "delete_first_row",
            "delete_first_col",
            "add_conditional_formatting",
            "move_formula_range",
        ),
    },
    "external_link_relationship_edges": {
        "label": "External-link and workbook relationship edge cases",
        "feature_keys": ("external_link",),
        "semantic_keys": ("external_links",),
        "structural_mutations": (
            "rename_first_sheet",
            "move_formula_range",
        ),
    },
}


@dataclass(frozen=True)
class FixtureCoverage:
    filename: str
    fixture_id: str | None
    tool: str | None
    source_class: str
    surfaces: list[str]
    passed_mutations: list[str]


def audit_coverage(
    fixture_dir: Path,
    reports: Iterable[Path] = (),
) -> dict:
    fixture_dir = fixture_dir.resolve()
    passed_mutations = _passed_mutations_by_fixture(reports)
    fixtures = []
    for entry in run_ooxml_fidelity_mutations.discover_fixtures(fixture_dir):
        path = fixture_dir / entry.filename
        if not path.is_file():
            continue
        surfaces = _surfaces_for_fixture(path)
        fixtures.append(
            FixtureCoverage(
                filename=entry.filename,
                fixture_id=entry.fixture_id,
                tool=entry.tool,
                source_class=_source_class(entry.tool),
                surfaces=surfaces,
                passed_mutations=sorted(passed_mutations.get(entry.filename, set())),
            )
        )

    fixture_dicts = [asdict(fixture) for fixture in fixtures]
    surface_results = {
        name: _surface_result(name, fixture_dicts)
        for name in SURFACES
    }
    return {
        "fixture_dir": str(fixture_dir),
        "required_evidence": [
            "external_tool_fixture",
            "real_excel_fixture",
            "structural_mutation_pass",
        ],
        "fixture_count": len(fixtures),
        "fixtures": fixture_dicts,
        "surfaces": surface_results,
        "ready": all(not surface["missing"] for surface in surface_results.values()),
    }


def _passed_mutations_by_fixture(reports: Iterable[Path]) -> dict[str, set[str]]:
    out: dict[str, set[str]] = {}
    for report_path in reports:
        payload = json.loads(Path(report_path).read_text())
        for result in payload.get("results", []):
            if result.get("status") not in PASSING_STATUSES:
                continue
            fixture = result.get("fixture")
            mutation = result.get("mutation")
            if fixture and mutation:
                out.setdefault(str(fixture), set()).add(str(mutation))
    return out


def _surfaces_for_fixture(path: Path) -> list[str]:
    snapshot = audit_ooxml_fidelity.snapshot(path)
    out: list[str] = []
    for surface, config in SURFACES.items():
        has_feature_part = any(
            snapshot.feature_parts.get(key)
            for key in config["feature_keys"]
        )
        has_semantic_fingerprint = any(
            snapshot.semantic_fingerprints.get(key)
            for key in config["semantic_keys"]
        )
        if has_feature_part or has_semantic_fingerprint:
            out.append(surface)
    return out


def _surface_result(surface: str, fixtures: list[dict]) -> dict:
    config = SURFACES[surface]
    matching = [fixture for fixture in fixtures if surface in fixture["surfaces"]]
    external = [
        fixture["filename"]
        for fixture in matching
        if fixture["source_class"] == "external_tool"
    ]
    real_excel = [
        fixture["filename"]
        for fixture in matching
        if fixture["source_class"] == "real_excel"
    ]
    structural = [
        fixture["filename"]
        for fixture in matching
        if any(
            mutation in fixture["passed_mutations"]
            for mutation in config["structural_mutations"]
        )
    ]
    missing = []
    if not external:
        missing.append("external_tool_fixture")
    if not real_excel:
        missing.append("real_excel_fixture")
    if not structural:
        missing.append("structural_mutation_pass")
    return {
        "label": config["label"],
        "fixture_count": len(matching),
        "fixtures": [fixture["filename"] for fixture in matching],
        "external_tool_fixtures": external,
        "real_excel_fixtures": real_excel,
        "structural_mutation_fixtures": structural,
        "accepted_structural_mutations": list(config["structural_mutations"]),
        "missing": missing,
        "clear": not missing,
    }


def _source_class(tool: str | None) -> str:
    if not tool:
        return "unknown"
    normalized = tool.strip().lower()
    if normalized in REAL_EXCEL_TOOLS:
        return "real_excel"
    return "external_tool"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("fixture_dir", type=Path)
    parser.add_argument(
        "--report",
        action="append",
        type=Path,
        default=[],
        help="Mutation runner report.json. May be passed multiple times.",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when any P0 surface lacks required evidence.",
    )
    args = parser.parse_args(argv)

    report = audit_coverage(args.fixture_dir, args.report)
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
