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
PASSING_APP_STATUSES = {"passed"}
PASSING_NO_OP_RENDER_STATUSES = {"passed", "sampled_passed"}
PASSING_INTENTIONAL_RENDER_STATUSES = {"rendered", "sampled_rendered"}
REAL_EXCEL_TOOLS = {"excel", "microsoft-excel", "excel-365", "excel-2021"}

SURFACES = {
    "pivot_slicer_preservation": {
        "label": "Pivot/slicer preservation across modify saves",
        "feature_keys": ("pivot", "slicer", "timeline"),
        "semantic_keys": ("pivots", "slicers", "timelines"),
        "required_feature_groups": {
            "pivot": ("pivot",),
            "slicer_or_timeline": ("slicer", "timeline"),
        },
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
    "style_theme_color_preservation": {
        "label": "Workbook style/theme/color dependency preservation",
        "feature_keys": (),
        "semantic_keys": ("style_theme",),
        "structural_mutations": (
            "marker_cell",
            "copy_first_sheet",
            "rename_first_sheet",
            "move_formula_range",
        ),
    },
    "conditional_formatting_extension_preservation": {
        "label": "Conditional formatting extension preservation",
        "feature_keys": (),
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
    "workbook_connections_query_metadata": {
        "label": "Workbook connections / query metadata",
        "feature_keys": ("connection",),
        "semantic_keys": ("connections",),
        "structural_mutations": (
            "marker_cell",
            "copy_first_sheet",
            "move_formula_range",
        ),
    },
    "powerpivot_data_model_preservation": {
        "label": "PowerPivot / workbook data model preservation",
        "feature_keys": ("data_model",),
        "semantic_keys": ("data_model",),
        "required_source_classes": ("real_excel",),
        "structural_mutations": (
            "marker_cell",
            "copy_first_sheet",
            "move_formula_range",
        ),
    },
    "ooxml_extension_payload_preservation": {
        "label": "OOXML extension payload preservation",
        "feature_keys": (),
        "semantic_keys": ("extensions",),
        "structural_mutations": (
            "marker_cell",
            "copy_first_sheet",
            "move_formula_range",
        ),
    },
    "table_structured_refs_validations": {
        "label": "Tables / structured refs / validations",
        "feature_keys": ("table",),
        "semantic_keys": ("data_validations", "structured_references"),
        "required_feature_groups": {
            "table": ("table",),
            "data_validation": ("data_validation",),
            "structured_reference": ("structured_reference",),
        },
        "structural_mutations": (
            "delete_first_row",
            "delete_first_col",
            "copy_first_sheet",
            "rename_first_sheet",
            "move_formula_range",
        ),
    },
    "drawings_comments_embedded_objects": {
        "label": "Drawings / comments / embedded objects",
        "feature_keys": ("drawing", "comment", "image_media", "embedded_object"),
        "semantic_keys": (),
        "required_feature_groups": {
            "drawing": ("drawing",),
            "comment": ("comment",),
            "image_or_media": ("image_media",),
            "embedded_object": ("embedded_object",),
        },
        "structural_mutations": (
            "copy_first_sheet",
            "rename_first_sheet",
            "delete_first_row",
            "delete_first_col",
        ),
    },
    "workbook_global_state": {
        "label": "Workbook global state",
        "feature_keys": (
            "calc_chain",
            "custom_xml",
            "page_setup",
            "printer_settings",
            "vba",
        ),
        "semantic_keys": ("page_setup", "workbook_globals"),
        "required_feature_groups": {
            "defined_names_or_calc_chain": ("defined_name", "calc_chain"),
            "workbook_protection": ("workbook_protection",),
            "vba_or_custom_xml": ("vba", "custom_xml"),
            "printer_or_page_setup": ("printer_settings", "page_setup"),
        },
        "structural_mutations": (
            "delete_first_row",
            "delete_first_col",
            "copy_first_sheet",
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
    feature_keys: list[str]
    surfaces: list[str]
    passed_mutations: list[str]
    render_passes: list[str]
    intentional_render_passes: list[str]
    app_passes: list[str]
    intentional_app_passes: list[str]


def audit_coverage(
    fixture_dir: Path,
    reports: Iterable[Path] = (),
    render_reports: Iterable[Path] = (),
    app_reports: Iterable[Path] = (),
    require_render: bool = False,
    require_intentional_render: bool = False,
    require_app: bool = False,
    require_intentional_app: bool = False,
) -> dict:
    fixture_dir = fixture_dir.resolve()
    report_paths = list(reports)
    render_report_paths = list(render_reports)
    app_report_paths = list(app_reports)
    passed_mutations = _passed_mutations_by_fixture(report_paths)
    render_passes, intentional_render_passes = _render_passes_by_fixture(
        render_report_paths
    )
    app_passes, intentional_app_passes = _app_passes_by_fixture(app_report_paths)
    fixtures = []
    for entry in run_ooxml_fidelity_mutations.discover_fixtures(fixture_dir):
        path = fixture_dir / entry.filename
        if not path.is_file():
            continue
        snapshot = audit_ooxml_fidelity.snapshot(path)
        feature_keys = _feature_keys_for_snapshot(snapshot)
        surfaces = _surfaces_for_snapshot(snapshot)
        fixtures.append(
            FixtureCoverage(
                filename=entry.filename,
                fixture_id=entry.fixture_id,
                tool=entry.tool,
                source_class=_source_class(entry.tool),
                feature_keys=feature_keys,
                surfaces=surfaces,
                passed_mutations=sorted(passed_mutations.get(entry.filename, set())),
                render_passes=sorted(render_passes.get(entry.filename, set())),
                intentional_render_passes=sorted(
                    intentional_render_passes.get(entry.filename, set())
                ),
                app_passes=sorted(app_passes.get(entry.filename, set())),
                intentional_app_passes=sorted(
                    intentional_app_passes.get(entry.filename, set())
                ),
            )
        )

    fixture_dicts = [asdict(fixture) for fixture in fixtures]
    surface_results = {
        name: _surface_result(
            name,
            fixture_dicts,
            require_render=require_render,
            require_intentional_render=require_intentional_render,
            require_app=require_app,
            require_intentional_app=require_intentional_app,
        )
        for name in SURFACES
    }
    required_evidence = [
        "external_tool_fixture",
        "real_excel_fixture",
        "structural_mutation_pass",
    ]
    if require_render:
        required_evidence.append("render_no_op_pass")
    if require_intentional_render:
        required_evidence.append("intentional_render_pass")
    if require_app:
        required_evidence.append("app_open_pass")
    if require_intentional_app:
        required_evidence.append("intentional_app_open_pass")
    return {
        "fixture_dir": str(fixture_dir),
        "required_evidence": required_evidence,
        "mutation_report_count": len(report_paths),
        "render_report_count": len(render_report_paths),
        "app_report_count": len(app_report_paths),
        "render_required": require_render,
        "intentional_render_required": require_intentional_render,
        "app_required": require_app,
        "intentional_app_required": require_intentional_app,
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


def _render_passes_by_fixture(
    reports: Iterable[Path],
) -> tuple[dict[str, set[str]], dict[str, set[str]]]:
    no_op: dict[str, set[str]] = {}
    intentional: dict[str, set[str]] = {}
    for report_path in reports:
        payload = json.loads(Path(report_path).read_text())
        for result in payload.get("results", []):
            fixture = result.get("fixture")
            mutation = result.get("mutation", "no_op")
            status = result.get("status")
            if not fixture or not status:
                continue
            if mutation == "no_op" and status in PASSING_NO_OP_RENDER_STATUSES:
                no_op.setdefault(str(fixture), set()).add(str(status))
            elif (
                mutation != "no_op"
                and status in PASSING_INTENTIONAL_RENDER_STATUSES
            ):
                intentional.setdefault(str(fixture), set()).add(
                    f"{mutation}:{status}"
                )
    return no_op, intentional


def _app_passes_by_fixture(
    reports: Iterable[Path],
) -> tuple[dict[str, set[str]], dict[str, set[str]]]:
    source: dict[str, set[str]] = {}
    intentional: dict[str, set[str]] = {}
    for report_path in reports:
        payload = json.loads(Path(report_path).read_text())
        for result in payload.get("results", []):
            fixture = result.get("fixture")
            mutation = result.get("mutation", "source")
            app = result.get("app", "app")
            status = result.get("status")
            if not fixture or status not in PASSING_APP_STATUSES:
                continue
            label = f"{app}:{mutation}"
            if mutation == "source":
                source.setdefault(str(fixture), set()).add(label)
            else:
                intentional.setdefault(str(fixture), set()).add(label)
    return source, intentional


def _feature_keys_for_snapshot(snapshot: object) -> list[str]:
    keys = {
        key
        for key, values in snapshot.feature_parts.items()
        if values
    }
    semantic_to_feature = {
        "data_validations": "data_validation",
        "charts": "chart",
        "chart_sheets": "chart_sheet",
        "chart_styles": "chart_style",
        "conditional_formatting": "conditional_formatting",
        "connections": "connection",
        "data_model": "data_model",
        "external_links": "external_link",
        "extensions": "extension_payload",
        "page_setup": "page_setup",
        "pivots": "pivot",
        "slicers": "slicer",
        "style_theme": "style_theme",
        "structured_references": "structured_reference",
        "timelines": "timeline",
        "workbook_globals": "workbook_global",
        "worksheet_formulas": "worksheet_formula",
    }
    for semantic_key, feature_key in semantic_to_feature.items():
        fingerprint = snapshot.semantic_fingerprints.get(semantic_key)
        if fingerprint:
            keys.add(feature_key)
            if semantic_key == "workbook_globals":
                keys.update(_workbook_global_feature_keys(fingerprint))
    return sorted(keys)


def _surfaces_for_snapshot(snapshot: object) -> list[str]:
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


def _workbook_global_feature_keys(fingerprint: dict[str, object]) -> set[str]:
    keys: set[str] = set()
    workbook_entries = fingerprint.get("xl/workbook.xml")
    if isinstance(workbook_entries, list):
        for label, value in workbook_entries:
            if label == "defined_names" and value:
                keys.add("defined_name")
            elif label == "workbook_protection" and value:
                keys.add("workbook_protection")
            elif label in {"calc_pr", "workbook_views", "extensions"} and value:
                keys.add(f"workbook_{label}")
    package_parts = fingerprint.get("package_parts")
    if isinstance(package_parts, list):
        for part in package_parts:
            if part == "xl/calcChain.xml":
                keys.add("calc_chain")
            elif part == "xl/vbaProject.bin":
                keys.add("vba")
            elif str(part).startswith(("customXml/", "xl/customXml/")):
                keys.add("custom_xml")
            elif str(part).startswith("xl/printerSettings/"):
                keys.add("printer_settings")
    return keys


def _surface_result(
    surface: str,
    fixtures: list[dict],
    require_render: bool,
    require_intentional_render: bool,
    require_app: bool,
    require_intentional_app: bool,
) -> dict:
    config = SURFACES[surface]
    matching = [fixture for fixture in fixtures if surface in fixture["surfaces"]]
    required_source_classes = config.get(
        "required_source_classes", ("external_tool", "real_excel")
    )
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
    rendered = [
        fixture["filename"]
        for fixture in matching
        if fixture["render_passes"]
    ]
    intentional_rendered = [
        fixture["filename"]
        for fixture in matching
        if fixture["intentional_render_passes"]
    ]
    app_opened = [
        fixture["filename"]
        for fixture in matching
        if fixture["app_passes"]
    ]
    intentional_app_opened = [
        fixture["filename"]
        for fixture in matching
        if fixture["intentional_app_passes"]
    ]
    missing = []
    if "external_tool" in required_source_classes and not external:
        missing.append("external_tool_fixture")
    if "real_excel" in required_source_classes and not real_excel:
        missing.append("real_excel_fixture")
    if not structural:
        missing.append("structural_mutation_pass")
    if require_render and not rendered:
        missing.append("render_no_op_pass")
    if require_intentional_render and not intentional_rendered:
        missing.append("intentional_render_pass")
    if require_app and not app_opened:
        missing.append("app_open_pass")
    if require_intentional_app and not intentional_app_opened:
        missing.append("intentional_app_open_pass")
    feature_groups = config.get("required_feature_groups", {})
    group_results = {}
    for group, keys in feature_groups.items():
        group_matching = [
            fixture
            for fixture in matching
            if any(key in fixture["feature_keys"] for key in keys)
        ]
        group_external = [
            fixture["filename"]
            for fixture in group_matching
            if fixture["source_class"] == "external_tool"
        ]
        group_real_excel = [
            fixture["filename"]
            for fixture in group_matching
            if fixture["source_class"] == "real_excel"
        ]
        group_structural = [
            fixture["filename"]
            for fixture in group_matching
            if any(
                mutation in fixture["passed_mutations"]
                for mutation in config["structural_mutations"]
            )
        ]
        group_rendered = [
            fixture["filename"]
            for fixture in group_matching
            if fixture["render_passes"]
        ]
        group_intentional_rendered = [
            fixture["filename"]
            for fixture in group_matching
            if fixture["intentional_render_passes"]
        ]
        group_app_opened = [
            fixture["filename"]
            for fixture in group_matching
            if fixture["app_passes"]
        ]
        group_intentional_app_opened = [
            fixture["filename"]
            for fixture in group_matching
            if fixture["intentional_app_passes"]
        ]
        group_missing = []
        if not group_matching:
            group_missing.append("fixture")
        if "external_tool" in required_source_classes and not group_external:
            group_missing.append("external_tool_fixture")
        if "real_excel" in required_source_classes and not group_real_excel:
            group_missing.append("real_excel_fixture")
        if not group_structural:
            group_missing.append("structural_mutation_pass")
        if require_render and not group_rendered:
            group_missing.append("render_no_op_pass")
        if require_intentional_render and not group_intentional_rendered:
            group_missing.append("intentional_render_pass")
        if require_app and not group_app_opened:
            group_missing.append("app_open_pass")
        if require_intentional_app and not group_intentional_app_opened:
            group_missing.append("intentional_app_open_pass")
        if group_missing:
            missing.extend(f"{group}_{item}" for item in group_missing)
        group_results[group] = {
            "feature_keys": list(keys),
            "fixtures": [fixture["filename"] for fixture in group_matching],
            "external_tool_fixtures": group_external,
            "real_excel_fixtures": group_real_excel,
            "structural_mutation_fixtures": group_structural,
            "render_fixtures": group_rendered,
            "intentional_render_fixtures": group_intentional_rendered,
            "app_open_fixtures": group_app_opened,
            "intentional_app_open_fixtures": group_intentional_app_opened,
            "missing": group_missing,
            "clear": not group_missing,
        }
    return {
        "label": config["label"],
        "fixture_count": len(matching),
        "fixtures": [fixture["filename"] for fixture in matching],
        "external_tool_fixtures": external,
        "real_excel_fixtures": real_excel,
        "structural_mutation_fixtures": structural,
        "render_fixtures": rendered,
        "intentional_render_fixtures": intentional_rendered,
        "app_open_fixtures": app_opened,
        "intentional_app_open_fixtures": intentional_app_opened,
        "accepted_structural_mutations": list(config["structural_mutations"]),
        "required_source_classes": list(required_source_classes),
        "feature_groups": group_results,
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
        "--render-report",
        action="append",
        type=Path,
        default=[],
        help=(
            "Rendered comparison render-compare-report.json. May be passed "
            "multiple times."
        ),
    )
    parser.add_argument(
        "--app-report",
        action="append",
        type=Path,
        default=[],
        help=(
            "Spreadsheet app smoke app-smoke-report.json. May be passed "
            "multiple times."
        ),
    )
    parser.add_argument(
        "--require-render",
        action="store_true",
        help=(
            "Require at least one passing no-op render comparison for each "
            "fidelity surface."
        ),
    )
    parser.add_argument(
        "--require-intentional-render",
        action="store_true",
        help=(
            "Require at least one passing non-no-op mutation render smoke for "
            "each fidelity surface."
        ),
    )
    parser.add_argument(
        "--require-app",
        action="store_true",
        help=(
            "Require at least one passing source fixture app-open smoke for "
            "each fidelity surface."
        ),
    )
    parser.add_argument(
        "--require-intentional-app",
        action="store_true",
        help=(
            "Require at least one passing non-source mutation app-open smoke "
            "for each fidelity surface."
        ),
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when any fidelity surface lacks required evidence.",
    )
    args = parser.parse_args(argv)

    if args.strict and not args.report:
        print(
            "error: --strict requires at least one --report from "
            "run_ooxml_fidelity_mutations.py so structural mutation evidence "
            "can be evaluated.",
            file=sys.stderr,
        )
        return 2
    if args.require_render and not args.render_report:
        print(
            "error: --require-render requires at least one --render-report from "
            "run_ooxml_render_compare.py so rendered no-op evidence can be "
            "evaluated.",
            file=sys.stderr,
        )
        return 2
    if args.require_intentional_render and not args.render_report:
        print(
            "error: --require-intentional-render requires at least one "
            "--render-report from run_ooxml_render_compare.py so intentional "
            "mutation render evidence can be evaluated.",
            file=sys.stderr,
        )
        return 2
    if args.require_app and not args.app_report:
        print(
            "error: --require-app requires at least one --app-report from "
            "run_ooxml_app_smoke.py so app-open evidence can be evaluated.",
            file=sys.stderr,
        )
        return 2
    if args.require_intentional_app and not args.app_report:
        print(
            "error: --require-intentional-app requires at least one "
            "--app-report from run_ooxml_app_smoke.py so intentional mutation "
            "app-open evidence can be evaluated.",
            file=sys.stderr,
        )
        return 2

    report = audit_coverage(
        args.fixture_dir,
        reports=args.report,
        render_reports=args.render_report,
        app_reports=args.app_report,
        require_render=args.require_render,
        require_intentional_render=args.require_intentional_render,
        require_app=args.require_app,
        require_intentional_app=args.require_intentional_app,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
