#!/usr/bin/env python3
"""Inventory real-world OOXML corpus diversity for fidelity gap discovery."""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_ooxml_fidelity  # noqa: E402
import audit_ooxml_fidelity_coverage  # noqa: E402
import run_ooxml_fidelity_mutations  # noqa: E402

APP_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"
SPREADSHEET_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}

REQUIRED_BUCKETS = {
    "excel_authored",
    "external_tool_authored",
    "macro_vba",
    "powerpivot_data_model",
    "slicer_or_timeline",
    "embedded_object_or_control",
    "external_link",
    "chart_or_chart_style",
    "conditional_formatting_extension",
    "table_structured_ref_or_validation",
    "drawing_comment_or_media",
    "workbook_global_state",
}


@dataclass(frozen=True)
class CorpusWorkbook:
    path: str
    source_label: str
    tool: str | None
    application: str | None
    feature_keys: list[str]
    buckets: list[str]
    surfaces: list[str]


def audit_corpus(paths: list[Path], recursive: bool = False) -> dict:
    workbooks: list[CorpusWorkbook] = []
    seen: set[Path] = set()
    for source in paths:
        for workbook_path, tool in _discover_workbooks(source, recursive=recursive):
            resolved = workbook_path.resolve()
            if resolved in seen or not workbook_path.is_file():
                continue
            seen.add(resolved)
            try:
                snapshot = audit_ooxml_fidelity.snapshot(workbook_path)
            except zipfile.BadZipFile:
                continue
            feature_keys = audit_ooxml_fidelity_coverage._feature_keys_for_snapshot(
                snapshot
            )
            surfaces = audit_ooxml_fidelity_coverage._surfaces_for_snapshot(snapshot)
            application = _application_name(workbook_path)
            buckets = _buckets_for_workbook(
                tool=tool,
                application=application,
                feature_keys=set(feature_keys),
                surfaces=set(surfaces),
            )
            workbooks.append(
                CorpusWorkbook(
                    path=str(workbook_path),
                    source_label=str(source),
                    tool=tool,
                    application=application,
                    feature_keys=feature_keys,
                    buckets=sorted(buckets),
                    surfaces=surfaces,
                )
            )

    bucket_fixtures: dict[str, list[str]] = {bucket: [] for bucket in REQUIRED_BUCKETS}
    for workbook in workbooks:
        for bucket in workbook.buckets:
            bucket_fixtures.setdefault(bucket, []).append(workbook.path)
    bucket_fixtures = {
        bucket: sorted(paths) for bucket, paths in sorted(bucket_fixtures.items())
    }
    missing = sorted(bucket for bucket in REQUIRED_BUCKETS if not bucket_fixtures[bucket])
    return {
        "workbook_count": len(workbooks),
        "required_buckets": sorted(REQUIRED_BUCKETS),
        "bucket_fixtures": bucket_fixtures,
        "missing_buckets": missing,
        "ready": not missing,
        "workbooks": [asdict(workbook) for workbook in workbooks],
    }


def _discover_workbooks(source: Path, recursive: bool) -> list[tuple[Path, str | None]]:
    if source.is_file() and source.suffix.lower() in SPREADSHEET_SUFFIXES:
        return [(source, None)]
    if not source.is_dir():
        return []

    manifest = source / run_ooxml_fidelity_mutations.MANIFEST_NAME
    if manifest.is_file():
        return [
            (source / entry.filename, entry.tool)
            for entry in run_ooxml_fidelity_mutations.discover_fixtures(source)
        ]

    pattern = "**/*" if recursive else "*"
    return [
        (path, None)
        for path in sorted(source.glob(pattern))
        if path.is_file()
        and path.suffix.lower() in SPREADSHEET_SUFFIXES
        and not path.name.startswith("~$")
    ]


def _application_name(path: Path) -> str | None:
    try:
        with zipfile.ZipFile(path) as archive:
            root = ElementTree.fromstring(archive.read("docProps/app.xml"))
    except (KeyError, ElementTree.ParseError, zipfile.BadZipFile):
        return None
    node = root.find(f"{APP_NS}Application")
    if node is None or node.text is None:
        return None
    return node.text.strip() or None


def _buckets_for_workbook(
    *, tool: str | None, application: str | None, feature_keys: set[str], surfaces: set[str]
) -> set[str]:
    buckets: set[str] = set()
    if _is_excel_authored(tool, application):
        buckets.add("excel_authored")
    if _is_external_tool_authored(tool, application):
        buckets.add("external_tool_authored")

    if "vba" in feature_keys:
        buckets.add("macro_vba")
    if "data_model" in feature_keys:
        buckets.add("powerpivot_data_model")
    if {"slicer", "timeline"} & feature_keys:
        buckets.add("slicer_or_timeline")
    if {"embedded_object", "drawing_object"} & feature_keys:
        buckets.add("embedded_object_or_control")
    if "external_link" in feature_keys:
        buckets.add("external_link")
    if {"chart", "chart_sheet", "chart_style"} & feature_keys:
        buckets.add("chart_or_chart_style")
    if "conditional_formatting" in feature_keys:
        buckets.add("conditional_formatting_extension")
    if {"table", "structured_reference", "data_validation"} & feature_keys:
        buckets.add("table_structured_ref_or_validation")
    if {"drawing", "comment", "image_media", "drawing_object"} & feature_keys:
        buckets.add("drawing_comment_or_media")
    if "workbook_global_state" in surfaces or "workbook_global" in feature_keys:
        buckets.add("workbook_global_state")
    return buckets


def _is_excel_authored(tool: str | None, application: str | None) -> bool:
    values = [value.lower() for value in (tool, application) if value]
    return any("excel" in value and "excelize" not in value for value in values)


def _is_external_tool_authored(tool: str | None, application: str | None) -> bool:
    values = [value.lower() for value in (tool, application) if value]
    return any("excel" not in value or "excelize" in value for value in values)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("paths", nargs="+", type=Path)
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Discover workbooks recursively for non-manifest directories.",
    )
    parser.add_argument("--json", action="store_true", help="Emit JSON")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when any required corpus bucket is missing.",
    )
    args = parser.parse_args(argv)

    report = audit_corpus(args.paths, recursive=args.recursive)
    if args.json:
        print(json.dumps(report, indent=2, sort_keys=True))
    else:
        print(f"Workbooks: {report['workbook_count']}")
        print(f"Missing buckets: {len(report['missing_buckets'])}")
        for bucket in report["missing_buckets"]:
            print(f"- {bucket}")
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
