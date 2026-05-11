#!/usr/bin/env python3
"""Inventory real-world OOXML corpus diversity for fidelity gap discovery."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from types import SimpleNamespace
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


@dataclass(frozen=True)
class SkippedWorkbook:
    path: str
    source_label: str
    tool: str | None
    reason: str


def audit_corpus(
    paths: list[Path],
    recursive: bool = False,
    workbook_timeout_seconds: float | None = None,
    package_only: bool = False,
) -> dict:
    workbooks: list[CorpusWorkbook] = []
    skipped_workbooks: list[SkippedWorkbook] = []
    seen: set[Path] = set()
    for source in paths:
        for workbook_path, tool in _discover_workbooks(source, recursive=recursive):
            resolved = workbook_path.resolve()
            if resolved in seen or not workbook_path.is_file():
                continue
            seen.add(resolved)
            try:
                workbook_features = _audit_workbook_features(
                    workbook_path,
                    timeout_seconds=workbook_timeout_seconds,
                    package_only=package_only,
                )
            except _SKIPPABLE_WORKBOOK_ERRORS as exc:
                skipped_workbooks.append(
                    SkippedWorkbook(
                        path=str(workbook_path),
                        source_label=str(source),
                        tool=tool,
                        reason=_error_reason(exc),
                    )
                )
                continue
            feature_keys = workbook_features["feature_keys"]
            surfaces = workbook_features["surfaces"]
            application = workbook_features["application"]
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
        "audit_mode": "package_only" if package_only else "full_snapshot",
        "workbook_count": len(workbooks),
        "skipped_workbook_count": len(skipped_workbooks),
        "required_buckets": sorted(REQUIRED_BUCKETS),
        "bucket_fixtures": bucket_fixtures,
        "missing_buckets": missing,
        "ready": not missing and not skipped_workbooks,
        "workbooks": [asdict(workbook) for workbook in workbooks],
        "skipped_workbooks": [
            asdict(skipped_workbook) for skipped_workbook in skipped_workbooks
        ],
    }


class WorkbookAuditTimeoutError(TimeoutError):
    """Raised when a single workbook takes too long to snapshot."""


class WorkbookAuditError(RuntimeError):
    """Raised when child-process workbook auditing reports a skippable failure."""


_SKIPPABLE_WORKBOOK_ERRORS = (
    ElementTree.ParseError,
    OSError,
    RuntimeError,
    WorkbookAuditError,
    WorkbookAuditTimeoutError,
    ValueError,
    zipfile.BadZipFile,
)


def _error_reason(exc: BaseException) -> str:
    message = str(exc).strip()
    if message:
        return f"{type(exc).__name__}: {message}"
    return type(exc).__name__


def _audit_workbook_features(
    workbook_path: Path,
    *,
    timeout_seconds: float | None,
    package_only: bool,
) -> dict:
    if package_only:
        return _audit_workbook_features_package_only(workbook_path)
    if not timeout_seconds or timeout_seconds <= 0:
        return _audit_workbook_features_unbounded(workbook_path)

    command = [
        sys.executable,
        str(Path(__file__).resolve()),
        "--_audit-one-workbook-json",
        str(workbook_path),
    ]
    try:
        result = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired:
        raise WorkbookAuditTimeoutError(
            f"timed out after {timeout_seconds:g}s while snapshotting {workbook_path}"
        )
    try:
        payload = json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        stderr = result.stderr.strip()
        detail = f": {stderr}" if stderr else ""
        raise WorkbookAuditError(
            f"worker returned invalid JSON for {workbook_path}{detail}"
        ) from exc
    if result.returncode != 0:
        raise WorkbookAuditError(str(payload.get("reason") or result.stderr.strip()))
    return payload


def _audit_workbook_features_unbounded(workbook_path: Path) -> dict:
    snapshot = audit_ooxml_fidelity.snapshot(workbook_path)
    return {
        "feature_keys": audit_ooxml_fidelity_coverage._feature_keys_for_snapshot(
            snapshot
        ),
        "surfaces": audit_ooxml_fidelity_coverage._surfaces_for_snapshot(snapshot),
        "application": _application_name(workbook_path),
    }


def _audit_workbook_features_package_only(workbook_path: Path) -> dict:
    with zipfile.ZipFile(workbook_path) as archive:
        parts = set(archive.namelist())
    snapshot = SimpleNamespace(
        feature_parts=audit_ooxml_fidelity._classify_feature_parts(parts),
        semantic_fingerprints=_package_only_semantic_fingerprints(parts),
    )
    return {
        "feature_keys": audit_ooxml_fidelity_coverage._feature_keys_for_snapshot(
            snapshot
        ),
        "surfaces": audit_ooxml_fidelity_coverage._surfaces_for_snapshot(snapshot),
        "application": _application_name(workbook_path),
    }


def _package_only_semantic_fingerprints(parts: set[str]) -> dict[str, dict[str, object]]:
    workbook_global_entries = []
    if "xl/workbook.xml" in parts:
        workbook_global_entries.append(("workbook_views", True))
    workbook_globals = {
        "xl/workbook.xml": workbook_global_entries,
        "package_parts": sorted(parts),
    }
    return {"workbook_globals": workbook_globals}


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
    if "python" in feature_keys:
        buckets.add("python_in_excel")
    if "sheet_metadata" in feature_keys:
        buckets.add("sheet_metadata")
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
    parser.add_argument(
        "--_audit-one-workbook-json",
        dest="audit_one_workbook_json",
        type=Path,
        help=argparse.SUPPRESS,
    )
    parser.add_argument("paths", nargs="*", type=Path)
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Discover workbooks recursively for non-manifest directories.",
    )
    parser.add_argument(
        "--workbook-timeout-seconds",
        type=float,
        default=None,
        help=(
            "Skip a workbook when snapshotting it exceeds this many seconds. "
            "The default preserves the historical no-timeout behavior."
        ),
    )
    parser.add_argument(
        "--package-only",
        action="store_true",
        help=(
            "Use package part names and workbook application metadata only. "
            "This is useful for large public corpora where full semantic "
            "snapshotting is too expensive."
        ),
    )
    parser.add_argument("--json", action="store_true", help="Emit JSON")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when any required corpus bucket is missing.",
    )
    args = parser.parse_args(argv)

    if args.audit_one_workbook_json is not None:
        try:
            payload = _audit_workbook_features_unbounded(args.audit_one_workbook_json)
        except _SKIPPABLE_WORKBOOK_ERRORS as exc:
            print(json.dumps({"reason": _error_reason(exc)}))
            return 1
        print(json.dumps(payload, sort_keys=True))
        return 0
    if not args.paths:
        parser.error("the following arguments are required: paths")

    report = audit_corpus(
        args.paths,
        recursive=args.recursive,
        workbook_timeout_seconds=args.workbook_timeout_seconds,
        package_only=args.package_only,
    )
    if args.json:
        print(json.dumps(report, indent=2, sort_keys=True))
    else:
        print(f"Workbooks: {report['workbook_count']}")
        print(f"Skipped workbooks: {report['skipped_workbook_count']}")
        print(f"Missing buckets: {len(report['missing_buckets'])}")
        for bucket in report["missing_buckets"]:
            print(f"- {bucket}")
    return 1 if args.strict and not report["ready"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
