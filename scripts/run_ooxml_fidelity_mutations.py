#!/usr/bin/env python3
"""Run OOXML fidelity mutation sweeps over workbook fixture directories."""

from __future__ import annotations

import argparse
import fnmatch
import hashlib
import json
import re
import shutil
import sys
from copy import deepcopy
from xml.etree import ElementTree
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_ooxml_fidelity  # noqa: E402
import wolfxl  # noqa: E402

CELL_REF_RE = re.compile(r"(?<![A-Za-z0-9_])\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})(?![A-Za-z0-9_])")
REF_SCAN_PART_PREFIXES = (
    "xl/workbook.xml",
    "xl/worksheets/",
    "xl/comments",
    "xl/drawings/",
    "xl/tables/",
    "xl/ctrlProps/",
    "xl/charts/",
    "xl/pivotTables/",
    "xl/pivotCache/",
    "xl/slicers/",
    "xl/timelines/",
)

DEFAULT_MUTATIONS = (
    "no_op",
    "marker_cell",
    "style_cell",
    "insert_tail_row",
    "insert_tail_col",
    "delete_marker_tail_row",
    "delete_marker_tail_col",
    "copy_remove_sheet",
    "move_marker_range",
)
SUPPORTED_MUTATIONS = (
    *DEFAULT_MUTATIONS,
    "delete_first_row",
    "delete_first_col",
    "copy_first_sheet",
    "rename_first_sheet",
    "add_remove_chart",
    "add_data_validation",
    "add_conditional_formatting",
    "move_formula_range",
    "retarget_external_links",
)
PASSING_STATUSES = {"passed", "passed_with_expected_drift"}
MARKER_CELL = "Z1"
MARKER_VALUE = "wolfxl_ooxml_fidelity_mutation"
STYLE_CELL = "AA1"
RENAMED_SHEET = "WolfXL Fidelity Rename"
SCRATCH_CHART_SHEET = "WolfXL Chart Scratch"
MANIFEST_NAME = "manifest.json"
RETARGETED_EXTERNAL_LINK = "wolfxl-retargeted-external-link.xlsx"
SPREADSHEET_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}
EXPECTED_DRAWING_ANCHOR_DRIFT_MUTATIONS = {
    "insert_tail_row",
    "insert_tail_col",
    "delete_marker_tail_row",
    "delete_marker_tail_col",
}
MISSING_RELATIONSHIP_RE = re.compile(
    r"^Relationship existed before save and is missing after save: "
    r"(?P<part>\S+) (?P<rid>\S+) (?P<rel_type>\S+) -> (?P<target>.*)$"
)
EXPECTED_ISSUE_KINDS_BY_MUTATION = {
    # Renaming a sheet is an intentional workbook semantic change. Feature
    # formulas and pivot cache worksheetSource sheet attrs may legitimately
    # change to keep pointing at the same sheet under its new title.
    "rename_first_sheet": {
        "charts_semantic_drift",
        "conditional_formatting_semantic_drift",
        "pivots_semantic_drift",
        "workbook_globals_semantic_drift",
        "worksheet_formulas_semantic_drift",
    },
    # Deleting the first row/column intentionally moves feature ranges. The
    # package-fidelity gate should still fail on part/relationship loss, while
    # declared semantic fingerprint shifts remain visible as expected issues.
    "delete_first_row": {
        "charts_semantic_drift",
        "conditional_formatting_semantic_drift",
        "data_validations_semantic_drift",
        "drawing_objects_semantic_drift",
        "external_links_semantic_drift",
        "extensions_semantic_drift",
        "feature_part_loss",
        "missing_part",
        "missing_relationship",
        "pivots_semantic_drift",
        "slicers_semantic_drift",
        "structured_references_semantic_drift",
        "workbook_globals_semantic_drift",
        "worksheet_formulas_semantic_drift",
    },
    "delete_first_col": {
        "charts_semantic_drift",
        "conditional_formatting_semantic_drift",
        "data_validations_semantic_drift",
        "drawing_objects_semantic_drift",
        "external_links_semantic_drift",
        "extensions_semantic_drift",
        "feature_part_loss",
        "missing_part",
        "missing_relationship",
        "pivots_semantic_drift",
        "slicers_semantic_drift",
        "structured_references_semantic_drift",
        "workbook_globals_semantic_drift",
        "worksheet_formulas_semantic_drift",
    },
    "copy_first_sheet": {
        "chart_styles_semantic_drift",
        "charts_semantic_drift",
        "conditional_formatting_semantic_drift",
        "data_validations_semantic_drift",
        "external_links_semantic_drift",
        "page_setup_semantic_drift",
        "pivots_semantic_drift",
        "slicers_semantic_drift",
        "structured_references_semantic_drift",
        "timelines_semantic_drift",
        "workbook_globals_semantic_drift",
        "worksheet_formulas_semantic_drift",
    },
    "add_data_validation": {
        "data_validations_semantic_drift",
    },
    "add_conditional_formatting": {
        "conditional_formatting_semantic_drift",
        "style_theme_semantic_drift",
    },
    "move_formula_range": {
        "worksheet_formulas_semantic_drift",
    },
    "retarget_external_links": {
        "external_links_semantic_drift",
        "missing_relationship",
    },
    "style_cell": {
        "style_theme_semantic_drift",
    },
}
EXPECTED_ISSUE_MARKERS_BY_MUTATION = {
    # Feature-add mutations should only accept semantic drift that contains
    # the newly authored range. A wholesale loss of the pre-existing feature
    # fingerprint must remain unexpected.
    "add_data_validation": {
        "data_validations_semantic_drift": "AB2:AB10",
    },
    "add_conditional_formatting": {
        "conditional_formatting_semantic_drift": "AC2:AC10",
    },
    "move_formula_range": {
        "worksheet_formulas_semantic_drift": "Z2",
    },
    "retarget_external_links": {
        "external_links_semantic_drift": "wolfxl-retargeted-external-link.xlsx",
    },
    "style_cell": {
        "style_theme_semantic_drift": "FF1F4E79",
    },
    "delete_first_row": {
        "external_links_semantic_drift": "worksheet_formulas",
        "feature_part_loss": "calc_chain",
        "missing_part": "xl/calcChain.xml",
        "missing_relationship": "relationships/calcChain",
    },
    "delete_first_col": {
        "external_links_semantic_drift": "worksheet_formulas",
        "feature_part_loss": "calc_chain",
        "missing_part": "xl/calcChain.xml",
        "missing_relationship": "relationships/calcChain",
    },
    "copy_first_sheet": {
        "external_links_semantic_drift": "worksheet_formulas",
    },
}
REQUIRED_EXPECTED_ISSUE_MARKERS_BY_MUTATION = {
    # This mutation is an oracle for formula translation. Passing requires the
    # audit to observe the formula move and the translated reference.
    "move_formula_range": {
        "worksheet_formulas_semantic_drift": "Z2",
    },
    "retarget_external_links": {
        "external_links_semantic_drift": "wolfxl-retargeted-external-link.xlsx",
    },
}


@dataclass
class FixtureEntry:
    filename: str
    sha256: str | None = None
    fixture_id: str | None = None
    tool: str | None = None
    app_unsupported_features: list[str] | None = None


@dataclass
class MutationResult:
    fixture: str
    mutation: str
    status: str
    issue_count: int
    issues: list[dict]
    expected_issue_count: int
    expected_issues: list[dict]
    before: str
    after: str
    error: str | None = None


def discover_fixtures(fixture_dir: Path, recursive: bool = False) -> list[FixtureEntry]:
    manifest = fixture_dir / MANIFEST_NAME
    if manifest.is_file():
        payload = json.loads(manifest.read_text())
        return [
            FixtureEntry(
                filename=entry["filename"],
                sha256=entry.get("sha256"),
                fixture_id=entry.get("fixture_id"),
                tool=entry.get("tool"),
                app_unsupported_features=entry.get("app_unsupported_features"),
            )
            for entry in payload.get("fixtures", [])
        ]

    pattern = "**/*" if recursive else "*"
    return [
        FixtureEntry(filename=path.relative_to(fixture_dir).as_posix())
        for path in sorted(fixture_dir.glob(pattern))
        if path.is_file()
        and path.suffix.lower() in SPREADSHEET_SUFFIXES
        and not path.name.startswith("~$")
    ]


def run_sweep(
    fixture_dir: Path,
    output_dir: Path,
    mutations: Iterable[str] = DEFAULT_MUTATIONS,
    verify_hashes: bool = True,
    recursive: bool = False,
    exclude_fixture_patterns: Iterable[str] = (),
) -> dict:
    fixture_dir = fixture_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    results: list[MutationResult] = []
    exclude_fixture_patterns = tuple(exclude_fixture_patterns)
    skipped_fixtures: list[str] = []

    for entry in discover_fixtures(fixture_dir, recursive=recursive):
        if _fixture_is_excluded(entry.filename, exclude_fixture_patterns):
            skipped_fixtures.append(entry.filename)
            continue
        fixture_path = fixture_dir / entry.filename
        if not fixture_path.is_file():
            results.append(
                MutationResult(
                    fixture=entry.filename,
                    mutation="discover",
                    status="missing_fixture",
                    issue_count=0,
                    issues=[],
                    expected_issue_count=0,
                    expected_issues=[],
                    before=str(fixture_path),
                    after="",
                    error=f"fixture missing: {fixture_path}",
                )
            )
            continue

        hash_error = _hash_error(fixture_path, entry.sha256, verify_hashes)
        for mutation in mutations:
            results.append(
                _run_single_mutation(
                    fixture_path=fixture_path,
                    fixture_label=entry.filename,
                    output_dir=output_dir,
                    mutation=mutation,
                    hash_error=hash_error,
                )
            )

    report = {
        "fixture_dir": str(fixture_dir),
        "output_dir": str(output_dir.resolve()),
        "mutations": list(mutations),
        "recursive": recursive,
        "exclude_fixture_patterns": list(exclude_fixture_patterns),
        "skipped_fixture_count": len(skipped_fixtures),
        "skipped_fixtures": skipped_fixtures,
        "result_count": len(results),
        "failure_count": sum(1 for result in results if result.status not in PASSING_STATUSES),
        "results": [asdict(result) for result in results],
    }
    (output_dir / "report.json").write_text(json.dumps(report, indent=2, sort_keys=True))
    return report


def _fixture_is_excluded(filename: str, patterns: Iterable[str]) -> bool:
    path = filename.replace("\\", "/")
    name = Path(path).name
    return any(
        fnmatch.fnmatch(path, pattern) or fnmatch.fnmatch(name, pattern) for pattern in patterns
    )


def _hash_error(path: Path, expected_hash: str | None, verify_hashes: bool) -> str | None:
    if not verify_hashes or not expected_hash:
        return None
    actual_hash = hashlib.sha256(path.read_bytes()).hexdigest()
    if actual_hash != expected_hash:
        return f"sha256 mismatch: expected {expected_hash}, got {actual_hash}"
    return None


def _run_single_mutation(
    fixture_path: Path,
    fixture_label: str,
    output_dir: Path,
    mutation: str,
    hash_error: str | None,
) -> MutationResult:
    mutation_dir = (
        output_dir / _safe_stem(Path(fixture_label).with_suffix("").as_posix()) / mutation
    )
    mutation_dir.mkdir(parents=True, exist_ok=True)
    before_path = mutation_dir / f"before-{fixture_path.name}"
    after_path = mutation_dir / f"after-{fixture_path.name}"
    shutil.copy2(fixture_path, before_path)
    shutil.copy2(fixture_path, after_path)

    if hash_error:
        return MutationResult(
            fixture=fixture_label,
            mutation=mutation,
            status="hash_mismatch",
            issue_count=0,
            issues=[],
            expected_issue_count=0,
            expected_issues=[],
            before=str(before_path),
            after=str(after_path),
            error=hash_error,
        )

    try:
        _prepare_mutation_baseline(before_path, after_path, mutation)
        _apply_mutation(after_path, mutation)
        audit_report = audit_ooxml_fidelity.audit(before_path, after_path)
        _assert_zip_integrity(after_path)
    except Exception as exc:
        return MutationResult(
            fixture=fixture_label,
            mutation=mutation,
            status="error",
            issue_count=0,
            issues=[],
            expected_issue_count=0,
            expected_issues=[],
            before=str(before_path),
            after=str(after_path),
            error=str(exc),
        )

    issues, expected_issues = _split_expected_issues(
        list(audit_report["issues"]),
        mutation,
        before_path=before_path,
        after_path=after_path,
    )
    missing_expected_issues = _missing_required_expected_issues(
        expected_issues, mutation, before_path=before_path
    )
    issues.extend(missing_expected_issues)
    status = "passed"
    if issues:
        status = "failed"
    elif expected_issues:
        status = "passed_with_expected_drift"
    return MutationResult(
        fixture=fixture_label,
        mutation=mutation,
        status=status,
        issue_count=len(issues),
        issues=issues,
        expected_issue_count=len(expected_issues),
        expected_issues=expected_issues,
        before=str(before_path),
        after=str(after_path),
    )


def _apply_mutation(path: Path, mutation: str) -> None:
    workbook = wolfxl.load_workbook(path, modify=True)
    try:
        if mutation == "no_op":
            pass
        elif mutation == "marker_cell":
            workbook[workbook.sheetnames[0]][MARKER_CELL] = MARKER_VALUE
        elif mutation == "style_cell":
            from wolfxl.styles import Font, PatternFill

            cell = workbook[workbook.sheetnames[0]][STYLE_CELL]
            cell.value = MARKER_VALUE
            cell.font = Font(bold=True, color="FF1F4E79")
            cell.fill = PatternFill(
                fill_type="solid",
                fgColor="FFEAF2F8",
            )
        elif mutation == "insert_tail_row":
            worksheet = workbook[workbook.sheetnames[0]]
            row_idx = _safe_tail_row_index(path, worksheet)
            worksheet.insert_rows(row_idx, amount=1)
            worksheet.cell(row=row_idx, column=1).value = MARKER_VALUE
        elif mutation == "insert_tail_col":
            worksheet = workbook[workbook.sheetnames[0]]
            col_idx = _safe_tail_col_index(path, worksheet)
            worksheet.insert_cols(col_idx, amount=1)
            worksheet.cell(row=1, column=col_idx).value = MARKER_VALUE
        elif mutation == "delete_marker_tail_row":
            worksheet = workbook[workbook.sheetnames[0]]
            row_idx = _safe_tail_row_index(path, worksheet)
            worksheet.cell(row=row_idx, column=1).value = MARKER_VALUE
            workbook.save(path)
            workbook.close()
            workbook = wolfxl.load_workbook(path, modify=True)
            worksheet = workbook[workbook.sheetnames[0]]
            worksheet.delete_rows(row_idx, amount=1)
        elif mutation == "delete_marker_tail_col":
            worksheet = workbook[workbook.sheetnames[0]]
            col_idx = _safe_tail_col_index(path, worksheet)
            worksheet.cell(row=1, column=col_idx).value = MARKER_VALUE
            workbook.save(path)
            workbook.close()
            workbook = wolfxl.load_workbook(path, modify=True)
            worksheet = workbook[workbook.sheetnames[0]]
            worksheet.delete_cols(col_idx, amount=1)
        elif mutation == "delete_first_row":
            workbook[workbook.sheetnames[0]].delete_rows(1, amount=1)
        elif mutation == "delete_first_col":
            workbook[workbook.sheetnames[0]].delete_cols(1, amount=1)
        elif mutation == "copy_first_sheet":
            workbook.copy_worksheet(workbook[workbook.sheetnames[0]])
        elif mutation == "copy_remove_sheet":
            clone = workbook.copy_worksheet(workbook[workbook.sheetnames[0]])
            clone_title = clone.title
            workbook.save(path)
            workbook.close()
            workbook = wolfxl.load_workbook(path, modify=True)
            workbook.remove(workbook[clone_title])
        elif mutation == "move_marker_range":
            worksheet = workbook[workbook.sheetnames[0]]
            worksheet["Z1"] = MARKER_VALUE
            worksheet["AA1"] = f"{MARKER_VALUE}_right"
            worksheet.move_range("Z1:AA1", rows=1, cols=0)
        elif mutation == "move_formula_range":
            worksheet = workbook[workbook.sheetnames[0]]
            worksheet.move_range("Z1:AA1", rows=1, cols=0, translate=True)
        elif mutation == "retarget_external_links":
            links = getattr(workbook, "_external_links", [])
            if links:
                links[0].update_target(RETARGETED_EXTERNAL_LINK)
        elif mutation == "rename_first_sheet":
            workbook[workbook.sheetnames[0]].title = RENAMED_SHEET
        elif mutation == "add_remove_chart":
            from wolfxl.chart import BarChart, Reference

            if SCRATCH_CHART_SHEET in workbook.sheetnames:
                workbook.remove(workbook[SCRATCH_CHART_SHEET])
            worksheet = workbook.create_sheet(SCRATCH_CHART_SHEET)
            worksheet["A1"] = "value"
            worksheet["A2"] = 1
            chart = BarChart()
            data = Reference(worksheet, min_col=1, min_row=1, max_row=2)
            chart.add_data(data, titles_from_data=True)
            worksheet.add_chart(chart, "C2")
            workbook.save(path)
            workbook.close()
            workbook = wolfxl.load_workbook(path, modify=True)
            worksheet = workbook[SCRATCH_CHART_SHEET]
            worksheet.remove_chart(worksheet._charts[-1])
            workbook.save(path)
            workbook.close()
            workbook = wolfxl.load_workbook(path, modify=True)
            worksheet = workbook[SCRATCH_CHART_SHEET]
            workbook.remove(worksheet)
        elif mutation == "add_data_validation":
            from wolfxl.worksheet.datavalidation import DataValidation

            worksheet = workbook[workbook.sheetnames[0]]
            worksheet.data_validations.append(
                DataValidation(
                    type="whole",
                    operator="between",
                    formula1="1",
                    formula2="100",
                    sqref="AB2:AB10",
                    showErrorMessage=True,
                )
            )
        elif mutation == "add_conditional_formatting":
            from wolfxl.formatting.rule import CellIsRule

            worksheet = workbook[workbook.sheetnames[0]]
            worksheet.conditional_formatting.add(
                "AC2:AC10",
                CellIsRule(
                    operator="greaterThan",
                    formula=["0"],
                    extra={"font_bold": True},
                ),
            )
        else:
            raise ValueError(f"unknown mutation: {mutation}")
        workbook.save(path)
    finally:
        close = getattr(workbook, "close", None)
        if close is not None:
            close()


def _safe_tail_row_index(path: Path, worksheet) -> int:
    max_row = int(getattr(worksheet, "max_row", 1) or 1)
    for _col_idx, row_idx in _package_cell_refs(path):
        max_row = max(max_row, row_idx)
    return min(max_row + 1, 1_048_576)


def _safe_tail_col_index(path: Path, worksheet) -> int:
    max_col = int(getattr(worksheet, "max_column", 1) or 1)
    for col_idx, _row_idx in _package_cell_refs(path):
        max_col = max(max_col, col_idx)
    return min(max_col + 1, 16_384)


def _package_cell_refs(path: Path) -> Iterable[tuple[int, int]]:
    try:
        with zipfile.ZipFile(path) as archive:
            for name in archive.namelist():
                if not name.endswith(".xml"):
                    continue
                if not name.startswith(REF_SCAN_PART_PREFIXES):
                    continue
                try:
                    text = archive.read(name).decode("utf-8", errors="ignore")
                except KeyError:
                    continue
                for col_letters, row_text in CELL_REF_RE.findall(text):
                    row_idx = int(row_text)
                    if row_idx > 1_048_576:
                        continue
                    col_idx = _column_index(col_letters)
                    if col_idx > 16_384:
                        continue
                    yield col_idx, row_idx
    except zipfile.BadZipFile:
        return


def _column_index(col_letters: str) -> int:
    idx = 0
    for char in col_letters:
        idx = idx * 26 + (ord(char) - ord("A") + 1)
    return idx


def _prepare_mutation_baseline(before_path: Path, after_path: Path, mutation: str) -> None:
    if mutation != "move_formula_range":
        return
    for path in (before_path, after_path):
        workbook = wolfxl.load_workbook(path, modify=True)
        try:
            worksheet = workbook[workbook.sheetnames[0]]
            worksheet["Z1"] = 10
            worksheet["AA1"] = "=Z1"
            workbook.save(path)
        finally:
            close = getattr(workbook, "close", None)
            if close is not None:
                close()


def _split_expected_issues(
    issues: list[dict],
    mutation: str,
    before_path: Path | None = None,
    after_path: Path | None = None,
) -> tuple[list[dict], list[dict]]:
    expected_kinds = EXPECTED_ISSUE_KINDS_BY_MUTATION.get(mutation, set())
    unexpected: list[dict] = []
    expected: list[dict] = []
    for issue in issues:
        if _is_expected_issue(
            issue,
            mutation,
            expected_kinds,
            before_path=before_path,
            after_path=after_path,
        ):
            expected.append(issue)
        else:
            unexpected.append(issue)
    return unexpected, expected


def _is_expected_issue(
    issue: dict,
    mutation: str,
    expected_kinds: set[str],
    before_path: Path | None = None,
    after_path: Path | None = None,
) -> bool:
    kind = issue.get("kind")
    if _is_expected_structural_drawing_anchor_drift(
        issue, mutation, before_path, after_path
    ):
        return True
    if kind not in expected_kinds:
        return False
    if _is_expected_deleted_first_axis_formula_loss(issue, mutation):
        return True
    if _looks_like_total_semantic_loss(issue):
        return False
    if mutation == "add_conditional_formatting" and kind == "style_theme_semantic_drift":
        return _is_expected_conditional_formatting_dxf_addition(before_path, after_path)
    if _is_expected_external_link_retarget_relationship_drift(issue, mutation, after_path):
        return True
    if mutation == "retarget_external_links" and kind == "missing_relationship":
        return False
    expected_markers = EXPECTED_ISSUE_MARKERS_BY_MUTATION.get(mutation, {})
    marker = expected_markers.get(kind)
    if marker is None:
        return True
    return marker in issue.get("message", "")


def _is_expected_structural_drawing_anchor_drift(
    issue: dict,
    mutation: str,
    before_path: Path | None,
    after_path: Path | None,
) -> bool:
    if mutation not in EXPECTED_DRAWING_ANCHOR_DRIFT_MUTATIONS:
        return False
    if issue.get("kind") != "drawing_objects_semantic_drift":
        return False
    if before_path is None or after_path is None:
        return False
    return _drawing_objects_match_except_structural_anchor_text(before_path, after_path)


def _drawing_objects_match_except_structural_anchor_text(
    before_path: Path, after_path: Path
) -> bool:
    with zipfile.ZipFile(before_path) as before_archive:
        before_fingerprint = audit_ooxml_fidelity._drawing_object_fingerprint(
            before_archive, set(before_archive.namelist())
        )
    with zipfile.ZipFile(after_path) as after_archive:
        after_fingerprint = audit_ooxml_fidelity._drawing_object_fingerprint(
            after_archive, set(after_archive.namelist())
        )
    return _normalize_structural_anchor_text(
        before_fingerprint
    ) == _normalize_structural_anchor_text(after_fingerprint)


def _normalize_structural_anchor_text(value, parent: str | None = None):
    if (
        isinstance(value, tuple)
        and len(value) == 4
        and value[0] == "Anchor"
        and isinstance(value[2], str)
    ):
        return (value[0], value[1], "<vml-anchor>", value[3])
    if (
        isinstance(value, tuple)
        and len(value) == 4
        and parent in {"from", "to"}
        and value[0] in {"row", "col"}
        and isinstance(value[2], str)
    ):
        return (
            value[0],
            value[1],
            "<drawingml-anchor>",
            _normalize_structural_anchor_text(value[3], value[0]),
        )
    if isinstance(value, dict):
        return {
            key: _normalize_structural_anchor_text(item, parent)
            for key, item in value.items()
        }
    if isinstance(value, list):
        return [_normalize_structural_anchor_text(item, parent) for item in value]
    if isinstance(value, tuple):
        next_parent = value[0] if len(value) == 4 and isinstance(value[0], str) else parent
        return tuple(_normalize_structural_anchor_text(item, next_parent) for item in value)
    return value


def _looks_like_total_semantic_loss(issue: dict) -> bool:
    message = issue.get("message", "")
    return "before={" in message and " after={}" in message


def _is_expected_deleted_first_axis_formula_loss(issue: dict, mutation: str) -> bool:
    if issue.get("kind") != "worksheet_formulas_semantic_drift":
        return False
    if mutation not in {"delete_first_row", "delete_first_col"}:
        return False
    message = issue.get("message", "")
    if " after={}" not in message:
        return False
    refs = re.findall(r"\('r', '([A-Z]+)([0-9]+)'\)", message)
    if not refs:
        return False
    if mutation == "delete_first_row":
        return all(row == "1" for _col, row in refs)
    return all(col == "A" for col, _row in refs)


def _is_expected_conditional_formatting_dxf_addition(
    before_path: Path | None, after_path: Path | None
) -> bool:
    if before_path is None or after_path is None:
        return False
    try:
        with (
            zipfile.ZipFile(before_path) as before_archive,
            zipfile.ZipFile(after_path) as after_archive,
        ):
            before_root = _read_styles_root(before_archive)
            after_root = _read_styles_root(after_archive)
    except (OSError, KeyError, ElementTree.ParseError, zipfile.BadZipFile):
        return False
    if before_root is None or after_root is None:
        return False
    before_dxf_count = _dxf_count(before_root)
    after_dxf_count = _dxf_count(after_root)
    if after_dxf_count != before_dxf_count + 1:
        return False
    return _strip_dxfs_fingerprint(before_root) == _strip_dxfs_fingerprint(after_root)


def _is_expected_external_link_retarget_relationship_drift(
    issue: dict,
    mutation: str,
    after_path: Path | None,
) -> bool:
    if mutation != "retarget_external_links":
        return False
    if issue.get("kind") != "missing_relationship":
        return False
    message = issue.get("message", "")
    if "relationships/externalLinkPath" not in message and "xlExternalLinkPath" not in message:
        return False
    return _external_link_retarget_replaces_missing_relationship(issue, after_path)


def _external_link_retarget_replaces_missing_relationship(issue: dict, path: Path | None) -> bool:
    if path is None:
        return False
    match = MISSING_RELATIONSHIP_RE.match(issue.get("message", ""))
    if not match:
        return False
    part = str(issue.get("part") or match.group("part"))
    if part != match.group("part"):
        return False
    if not part.startswith("xl/externalLinks/_rels/") or not part.endswith(".rels"):
        return False
    missing_rid = match.group("rid")
    try:
        with zipfile.ZipFile(path) as archive:
            rels_root = ElementTree.fromstring(archive.read(part))
    except (OSError, KeyError, ElementTree.ParseError, zipfile.BadZipFile):
        return False
    for relationship in rels_root:
        if relationship.attrib.get("Id") != missing_rid:
            continue
        return relationship.attrib.get("Target") == RETARGETED_EXTERNAL_LINK
    return False


def _read_styles_root(archive: zipfile.ZipFile) -> ElementTree.Element | None:
    try:
        return ElementTree.fromstring(archive.read("xl/styles.xml"))
    except KeyError:
        return None


def _dxf_count(root: ElementTree.Element) -> int:
    dxfs = _first_child_by_local_name(root, "dxfs")
    if dxfs is None:
        return 0
    return sum(1 for child in list(dxfs) if _local_name(child.tag) == "dxf")


def _strip_dxfs_fingerprint(root: ElementTree.Element) -> bytes:
    root_copy = deepcopy(root)
    for child in list(root_copy):
        if _local_name(child.tag) == "dxfs":
            root_copy.remove(child)
    return ElementTree.tostring(root_copy, encoding="utf-8")


def _first_child_by_local_name(
    root: ElementTree.Element, local_name: str
) -> ElementTree.Element | None:
    for child in list(root):
        if _local_name(child.tag) == local_name:
            return child
    return None


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _missing_required_expected_issues(
    expected_issues: list[dict],
    mutation: str,
    before_path: Path | None = None,
) -> list[dict]:
    required = _required_expected_issue_markers(mutation, before_path)
    missing: list[dict] = []
    for kind, marker in required.items():
        if any(
            issue.get("kind") == kind and marker in issue.get("message", "")
            for issue in expected_issues
        ):
            continue
        missing.append(
            {
                "severity": "error",
                "kind": "missing_required_expected_drift",
                "part": mutation,
                "message": (
                    f"{mutation} expected {kind} containing marker {marker!r}, "
                    "but the audit did not report it"
                ),
            }
        )
    return missing


def _required_expected_issue_markers(
    mutation: str,
    before_path: Path | None,
) -> dict[str, str]:
    required = REQUIRED_EXPECTED_ISSUE_MARKERS_BY_MUTATION.get(mutation, {})
    if mutation == "retarget_external_links" and not _has_external_link_parts(before_path):
        return {}
    return required


def _has_external_link_parts(path: Path | None) -> bool:
    if path is None:
        return False
    try:
        with zipfile.ZipFile(path) as archive:
            return any(name.startswith("xl/externalLinks/") for name in archive.namelist())
    except (OSError, zipfile.BadZipFile):
        return False


def _assert_zip_integrity(path: Path) -> None:
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
    if bad is not None:
        raise ValueError(f"ZIP integrity failure: {bad}")


def _safe_stem(stem: str) -> str:
    return "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in stem)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("fixture_dir", type=Path, help="Directory of OOXML spreadsheet fixtures")
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument(
        "--mutation",
        action="append",
        choices=SUPPORTED_MUTATIONS,
        dest="mutations",
        help="Mutation to run. May be passed multiple times.",
    )
    parser.add_argument("--no-verify-hashes", action="store_true")
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Discover OOXML spreadsheet fixtures recursively when no manifest.json is present.",
    )
    parser.add_argument(
        "--exclude-fixture",
        action="append",
        default=[],
        metavar="GLOB",
        help=(
            "Skip discovered fixtures whose relative path or basename matches GLOB. "
            "May be passed multiple times."
        ),
    )
    args = parser.parse_args(argv)

    report = run_sweep(
        fixture_dir=args.fixture_dir,
        output_dir=args.output_dir,
        mutations=args.mutations or DEFAULT_MUTATIONS,
        verify_hashes=not args.no_verify_hashes,
        recursive=args.recursive,
        exclude_fixture_patterns=args.exclude_fixture,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if report["failure_count"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
