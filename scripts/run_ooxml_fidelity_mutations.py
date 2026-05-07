#!/usr/bin/env python3
"""Run OOXML fidelity mutation sweeps over workbook fixture directories."""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_ooxml_fidelity  # noqa: E402
import wolfxl  # noqa: E402

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
)
PASSING_STATUSES = {"passed", "passed_with_expected_drift"}
MARKER_CELL = "Z1"
MARKER_VALUE = "wolfxl_ooxml_fidelity_mutation"
STYLE_CELL = "AA1"
RENAMED_SHEET = "WolfXL Fidelity Rename"
SCRATCH_CHART_SHEET = "WolfXL Chart Scratch"
MANIFEST_NAME = "manifest.json"
EXPECTED_ISSUE_KINDS_BY_MUTATION = {
    # Renaming a sheet is an intentional workbook semantic change. Feature
    # formulas and pivot cache worksheetSource sheet attrs may legitimately
    # change to keep pointing at the same sheet under its new title.
    "rename_first_sheet": {
        "charts_semantic_drift",
        "conditional_formatting_semantic_drift",
        "pivots_semantic_drift",
        "worksheet_formulas_semantic_drift",
    },
    # Deleting the first row/column intentionally moves feature ranges. The
    # package-fidelity gate should still fail on part/relationship loss, while
    # declared semantic fingerprint shifts remain visible as expected issues.
    "delete_first_row": {
        "charts_semantic_drift",
        "conditional_formatting_semantic_drift",
        "data_validations_semantic_drift",
        "external_links_semantic_drift",
        "missing_part",
        "missing_relationship",
        "pivots_semantic_drift",
        "slicers_semantic_drift",
        "worksheet_formulas_semantic_drift",
    },
    "delete_first_col": {
        "charts_semantic_drift",
        "conditional_formatting_semantic_drift",
        "data_validations_semantic_drift",
        "external_links_semantic_drift",
        "missing_part",
        "missing_relationship",
        "pivots_semantic_drift",
        "slicers_semantic_drift",
        "worksheet_formulas_semantic_drift",
    },
    "copy_first_sheet": {
        "charts_semantic_drift",
        "conditional_formatting_semantic_drift",
        "data_validations_semantic_drift",
        "external_links_semantic_drift",
        "pivots_semantic_drift",
        "slicers_semantic_drift",
        "timelines_semantic_drift",
        "worksheet_formulas_semantic_drift",
    },
    "add_data_validation": {
        "data_validations_semantic_drift",
    },
    "add_conditional_formatting": {
        "conditional_formatting_semantic_drift",
    },
    "move_formula_range": {
        "worksheet_formulas_semantic_drift",
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
    "delete_first_row": {
        "external_links_semantic_drift": "worksheet_formulas",
        "missing_part": "xl/calcChain.xml",
        "missing_relationship": "relationships/calcChain",
    },
    "delete_first_col": {
        "external_links_semantic_drift": "worksheet_formulas",
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
}


@dataclass
class FixtureEntry:
    filename: str
    sha256: str | None = None
    fixture_id: str | None = None
    tool: str | None = None


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


def discover_fixtures(fixture_dir: Path) -> list[FixtureEntry]:
    manifest = fixture_dir / MANIFEST_NAME
    if manifest.is_file():
        payload = json.loads(manifest.read_text())
        return [
            FixtureEntry(
                filename=entry["filename"],
                sha256=entry.get("sha256"),
                fixture_id=entry.get("fixture_id"),
                tool=entry.get("tool"),
            )
            for entry in payload.get("fixtures", [])
        ]

    return [
        FixtureEntry(filename=path.name)
        for path in sorted(fixture_dir.glob("*.xlsx"))
        if path.is_file() and not path.name.startswith("~$")
    ]


def run_sweep(
    fixture_dir: Path,
    output_dir: Path,
    mutations: Iterable[str] = DEFAULT_MUTATIONS,
    verify_hashes: bool = True,
) -> dict:
    fixture_dir = fixture_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    results: list[MutationResult] = []

    for entry in discover_fixtures(fixture_dir):
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
                    output_dir=output_dir,
                    mutation=mutation,
                    hash_error=hash_error,
                )
            )

    report = {
        "fixture_dir": str(fixture_dir),
        "output_dir": str(output_dir.resolve()),
        "mutations": list(mutations),
        "result_count": len(results),
        "failure_count": sum(
            1 for result in results if result.status not in PASSING_STATUSES
        ),
        "results": [asdict(result) for result in results],
    }
    (output_dir / "report.json").write_text(json.dumps(report, indent=2, sort_keys=True))
    return report


def _hash_error(path: Path, expected_hash: str | None, verify_hashes: bool) -> str | None:
    if not verify_hashes or not expected_hash:
        return None
    actual_hash = hashlib.sha256(path.read_bytes()).hexdigest()
    if actual_hash != expected_hash:
        return f"sha256 mismatch: expected {expected_hash}, got {actual_hash}"
    return None


def _run_single_mutation(
    fixture_path: Path, output_dir: Path, mutation: str, hash_error: str | None
) -> MutationResult:
    mutation_dir = output_dir / _safe_stem(fixture_path.stem) / mutation
    mutation_dir.mkdir(parents=True, exist_ok=True)
    before_path = mutation_dir / f"before-{fixture_path.name}"
    after_path = mutation_dir / f"after-{fixture_path.name}"
    shutil.copy2(fixture_path, before_path)
    shutil.copy2(fixture_path, after_path)

    if hash_error:
        return MutationResult(
            fixture=fixture_path.name,
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
            fixture=fixture_path.name,
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
        list(audit_report["issues"]), mutation
    )
    missing_expected_issues = _missing_required_expected_issues(
        expected_issues, mutation
    )
    issues.extend(missing_expected_issues)
    status = "passed"
    if issues:
        status = "failed"
    elif expected_issues:
        status = "passed_with_expected_drift"
    return MutationResult(
        fixture=fixture_path.name,
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
            row_idx = int(getattr(worksheet, "max_row", 1) or 1) + 1
            worksheet.insert_rows(row_idx, amount=1)
            worksheet.cell(row=row_idx, column=1).value = MARKER_VALUE
        elif mutation == "insert_tail_col":
            worksheet = workbook[workbook.sheetnames[0]]
            col_idx = int(getattr(worksheet, "max_column", 1) or 1) + 1
            worksheet.insert_cols(col_idx, amount=1)
            worksheet.cell(row=1, column=col_idx).value = MARKER_VALUE
        elif mutation == "delete_marker_tail_row":
            worksheet = workbook[workbook.sheetnames[0]]
            row_idx = int(getattr(worksheet, "max_row", 1) or 1) + 1
            worksheet.cell(row=row_idx, column=1).value = MARKER_VALUE
            workbook.save(path)
            workbook.close()
            workbook = wolfxl.load_workbook(path, modify=True)
            worksheet = workbook[workbook.sheetnames[0]]
            worksheet.delete_rows(row_idx, amount=1)
        elif mutation == "delete_marker_tail_col":
            worksheet = workbook[workbook.sheetnames[0]]
            col_idx = int(getattr(worksheet, "max_column", 1) or 1) + 1
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
    issues: list[dict], mutation: str
) -> tuple[list[dict], list[dict]]:
    expected_kinds = EXPECTED_ISSUE_KINDS_BY_MUTATION.get(mutation, set())
    unexpected: list[dict] = []
    expected: list[dict] = []
    for issue in issues:
        if _is_expected_issue(issue, mutation, expected_kinds):
            expected.append(issue)
        else:
            unexpected.append(issue)
    return unexpected, expected


def _is_expected_issue(issue: dict, mutation: str, expected_kinds: set[str]) -> bool:
    kind = issue.get("kind")
    if kind not in expected_kinds:
        return False
    expected_markers = EXPECTED_ISSUE_MARKERS_BY_MUTATION.get(mutation, {})
    marker = expected_markers.get(kind)
    if marker is None:
        return True
    return marker in issue.get("message", "")


def _missing_required_expected_issues(
    expected_issues: list[dict], mutation: str
) -> list[dict]:
    required = REQUIRED_EXPECTED_ISSUE_MARKERS_BY_MUTATION.get(mutation, {})
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


def _assert_zip_integrity(path: Path) -> None:
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
    if bad is not None:
        raise ValueError(f"ZIP integrity failure: {bad}")


def _safe_stem(stem: str) -> str:
    return "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in stem)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("fixture_dir", type=Path, help="Directory of .xlsx fixtures")
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument(
        "--mutation",
        action="append",
        choices=SUPPORTED_MUTATIONS,
        dest="mutations",
        help="Mutation to run. May be passed multiple times.",
    )
    parser.add_argument("--no-verify-hashes", action="store_true")
    args = parser.parse_args(argv)

    report = run_sweep(
        fixture_dir=args.fixture_dir,
        output_dir=args.output_dir,
        mutations=args.mutations or DEFAULT_MUTATIONS,
        verify_hashes=not args.no_verify_hashes,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if report["failure_count"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
