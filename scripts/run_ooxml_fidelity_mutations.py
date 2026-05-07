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
    "move_marker_range",
)
SUPPORTED_MUTATIONS = (*DEFAULT_MUTATIONS, "rename_first_sheet")
PASSING_STATUSES = {"passed", "passed_with_expected_drift"}
MARKER_CELL = "Z1"
MARKER_VALUE = "wolfxl_ooxml_fidelity_mutation"
STYLE_CELL = "AA1"
RENAMED_SHEET = "WolfXL Fidelity Rename"
MANIFEST_NAME = "manifest.json"
EXPECTED_ISSUE_KINDS_BY_MUTATION = {
    # Renaming a sheet is an intentional workbook semantic change. Feature
    # formulas and pivot cache worksheetSource sheet attrs may legitimately
    # change to keep pointing at the same sheet under its new title.
    "rename_first_sheet": {
        "charts_semantic_drift",
        "conditional_formatting_semantic_drift",
        "pivots_semantic_drift",
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
        elif mutation == "move_marker_range":
            worksheet = workbook[workbook.sheetnames[0]]
            worksheet["Z1"] = MARKER_VALUE
            worksheet["AA1"] = f"{MARKER_VALUE}_right"
            worksheet.move_range("Z1:AA1", rows=1, cols=0)
        elif mutation == "rename_first_sheet":
            workbook[workbook.sheetnames[0]].title = RENAMED_SHEET
        else:
            raise ValueError(f"unknown mutation: {mutation}")
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
        if issue.get("kind") in expected_kinds:
            expected.append(issue)
        else:
            unexpected.append(issue)
    return unexpected, expected


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
