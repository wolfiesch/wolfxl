#!/usr/bin/env python3
"""Run narrow interactive Excel probes for OOXML fidelity fixtures."""

from __future__ import annotations

import argparse
import json
import shutil
import sys
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
SUPPORTED_PROBES = (
    "macro_project_presence",
    "embedded_control_openability",
    "external_link_update_prompt",
)
PROBE_FEATURE_KEYS = {
    "macro_project_presence": "vba",
    "embedded_control_openability": "embedded_object",
    "external_link_update_prompt": "external_link",
}


@dataclass
class InteractiveProbeResult:
    fixture: str
    probe: str
    mutation: str
    app: str
    status: str
    output: str | None
    message: str


def run_interactive_probes(
    fixture_dir: Path,
    output_dir: Path,
    probes: tuple[str, ...] = SUPPORTED_PROBES,
    mutation: str = SOURCE_MUTATION,
    timeout: int = 90,
) -> dict:
    fixture_dir = fixture_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    results: list[InteractiveProbeResult] = []
    for entry in run_ooxml_fidelity_mutations.discover_fixtures(fixture_dir):
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
                )
            )

    report = {
        "fixture_dir": str(fixture_dir),
        "output_dir": str(output_dir.resolve()),
        "probes": list(probes),
        "mutation": mutation,
        "result_count": len(results),
        "failure_count": sum(
            1 for result in results if result.status not in PASSING_STATUSES
        ),
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
) -> InteractiveProbeResult:
    if probe not in SUPPORTED_PROBES:
        return InteractiveProbeResult(
            fixture=fixture_label,
            probe=probe,
            mutation=mutation,
            app="excel",
            status="failed",
            output=None,
            message=f"unsupported probe: {probe}",
        )

    work = (
        output_dir
        / run_ooxml_fidelity_mutations._safe_stem(
            Path(fixture_label).with_suffix("").as_posix()
        )
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
        mutation=mutation,
        app="excel",
        status="passed",
        output=str(probe_path),
        message=f"Microsoft Excel opened workbook and {part_label} is present",
    )


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
            name.startswith(("xl/embeddings/", "xl/ctrlProps/", "xl/activeX/"))
            for name in names
        )
    if probe == "external_link_update_prompt":
        return any(name.startswith("xl/externalLinks/") for name in names)
    return False


def _probe_part_label(probe: str) -> str:
    if probe == "macro_project_presence":
        return "xl/vbaProject.bin"
    if probe == "embedded_control_openability":
        return "embedded/control OOXML parts"
    if probe == "external_link_update_prompt":
        return "external-link OOXML parts"
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
    parser.add_argument("--timeout", type=int, default=90)
    args = parser.parse_args(argv)
    probes = tuple(args.probes) if args.probes else SUPPORTED_PROBES

    report = run_interactive_probes(
        args.fixture_dir,
        args.output_dir,
        probes=probes,
        mutation=args.mutation,
        timeout=args.timeout,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if report["failure_count"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
