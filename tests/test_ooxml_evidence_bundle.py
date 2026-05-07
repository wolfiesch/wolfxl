from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from types import ModuleType


def _load_bundle_module() -> ModuleType:
    script = (
        Path(__file__).resolve().parents[1]
        / "scripts"
        / "audit_ooxml_evidence_bundle.py"
    )
    spec = importlib.util.spec_from_file_location("audit_ooxml_evidence_bundle", script)
    assert spec is not None
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


bundle = _load_bundle_module()


def test_bundle_audit_accepts_expected_report_values(tmp_path: Path) -> None:
    report_path = tmp_path / "coverage.json"
    report_path.write_text(
        json.dumps(
            {
                "ready": True,
                "fixture_count": 22,
                "surfaces": {
                    "pivot_slicer_preservation": {
                        "clear": True,
                    }
                },
            }
        )
    )
    manifest = tmp_path / "bundle.json"
    manifest.write_text(
        json.dumps(
            {
                "reports": [
                    {
                        "name": "coverage",
                        "path": str(report_path),
                        "producer": "uv run --no-sync python scripts/example.py",
                        "expect": [
                            {"path": "ready", "equals": True},
                            {"path": "fixture_count", "at_least": 20},
                            {
                                "path": "surfaces.pivot_slicer_preservation.clear",
                                "equals": True,
                            },
                        ],
                    }
                ]
            }
        )
    )

    audit = bundle.audit_bundle(manifest)

    assert audit["ready"] is True
    assert audit["issue_count"] == 0
    assert audit["producer_count"] == 1
    assert audit["reports"][0]["producer"] == "uv run --no-sync python scripts/example.py"


def test_bundle_audit_reports_missing_and_stale_evidence(tmp_path: Path) -> None:
    report_path = tmp_path / "coverage.json"
    report_path.write_text(json.dumps({"ready": False}))
    manifest = tmp_path / "bundle.json"
    manifest.write_text(
        json.dumps(
            {
                "reports": [
                    {
                        "name": "coverage",
                        "path": str(report_path),
                        "producer": "uv run --no-sync python scripts/example.py",
                        "expect": [
                            {"path": "ready", "equals": True},
                            {"path": "fixture_count", "equals": 22},
                        ],
                    },
                    {
                        "name": "missing",
                        "path": str(tmp_path / "missing.json"),
                        "producer": "uv run --no-sync python scripts/missing.py",
                        "expect": [
                            {"path": "ready", "equals": True},
                        ],
                    },
                ]
            }
        )
    )

    audit = bundle.audit_bundle(manifest)

    assert audit["ready"] is False
    assert audit["issue_count"] == 3
    messages = [issue["message"] for issue in audit["issues"]]
    assert "expected ready == True, got False" in messages
    assert any("missing path" in message for message in messages)
    assert "report file is missing" in messages


def test_bundle_audit_reports_missing_producer(tmp_path: Path) -> None:
    report_path = tmp_path / "coverage.json"
    report_path.write_text(json.dumps({"ready": True}))
    manifest = tmp_path / "bundle.json"
    manifest.write_text(
        json.dumps(
            {
                "reports": [
                    {
                        "name": "coverage",
                        "path": str(report_path),
                        "expect": [
                            {"path": "ready", "equals": True},
                        ],
                    }
                ]
            }
        )
    )

    audit = bundle.audit_bundle(manifest)

    assert audit["ready"] is False
    assert audit["producer_count"] == 0
    assert audit["issues"] == [
        {
            "report": "coverage",
            "path": str(report_path),
            "check": "producer",
            "message": "producer command is missing",
        }
    ]


def test_bundle_strict_cli_fails_for_stale_evidence(tmp_path: Path, capsys) -> None:
    report_path = tmp_path / "coverage.json"
    report_path.write_text(json.dumps({"ready": False}))
    manifest = tmp_path / "bundle.json"
    manifest.write_text(
        json.dumps(
            {
                "reports": [
                    {
                        "name": "coverage",
                        "path": str(report_path),
                        "producer": "uv run --no-sync python scripts/example.py",
                        "expect": [
                            {"path": "ready", "equals": True},
                        ],
                    }
                ]
            }
        )
    )

    code = bundle.main([str(manifest), "--strict"])

    captured = capsys.readouterr()
    assert code == 1
    payload = json.loads(captured.out)
    assert payload["ready"] is False
