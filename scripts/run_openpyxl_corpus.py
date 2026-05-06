#!/usr/bin/env python3
"""Run a cached openpyxl test corpus with ``openpyxl`` shimmed to ``wolfxl``.

This is intentionally non-networked. Point it at a local openpyxl source
checkout or a previously vendored test directory:

    uv run --no-sync python scripts/run_openpyxl_corpus.py --corpus /tmp/openpyxl/tests

The runner writes a machine-readable JSON summary and returns pytest's exit
code when a corpus is present. With no corpus it exits 0 and reports
``status=skipped`` unless ``--require-corpus`` is set.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from textwrap import dedent


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CORPUS = ROOT / "tests" / "vendored_openpyxl"
DEFAULT_ALLOWLIST = ROOT / "tests" / "vendored_openpyxl_allowlist.json"
DEFAULT_REPORT = ROOT / "logs" / "openpyxl-corpus-summary.json"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--corpus", type=Path, default=DEFAULT_CORPUS)
    parser.add_argument("--allowlist", type=Path, default=DEFAULT_ALLOWLIST)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--require-corpus", action="store_true")
    parser.add_argument("pytest_args", nargs=argparse.REMAINDER)
    args = parser.parse_args()
    pytest_args = list(args.pytest_args)
    if pytest_args and pytest_args[0] == "--":
        pytest_args = pytest_args[1:]

    corpus = args.corpus
    report = args.report
    report.parent.mkdir(parents=True, exist_ok=True)

    if not corpus.exists():
        payload = {
            "status": "missing_corpus",
            "corpus": str(corpus),
            "allowlist": _load_allowlist(args.allowlist),
            "message": "No cached openpyxl corpus found; pass --corpus or vendor tests first.",
        }
        report.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n")
        print(payload["message"])
        return 2 if args.require_corpus else 0

    with tempfile.TemporaryDirectory(prefix="wolfxl-openpyxl-corpus-") as tmp:
        sitecustomize = Path(tmp) / "sitecustomize.py"
        openpyxl_tests_dir = _openpyxl_tests_dir(corpus)
        sitecustomize.write_text(
            dedent(
                f"""
                import sys
                import types
                import wolfxl

                sys.modules.setdefault("openpyxl", wolfxl)
                # Upstream openpyxl tests import helpers as ``openpyxl.tests``.
                # Keep library imports routed to wolfxl, and expose only the
                # test-helper package so missing wolfxl shims do not silently
                # fall back to upstream openpyxl implementation modules.
                tests_dir = {str(openpyxl_tests_dir)!r}
                if tests_dir:
                    tests_pkg = types.ModuleType("openpyxl.tests")
                    tests_pkg.__path__ = [tests_dir]
                    sys.modules.setdefault("openpyxl.tests", tests_pkg)
                    setattr(wolfxl, "tests", tests_pkg)
                """
            )
        )
        env = os.environ.copy()
        env["PYTHONPATH"] = f"{tmp}{os.pathsep}{env.get('PYTHONPATH', '')}"
        cmd = [
            sys.executable,
            "-m",
            "pytest",
            str(corpus),
            "-q",
            *pytest_args,
        ]
        proc = subprocess.run(cmd, cwd=ROOT, env=env, text=True, capture_output=True)

    payload = {
        "status": "passed" if proc.returncode == 0 else "failed",
        "returncode": proc.returncode,
        "corpus": str(corpus),
        "allowlist": _load_allowlist(args.allowlist),
        "stdout_tail": proc.stdout[-4000:],
        "stderr_tail": proc.stderr[-4000:],
    }
    report.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n")
    print(proc.stdout, end="")
    print(proc.stderr, end="", file=sys.stderr)
    print(f"wrote {report}")
    return proc.returncode


def _load_allowlist(path: Path) -> dict[str, object]:
    if not path.exists():
        return {"entries": []}
    try:
        return json.loads(path.read_text())
    except json.JSONDecodeError as exc:
        return {"error": f"invalid allowlist JSON: {exc}"}


def _openpyxl_tests_dir(corpus: Path) -> Path | str:
    """Return the upstream ``openpyxl.tests`` package dir for ``corpus``.

    Typical input is ``.../openpyxl/tests`` from an upstream source archive.
    When callers point at a custom directory that is not nested under an
    ``openpyxl/tests`` package, return an empty path sentinel so the shim
    remains compatible with ad-hoc corpora.
    """
    resolved = corpus.resolve()
    if resolved.name == "tests" and resolved.parent.name == "openpyxl":
        return resolved
    return ""


if __name__ == "__main__":
    raise SystemExit(main())
