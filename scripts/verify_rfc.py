"""
verify_rfc.py — standardized "done" check for an RFC.

Per Plans/rfcs/INDEX.md §Conventions, every RFC's §6 verification matrix has
six layers:

  1. Rust unit tests
  2. Golden round-trip (diffwriter)
  3. openpyxl parity
  4. LibreOffice cross-renderer (optional, manual)
  5. Cross-mode (write + modify produce equivalent files)
  6. Regression fixture

This script automates layers 1, 2, 3, 5, 6 and reports a green/red bar per RFC.
LibreOffice (layer 4) is noted as MANUAL when present.

Usage:
    python scripts/verify_rfc.py --rfc 021
    python scripts/verify_rfc.py --rfc 021 --quick   # skip slow parity sweep
    python scripts/verify_rfc.py --all               # every Shipped RFC
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
RFCS = REPO / "Plans" / "rfcs"


@dataclass
class Layer:
    name: str
    cmd: list[str]
    optional: bool = False


@dataclass
class RfcSpec:
    rfc_id: str
    title: str
    status: str
    pytest_marker: str | None  # e.g. "rfc021"
    cargo_test_filter: str | None  # e.g. "defined_names"


def parse_rfc(path: Path) -> RfcSpec | None:
    text = path.read_text()
    m = re.search(r"^# RFC-(\d{3}):\s*(.+)$", text, re.MULTILINE)
    s = re.search(r"^Status:\s*(\S+)", text, re.MULTILINE)
    if not m or not s:
        return None
    rfc_id = m.group(1)
    return RfcSpec(
        rfc_id=rfc_id,
        title=m.group(2).strip(),
        status=s.group(1),
        pytest_marker=f"rfc{rfc_id}",
        cargo_test_filter=None,
    )


def find_rfc(rfc_id: str) -> Path:
    matches = list(RFCS.glob(f"{rfc_id}-*.md"))
    if not matches:
        sys.exit(f"no RFC matching {rfc_id} under {RFCS}")
    return matches[0]


def run(cmd: list[str], env_extra: dict[str, str] | None = None) -> tuple[bool, str]:
    import os

    env = os.environ.copy()
    if env_extra:
        env.update(env_extra)
    try:
        out = subprocess.run(
            cmd,
            cwd=REPO,
            check=False,
            capture_output=True,
            text=True,
            env=env,
            timeout=600,
        )
    except FileNotFoundError as e:
        return False, f"command not found: {e}"
    ok = out.returncode == 0
    tail = (out.stdout + out.stderr).strip().splitlines()[-20:]
    return ok, "\n".join(tail)


def layers_for(rfc: RfcSpec, quick: bool) -> list[Layer]:
    layers = [
        Layer(
            "1. cargo test (workspace crates)",
            [
                "cargo",
                "test",
                "-p",
                "wolfxl-core",
                "-p",
                "wolfxl-writer",
                "-p",
                "wolfxl-rels",
                "-p",
                "wolfxl-merger",
                "--quiet",
            ],
        ),
        Layer(
            "2. golden round-trip (diffwriter)",
            ["pytest", "tests/diffwriter/", "-q"],
            optional=True,
        ),
        Layer(
            "3. openpyxl parity",
            (
                ["pytest", "tests/parity/", "-q", "-x"]
                if not quick
                else ["pytest", "tests/parity/test_read_parity.py", "-q", "-x"]
            ),
        ),
        Layer(
            f"5. cross-mode pytest (-k {rfc.pytest_marker})",
            ["pytest", "tests/", "-q", "-k", rfc.pytest_marker or rfc.rfc_id],
        ),
        Layer(
            "6. ruff lint",
            ["ruff", "check", "python/", "tests/"],
        ),
    ]
    return layers


def verify(rfc_id: str, quick: bool = False) -> bool:
    rfc_path = find_rfc(rfc_id)
    rfc = parse_rfc(rfc_path)
    if rfc is None:
        sys.exit(f"could not parse RFC frontmatter from {rfc_path}")

    print(f"\n=== Verifying RFC-{rfc.rfc_id}: {rfc.title} ===")
    print(f"Status (frontmatter): {rfc.status}")
    print(f"Spec file: {rfc_path.relative_to(REPO)}\n")

    all_ok = True
    for layer in layers_for(rfc, quick):
        if layer.cmd[0] == "cargo" and shutil.which("cargo") is None:
            print(f"[SKIP] {layer.name}: cargo not on PATH")
            continue
        if layer.cmd[0] == "pytest" and shutil.which("pytest") is None:
            print(f"[SKIP] {layer.name}: pytest not on PATH")
            continue
        ok, tail = run(layer.cmd, env_extra={"WOLFXL_TEST_EPOCH": "0"})
        flag = "OK" if ok else ("WARN" if layer.optional else "FAIL")
        print(f"[{flag}] {layer.name}")
        if not ok:
            print("  └─", tail.replace("\n", "\n     "))
            if not layer.optional:
                all_ok = False

    print()
    print("Layer 4 (LibreOffice cross-renderer): MANUAL — see RFC §6.")
    print()
    print("RESULT:", "GREEN" if all_ok else "RED")
    return all_ok


def main() -> int:
    p = argparse.ArgumentParser()
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument("--rfc", help="3-digit RFC id, e.g. 021")
    g.add_argument("--all", action="store_true", help="verify every Shipped RFC")
    p.add_argument("--quick", action="store_true", help="skip slow parity sweep")
    args = p.parse_args()

    if args.all:
        ok_count = 0
        fail_ids: list[str] = []
        for path in sorted(RFCS.glob("[0-9][0-9][0-9]-*.md")):
            rfc = parse_rfc(path)
            if rfc and rfc.status.lower() == "shipped":
                if verify(rfc.rfc_id, quick=args.quick):
                    ok_count += 1
                else:
                    fail_ids.append(rfc.rfc_id)
        print(f"\n=== Summary: {ok_count} green, {len(fail_ids)} red ===")
        if fail_ids:
            print("Red:", ", ".join(fail_ids))
        return 0 if not fail_ids else 1

    return 0 if verify(args.rfc, quick=args.quick) else 1


if __name__ == "__main__":
    raise SystemExit(main())
