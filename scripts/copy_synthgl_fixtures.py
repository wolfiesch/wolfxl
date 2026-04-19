#!/usr/bin/env python3
"""Refresh the SynthGL fixture snapshot used by the parity harness.

Usage::

    python scripts/copy_synthgl_fixtures.py \
        --synthgl-root /Users/wolfgangschoenberger/Projects/SynthGL \
        [--per-category 3]

Idempotent: files that already exist with matching size are skipped.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

CATEGORIES = (
    "aging",
    "flat_register",
    "rollforward",
    "time_series",
    "cross_ref",
    "key_value",
    "stress",
)


def parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    p.add_argument(
        "--synthgl-root",
        type=Path,
        default=Path("/Users/wolfgangschoenberger/Projects/SynthGL"),
    )
    p.add_argument("--per-category", type=int, default=3)
    p.add_argument(
        "--dest-root",
        type=Path,
        default=Path(__file__).parent.parent / "tests" / "parity" / "fixtures" / "synthgl_snapshot",
    )
    return p.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    src_root = args.synthgl_root / "tests" / "app" / "fixtures" / "ingestion"
    if not src_root.exists():
        print(f"Source not found: {src_root}", file=sys.stderr)
        return 1

    copied = 0
    skipped = 0
    for category in CATEGORIES:
        src = src_root / category
        if not src.exists():
            continue
        dst = args.dest_root / category
        dst.mkdir(parents=True, exist_ok=True)
        xlsx_files = sorted(src.glob("*.xlsx"))[: args.per_category]
        for f in xlsx_files:
            target = dst / f.name
            if target.exists() and target.stat().st_size == f.stat().st_size:
                skipped += 1
                continue
            shutil.copy2(f, target)
            copied += 1

    print(f"Copied {copied}, skipped {skipped} (already present).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
