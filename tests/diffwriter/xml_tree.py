"""Layer 2 — XML-structural diff (blocking).

Stubs land in the W4C contract sub-commit. The implementation lands in the
W4C implementation sub-commit. Until then, ``compute_diffs`` returns an
empty list so harness assertions that ignore Layer 2 results compile.
"""
from __future__ import annotations

from pathlib import Path


def compute_diffs(oracle_path: Path, native_path: Path) -> list[str]:
    """Return a list of human-readable diff strings (empty when clean).

    Stub returns ``[]`` until the implementation sub-commit lands.
    """
    return []


def assert_structural_clean(oracle_path: Path, native_path: Path) -> None:
    """Assert no structural differences. Stub is a no-op for now."""
    diffs = compute_diffs(oracle_path, native_path)
    if diffs:
        raise AssertionError(
            f"Layer 2 structural differences:\n" + "\n".join(diffs)
        )
