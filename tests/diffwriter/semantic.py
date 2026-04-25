"""Layer 3 — semantic diff via ``tests/parity/_scoring.py`` (blocking on HARD).

Stubs land in the W4C contract sub-commit. The implementation lands in the
W4C implementation sub-commit. Until then, ``assert_semantic_clean`` is a
no-op so harness assertions that ignore Layer 3 results compile.
"""
from __future__ import annotations

from pathlib import Path


def assert_semantic_clean(oracle_path: Path, native_path: Path) -> None:
    """Assert no HARD-tier semantic differences. Stub is a no-op for now."""
    return None
