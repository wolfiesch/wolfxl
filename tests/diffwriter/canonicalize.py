"""Layer 1 — byte-canonical diff (gold-star, non-blocking).

Stubs land in the W4C contract sub-commit. The implementation lands in the
W4C implementation sub-commit. Until then, ``canonical_part_hashes`` returns
an empty mapping so harness assertions that ignore Layer 1 results compile.
"""
from __future__ import annotations

from pathlib import Path
from typing import Mapping


def canonical_part_hashes(xlsx_path: Path, fuzzy: list[str]) -> Mapping[str, str]:
    """Return ``{part_path: sha256(canonical_bytes)}`` for every part of an xlsx.

    Stub returns an empty mapping until the implementation sub-commit lands.
    """
    return {}
