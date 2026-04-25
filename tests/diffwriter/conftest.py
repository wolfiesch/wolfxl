"""Pytest fixtures for the differential writer harness.

The session-scoped autouse fixture below forces ``WOLFXL_TEST_EPOCH=0`` so
ZIP entry mtimes in both backends are deterministic — without this, Layer 1
byte canonicalization is impossible (mtimes drift second-to-second). The
``dual_workbook`` fixture is the canonical entry point for harness tests:
it constructs a ``DualWorkbook`` and yields it together with the output
path stem; the test populates the workbook, calls ``save()``, and gets
``_oracle_path`` + ``_native_path`` to feed into ``assert_layers_clean``.
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Iterator

import pytest


@pytest.fixture(autouse=True, scope="session")
def _force_test_epoch() -> Iterator[None]:
    """Force ``WOLFXL_TEST_EPOCH=0`` so ZIP mtimes are deterministic across
    both backends. Without this Layer 1 byte parity is impossible.

    Restores the previous value (or unsets) on teardown so non-harness
    tests that import this conftest indirectly are unaffected.
    """
    prev = os.environ.get("WOLFXL_TEST_EPOCH")
    os.environ["WOLFXL_TEST_EPOCH"] = "0"
    try:
        yield
    finally:
        if prev is None:
            os.environ.pop("WOLFXL_TEST_EPOCH", None)
        else:
            os.environ["WOLFXL_TEST_EPOCH"] = prev


@pytest.fixture
def dual_workbook(tmp_path: Path):
    """Yield ``(DualWorkbook, output_stem)``.

    The build closure populates the wb, then the test calls
    ``wb.save(str(output_stem))``. After save, ``wb._oracle_path`` and
    ``wb._native_path`` carry the two written files.
    """
    from wolfxl._dual_workbook import DualWorkbook

    wb = DualWorkbook()
    yield wb, tmp_path / "out.xlsx"
