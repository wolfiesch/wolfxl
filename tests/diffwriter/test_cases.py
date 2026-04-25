"""Pytest driver — one parametrized test per case across all case modules.

Discovers every ``CASES`` list in ``tests.diffwriter.cases.*`` at collection
time and runs Layer 2 (structural) + Layer 3 (semantic HARD) assertions on
the oracle+native pair produced by ``DualWorkbook``. Layer 1 byte-canonical
parity is computed but reported only — gold-star, not blocking.

Each case is one (id, build_fn) pair; the driver:
  1. Sets ``WOLFXL_WRITER=both`` so ``wolfxl.Workbook()`` routes through
     ``DualWorkbook``.
  2. Calls ``build_fn(wb)`` to populate the workbook via the public API.
  3. Saves to a tmp directory; ``DualWorkbook.save`` writes ``<stem>.oracle.xlsx``
     and ``<stem>.native.xlsx`` and stamps the paths on the wrapper.
  4. Asserts Layer 2 clean (after platform-gap filtering) and Layer 3
     HARD-tier clean.
"""
from __future__ import annotations

import importlib
import pkgutil
from pathlib import Path
from typing import Any, Callable

import pytest

from . import cases as cases_pkg
from .canonicalize import compare_canonical
from .semantic import assert_semantic_clean
from .xml_tree import assert_structural_clean


def _discover_cases() -> list[tuple[str, Callable[[Any], None]]]:
    """Walk ``tests/diffwriter/cases`` and return every (id, build) pair."""
    out: list[tuple[str, Callable[[Any], None]]] = []
    for mod_info in pkgutil.iter_modules(cases_pkg.__path__):
        mod = importlib.import_module(f"tests.diffwriter.cases.{mod_info.name}")
        cases = getattr(mod, "CASES", None)
        if cases is None:
            continue
        for case_id, build_fn in cases:
            out.append((case_id, build_fn))
    return out


_ALL_CASES = _discover_cases()


# Case IDs whose dual-backend diff cannot pass because oracle itself has a
# documented bug that native intentionally fixes. ``pytest.xfail`` is called
# inside the test body so these still execute the build + save paths (which
# proves the round-trip works on each backend), but the cross-backend
# assertion is an expected failure. The reverse — these unexpectedly passing
# — would indicate oracle has been patched upstream and the rewrite's
# regression coverage needs re-evaluating.
_ORACLE_BUGS_DOCUMENTED: dict[str, str] = {
    "comments_multi_author_insertion_order": (
        "rust_xlsxwriter alphabetizes <authors> via BTreeMap; native "
        "preserves IndexMap insertion order. This case is the canonical "
        "regression test for the bug that motivated the writer rewrite. "
        "Native-side correctness is verified directly in "
        "tests/parity/test_write_parity.py and the Wave 3 integration test "
        "wave3_rich_features_roundtrip; this dual-backend comparison cannot "
        "pass and is xfailed so a passing result would also flag a regression."
    ),
}


@pytest.mark.parametrize(
    "case_id,build_fn",
    _ALL_CASES,
    ids=[c[0] for c in _ALL_CASES],
)
def test_dual_backend_parity(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    case_id: str,
    build_fn: Callable[[Any], None],
) -> None:
    """Build the case under ``WOLFXL_WRITER=both`` and assert L2 + L3 clean."""
    monkeypatch.setenv("WOLFXL_WRITER", "both")
    import wolfxl

    wb = wolfxl.Workbook()
    build_fn(wb)
    output_stem = tmp_path / "out.xlsx"
    wb.save(str(output_stem))

    dual = wb._rust_writer
    oracle_path = Path(dual._oracle_path)
    native_path = Path(dual._native_path)
    assert oracle_path.exists(), f"{case_id}: oracle xlsx missing at {oracle_path}"
    assert native_path.exists(), f"{case_id}: native xlsx missing at {native_path}"

    # W4E.P6 gate: DualWorkbook captures asymmetric exceptions instead of
    # re-raising. If oracle accepted a call that native rejected (or vice
    # versa), the divergence is recorded but didn't kill fan-out — we
    # surface it here so the harness can't falsely report "clean."
    assert dual._oracle_errors == [], (
        f"{case_id}: oracle raised on fan-out: {dual._oracle_errors}"
    )
    assert dual._native_errors == [], (
        f"{case_id}: native raised on fan-out: {dual._native_errors}"
    )

    if case_id in _ORACLE_BUGS_DOCUMENTED:
        pytest.xfail(_ORACLE_BUGS_DOCUMENTED[case_id])

    assert_structural_clean(oracle_path, native_path)
    assert_semantic_clean(oracle_path, native_path)


@pytest.mark.parametrize(
    "case_id,build_fn",
    _ALL_CASES,
    ids=[c[0] for c in _ALL_CASES],
)
def test_layer1_canonical_report(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    case_id: str,
    build_fn: Callable[[Any], None],
    request: pytest.FixtureRequest,
) -> None:
    """Compute Layer 1 canonical-byte parity and stash the result on the node.

    Layer 1 is gold-star — never fatal. We record the per-case outcome on
    the request node so a future ``python -m diffwriter status`` can read it
    without re-running the cases.
    """
    monkeypatch.setenv("WOLFXL_WRITER", "both")
    import wolfxl

    wb = wolfxl.Workbook()
    build_fn(wb)
    output_stem = tmp_path / "out.xlsx"
    wb.save(str(output_stem))

    dual = wb._rust_writer
    oracle_path = Path(dual._oracle_path)
    native_path = Path(dual._native_path)
    mismatches = compare_canonical(oracle_path, native_path)
    request.node.user_properties.append((case_id, len(mismatches)))
