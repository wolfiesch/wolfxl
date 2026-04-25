"""W4E.P6 regression: DualWorkbook captures fan-out exceptions.

Previously ``__getattr__``'s ``fan_out`` re-raised whichever side
raised first, killing the test before the other side could run. If
oracle raised, native never ran — and vice versa. Either way, the
diff harness was left with one or zero outputs to compare and could
falsely report "clean" because there was nothing structurally wrong
with the surviving file.

The fix captures both sides' exceptions in
``self._oracle_errors`` / ``self._native_errors``. Oracle-side
errors still propagate to the caller (so test bodies see the same
behavior as under ``WOLFXL_WRITER=oracle``); native-side errors are
captured only. The harness asserts both lists are empty.
"""
from __future__ import annotations

from pathlib import Path

import pytest


class _StubBackend:
    """Plain-Python stub mimicking the rust pyclass surface enough to
    exercise ``__getattr__`` fan-out. pyo3 pyclasses are read-only at the
    Python attribute level, so we can't monkey-patch ``wb._oracle.foo``
    directly — instead we sub in a ``_StubBackend`` whose methods can be
    swapped in-place to simulate per-call divergence."""

    def __init__(self, name: str) -> None:
        self._name = name
        self.calls: list[tuple[str, tuple, dict]] = []

    def __dir__(self) -> list[str]:
        return ["add_sheet", "write_cell_value", "save", "calls"]

    def add_sheet(self, *args, **kwargs) -> None:
        self.calls.append(("add_sheet", args, kwargs))

    def write_cell_value(self, *args, **kwargs) -> None:
        self.calls.append(("write_cell_value", args, kwargs))

    def save(self, *args, **kwargs) -> None:
        self.calls.append(("save", args, kwargs))


def _make_dual_with_stubs() -> object:
    """Build a DualWorkbook whose ``_oracle`` / ``_native`` are stubs.
    The dir() consistency check passes because both stubs expose the
    same attribute set."""
    from wolfxl._dual_workbook import DualWorkbook

    wb = DualWorkbook.__new__(DualWorkbook)
    wb._oracle = _StubBackend("oracle")  # type: ignore[attr-defined]
    wb._native = _StubBackend("native")  # type: ignore[attr-defined]
    wb._oracle_path = None
    wb._native_path = None
    wb._oracle_errors = []
    wb._native_errors = []
    return wb


def test_native_only_failure_is_captured() -> None:
    """If native rejects a call that oracle accepts, oracle's call
    completes and the native exception lands in ``_native_errors``.
    The user-visible call returns oracle's result — no exception."""
    wb = _make_dual_with_stubs()

    def _explode(*args: object, **kwargs: object) -> None:
        raise RuntimeError("simulated native rejection")

    wb._native.add_sheet = _explode  # type: ignore[attr-defined]

    # Caller perspective: the call does NOT raise; the native exception
    # is captured silently while oracle's call completes.
    wb.add_sheet("Probe")  # type: ignore[attr-defined]
    assert len(wb._native_errors) == 1  # type: ignore[attr-defined]
    captured = wb._native_errors[0]  # type: ignore[attr-defined]
    assert captured.method == "add_sheet"
    assert captured.args == ("Probe",)
    assert isinstance(captured.exception, RuntimeError)
    assert "simulated native rejection" in str(captured.exception)
    assert wb._oracle_errors == []  # type: ignore[attr-defined]
    # Oracle's call still ran.
    assert wb._oracle.calls == [("add_sheet", ("Probe",), {})]  # type: ignore[attr-defined]


def test_oracle_failure_propagates_and_is_recorded() -> None:
    """Oracle-side failures still propagate to the caller (matching
    pure-oracle behavior) AND land in ``_oracle_errors``. Native runs
    independently — if it accepts the call, no native error logged."""
    wb = _make_dual_with_stubs()

    def _explode(*args: object, **kwargs: object) -> None:
        raise RuntimeError("simulated oracle rejection")

    wb._oracle.add_sheet = _explode  # type: ignore[attr-defined]

    with pytest.raises(RuntimeError, match="simulated oracle rejection"):
        wb.add_sheet("Probe")  # type: ignore[attr-defined]
    assert len(wb._oracle_errors) == 1  # type: ignore[attr-defined]
    assert wb._oracle_errors[0].method == "add_sheet"  # type: ignore[attr-defined]
    # Native ran successfully despite oracle's failure.
    assert wb._native_errors == []  # type: ignore[attr-defined]
    assert wb._native.calls == [("add_sheet", ("Probe",), {})]  # type: ignore[attr-defined]


def test_clean_call_records_no_errors(tmp_path: Path) -> None:
    """The common case end-to-end with the real pyclasses: oracle and
    native both accept the call, save() writes both files, neither
    error list is touched."""
    from wolfxl._dual_workbook import DualWorkbook

    wb = DualWorkbook()
    wb.add_sheet("Probe")
    wb.write_cell_value("Probe", "A1", {"type": "string", "value": "hello"})

    out = tmp_path / "out.xlsx"
    wb.save(str(out))

    assert wb._oracle_errors == []
    assert wb._native_errors == []
    assert wb._oracle_path is not None and Path(wb._oracle_path).exists()
    assert wb._native_path is not None and Path(wb._native_path).exists()
