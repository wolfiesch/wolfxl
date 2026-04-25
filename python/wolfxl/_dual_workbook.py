"""``DualWorkbook`` — fan-out wrapper used by the differential writer harness.

Holds an oracle (``RustXlsxWriterBook``) and a native (``NativeWorkbook``)
pyclass instance, fans every method call to both, and on ``save()`` writes
two sibling xlsx files. The harness reads ``_oracle_path`` / ``_native_path``
after save and feeds them into the 3-layer diff (canonicalize/xml_tree/semantic).

This class is NOT a pyo3 pyclass. It is a pure-Python wrapper because the
fan-out is trivial in Python and trying to clone the whole pymethod surface
in Rust would require a third pyclass that mirrors the other two — exactly
the kind of drift the constructor's ``dir()`` consistency check catches.

The ``save(path)`` method writes oracle to ``<stem>.oracle.xlsx`` and native
to ``<stem>.native.xlsx`` (the original path is not written). The harness
fixture in ``tests/diffwriter/conftest.py`` reads the two attributes after
``save()`` returns. Diff triggering lives in the test runner — keeping
``DualWorkbook`` decoupled from harness internals.

Native-side errors are captured rather than re-raised: if oracle accepts
a call but native rejects it (or vice versa), the divergence shows up in
``self._native_errors`` / ``self._oracle_errors`` rather than killing the
fan-out mid-flight. The harness asserts both lists are empty as part of
the per-case gate. Without capture, an asymmetric rejection would crash
the test on whichever backend was called second, and the diff harness
would never run on the surviving file — so the diff would falsely
report "clean" because there was nothing to compare.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from . import _rust  # type: ignore[attr-defined]


@dataclass
class _CapturedError:
    """A method-call exception captured during DualWorkbook fan-out."""
    method: str
    args: tuple[Any, ...]
    kwargs: dict[str, Any]
    exception: BaseException

    def __repr__(self) -> str:
        return (
            f"_CapturedError(method={self.method!r}, "
            f"exception={type(self.exception).__name__}: {self.exception})"
        )


class DualWorkbook:
    """Fan-out wrapper writing both oracle and native xlsx files per ``save()``."""

    def __init__(self) -> None:
        self._oracle: Any = _rust.RustXlsxWriterBook()
        self._native: Any = _rust.NativeWorkbook()
        self._oracle_path: str | None = None
        self._native_path: str | None = None
        # Captured exceptions from fan-out — the harness asserts both
        # are empty after ``save()`` to detect asymmetric rejection.
        self._oracle_errors: list[_CapturedError] = []
        self._native_errors: list[_CapturedError] = []

        # Sanity gate — both backends must expose the same pymethod surface.
        # Drift here means a future pymethod was added to one backend but
        # not the other, and every harness test would silently miss the new
        # behavior on whichever side lacks it. Fail loud at construction.
        oracle_methods = {m for m in dir(self._oracle) if not m.startswith("_")}
        native_methods = {m for m in dir(self._native) if not m.startswith("_")}
        missing_on_native = oracle_methods - native_methods
        missing_on_oracle = native_methods - oracle_methods
        if missing_on_native or missing_on_oracle:
            raise RuntimeError(
                "DualWorkbook surface drift detected. "
                f"Missing on native: {sorted(missing_on_native)}. "
                f"Missing on oracle: {sorted(missing_on_oracle)}."
            )

    def __getattr__(self, name: str) -> Any:
        # ``__getattr__`` only fires on the normal-lookup miss path. Anything
        # set in ``__init__`` (``_oracle``, ``_native``, ``_oracle_path``,
        # ``_native_path``, ``_oracle_errors``, ``_native_errors``) and
        # anything defined on the class (``save``) is found by normal
        # lookup before this runs.
        oracle_attr = getattr(self._oracle, name)
        native_attr = getattr(self._native, name)
        if not callable(oracle_attr):
            # Property-style access — return oracle's value. Both should
            # match for read-only state; if they ever diverge that's a real
            # bug we want surfaced via the diff harness, not here.
            return oracle_attr

        def fan_out(*args: Any, **kwargs: Any) -> Any:
            # Capture each side's exception independently so the other
            # side still runs. Without this, a backend-specific bug
            # would prevent the diff harness from ever comparing the
            # two outputs — silently masking divergence as "clean."
            o_result = None
            o_exc: BaseException | None = None
            try:
                o_result = oracle_attr(*args, **kwargs)
            except BaseException as exc:  # noqa: BLE001 — capture EVERYTHING
                o_exc = exc
                self._oracle_errors.append(
                    _CapturedError(name, args, dict(kwargs), exc)
                )

            try:
                native_attr(*args, **kwargs)
            except BaseException as exc:  # noqa: BLE001
                self._native_errors.append(
                    _CapturedError(name, args, dict(kwargs), exc)
                )

            # If oracle raised, propagate that to the caller — the test
            # body still sees the same behavior it would under
            # ``WOLFXL_WRITER=oracle`` alone. Native errors are captured
            # only; the harness inspects ``_native_errors`` post-fan-out.
            if o_exc is not None:
                raise o_exc
            return o_result

        fan_out.__name__ = name
        return fan_out

    def save(self, path: str) -> None:
        """Save oracle to ``<stem>.oracle.xlsx`` and native to ``<stem>.native.xlsx``.

        The original ``path`` is intentionally NOT written — the harness
        reads the two sibling files via ``self._oracle_path`` /
        ``self._native_path``. ``<stem>`` strips a trailing ``.xlsx`` so a
        caller passing ``out.xlsx`` gets clean ``out.oracle.xlsx`` instead of
        ``out.xlsx.oracle.xlsx``.
        """
        p = Path(path)
        stem = p.with_suffix("") if p.suffix.lower() == ".xlsx" else p
        oracle_path = f"{stem}.oracle.xlsx"
        native_path = f"{stem}.native.xlsx"
        self._oracle.save(oracle_path)
        self._native.save(native_path)
        self._oracle_path = oracle_path
        self._native_path = native_path
