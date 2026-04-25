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
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from . import _rust  # type: ignore[attr-defined]


class DualWorkbook:
    """Fan-out wrapper writing both oracle and native xlsx files per ``save()``."""

    def __init__(self) -> None:
        self._oracle: Any = _rust.RustXlsxWriterBook()
        self._native: Any = _rust.NativeWorkbook()
        self._oracle_path: str | None = None
        self._native_path: str | None = None

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
        # ``_native_path``) and anything defined on the class (``save``) is
        # found by normal lookup before this runs.
        oracle_attr = getattr(self._oracle, name)
        native_attr = getattr(self._native, name)
        if not callable(oracle_attr):
            # Property-style access — return oracle's value. Both should
            # match for read-only state; if they ever diverge that's a real
            # bug we want surfaced via the diff harness, not here.
            return oracle_attr

        def fan_out(*args: Any, **kwargs: Any) -> Any:
            o_result = oracle_attr(*args, **kwargs)
            native_attr(*args, **kwargs)
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
