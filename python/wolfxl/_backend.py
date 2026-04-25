"""Backend dispatcher reading WOLFXL_WRITER env var.

W4A grammar: ``'oracle' | 'native'``.
W4C grammar: adds ``'both'`` (DualWorkbook fan-out, used by the diff harness)
and ``'auto'`` (alias to native — no per-feature routing in MVP; Wave 5 may
flip the default once the 30-day soak is clean).
"""
from __future__ import annotations

import os
from typing import Any

from . import _rust  # type: ignore[attr-defined]

_VALID = ("oracle", "native", "both", "auto")


def make_writer() -> Any:
    """Construct the active write-mode backend per ``WOLFXL_WRITER``."""
    # TODO(W5): flip the default from "oracle" to "native" (or "auto",
    # which currently aliases native) AFTER the 30-day soak is clean. Once
    # ``RustXlsxWriterBook`` is removed in Wave 5, ``"oracle"`` and
    # ``"both"`` arms must also be deleted — leaving them in would crash
    # on the missing pyclass at the first ``make_writer()`` call.
    choice = os.environ.get("WOLFXL_WRITER", "oracle").lower()
    if choice == "oracle":
        return _rust.RustXlsxWriterBook()
    if choice == "native":
        return _rust.NativeWorkbook()
    if choice == "both":
        # Lazy import — DualWorkbook lives in pure Python and only loads
        # when the harness asks for it.
        from ._dual_workbook import DualWorkbook
        return DualWorkbook()
    if choice == "auto":
        # MVP aliases auto -> native. Wave 5 may flip the default to "auto"
        # once a 30-day soak under "both" is clean; future waves may refine
        # auto into per-feature routing if the diff harness surfaces a real
        # need for it.
        return _rust.NativeWorkbook()
    raise ValueError(
        f"WOLFXL_WRITER={choice!r} is not a valid backend. "
        f"Choose one of {_VALID}."
    )
