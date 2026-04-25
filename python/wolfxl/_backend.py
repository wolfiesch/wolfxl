"""Backend dispatcher reading WOLFXL_WRITER env var.

W4A grammar: ``'oracle' | 'native'``. Default is ``'oracle'`` for safety.
W4C will add ``'both'`` and ``'auto'`` once the diff harness is in place.
"""
from __future__ import annotations

import os
from typing import Any

from . import _rust  # type: ignore[attr-defined]

_VALID = ("oracle", "native")


def make_writer() -> Any:
    """Construct the active write-mode Rust pyclass per ``WOLFXL_WRITER``."""
    choice = os.environ.get("WOLFXL_WRITER", "oracle").lower()
    if choice == "oracle":
        return _rust.RustXlsxWriterBook()
    if choice == "native":
        return _rust.NativeWorkbook()
    raise ValueError(
        f"WOLFXL_WRITER={choice!r} is not a valid backend. "
        f"Choose one of {_VALID}. (W4C adds 'both' + 'auto'.)"
    )
