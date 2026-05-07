"""XML backend flags compatible with ``openpyxl.xml``."""

from __future__ import annotations

import os


def _lxml_available() -> bool:
    try:
        from lxml.etree import LXML_VERSION
    except ImportError:
        return False
    return LXML_VERSION >= (3, 3, 1, 0)


def _defusedxml_available() -> bool:
    try:
        import defusedxml  # noqa: F401
    except ImportError:
        return False
    return True


LXML = _lxml_available() and os.environ.get("OPENPYXL_LXML", "True") == "True"
DEFUSEDXML = (
    _defusedxml_available()
    and os.environ.get("OPENPYXL_DEFUSEDXML", "True") == "True"
)

__all__ = ["DEFUSEDXML", "LXML"]
