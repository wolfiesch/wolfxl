"""Shared infrastructure for openpyxl-API shims.

Many openpyxl modules expose classes wolfxl doesn't implement yet — charts,
images, pivots, data validations, conditional formatting rules, named styles.
Instead of silently missing the module (``ModuleNotFoundError``, which masks
the real story from users migrating from openpyxl), we expose the module
paths with stub classes that raise a clear ``NotImplementedError`` on
instantiation. The error message tells the user what's missing and points
at modify mode (which preserves most of these features on round-trip without
needing a Python-side class).
"""

from __future__ import annotations

from typing import Any


class _UnsupportedFeature:
    """Base for openpyxl classes that wolfxl exposes as shims.

    Subclasses carry a human-readable ``_feature_name`` and a ``_hint``
    pointing to the recommended path: modify-mode preservation, native
    WolfXL support when available, or an openpyxl fallback. Attempting to
    instantiate raises with the full message so users discover the gap at
    the construction site rather than at a downstream ``AttributeError``.
    """

    _feature_name: str = "<unknown>"
    _hint: str = ""

    def __init__(self, *args: Any, **kwargs: Any) -> None:  # noqa: ARG002
        raise NotImplementedError(
            f"wolfxl does not implement {self._feature_name}. {self._hint} "
            "See https://github.com/SynthGL/wolfxl#openpyxl-compatibility "
            "for compatibility notes."
        )


def _make_stub(name: str, hint: str) -> type:
    """Create a named subclass of ``_UnsupportedFeature`` for shim modules."""
    return type(name, (_UnsupportedFeature,), {"_feature_name": name, "_hint": hint})
