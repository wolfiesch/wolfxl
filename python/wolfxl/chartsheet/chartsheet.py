"""Chartsheet containers compatible with ``openpyxl.chartsheet``."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass
class ChartsheetProperties:
    """Small openpyxl-shaped chartsheet properties container."""

    published: bool | None = None
    codeName: str | None = None  # noqa: N815 - openpyxl camelCase


@dataclass
class ChartsheetProtection:
    """Small openpyxl-shaped chartsheet protection container."""

    content: bool | None = None
    objects: bool | None = None


class Chartsheet:
    """A workbook tab that contains a single full-sheet chart."""

    def __init__(self, parent: Any, title: str) -> None:
        self._parent = parent
        self.title = title
        self.sheet_state = "visible"
        self.sheet_properties = ChartsheetProperties()
        self.protection = ChartsheetProtection()
        self._charts: list[Any] = []
        self._source_chartsheet = False

    def add_chart(self, chart: Any) -> None:
        """Attach ``chart`` to this chartsheet."""
        if chart is None or not hasattr(chart, "to_rust_dict"):
            raise TypeError(
                f"Chartsheet.add_chart expects a wolfxl chart object, got {type(chart).__name__}"
            )
        self._charts[:] = [chart]

    @property
    def _drawing(self) -> Any | None:
        """openpyxl-shaped private drawing placeholder."""
        return self._charts[0] if self._charts else None
