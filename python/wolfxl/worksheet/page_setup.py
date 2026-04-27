"""Print / page-setup classes for worksheets (RFC-055 §2.1 / §2.2).

These classes back ``ws.page_setup``, ``ws.page_margins``, and the
``PrintOptions`` accessor. They are openpyxl-shaped dataclasses with
``to_rust_dict()`` helpers that emit the §10 dict contract for the
PyO3 boundary.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


_VALID_ORIENTATION = ("default", "portrait", "landscape")
_VALID_CELL_COMMENTS = ("asDisplayed", "atEnd", "none")
_VALID_ERRORS = ("displayed", "blank", "dash", "NA")


@dataclass
class PageSetup:
    """Page setup (CT_PageSetup, ECMA-376 §18.3.1.51).

    Attributes correspond 1:1 to OOXML ``pageSetup`` element attributes.
    All attributes are optional — ``None`` means "let Excel default it".
    """

    orientation: str = "default"
    paperSize: int | None = None  # noqa: N815 - openpyxl public API
    fitToWidth: int | None = None  # noqa: N815
    fitToHeight: int | None = None  # noqa: N815
    scale: int | None = None
    firstPageNumber: int | None = None  # noqa: N815
    horizontalDpi: int | None = None  # noqa: N815
    verticalDpi: int | None = None  # noqa: N815
    cellComments: str | None = None  # noqa: N815
    errors: str | None = None
    useFirstPageNumber: bool | None = None  # noqa: N815
    usePrinterDefaults: bool | None = None  # noqa: N815
    blackAndWhite: bool | None = None  # noqa: N815
    draft: bool | None = None

    # openpyxl aliases (snake_case alternatives for the camelCase OOXML names)
    @property
    def paper_size(self) -> int | None:
        return self.paperSize

    @paper_size.setter
    def paper_size(self, value: int | None) -> None:
        self.paperSize = value

    @property
    def fit_to_width(self) -> int | None:
        return self.fitToWidth

    @fit_to_width.setter
    def fit_to_width(self, value: int | None) -> None:
        self.fitToWidth = value

    @property
    def fit_to_height(self) -> int | None:
        return self.fitToHeight

    @fit_to_height.setter
    def fit_to_height(self, value: int | None) -> None:
        self.fitToHeight = value

    def __post_init__(self) -> None:
        self._validate()

    def _validate(self) -> None:
        if self.orientation not in _VALID_ORIENTATION:
            raise ValueError(
                f"orientation must be one of {_VALID_ORIENTATION}, got {self.orientation!r}"
            )
        if self.cellComments is not None and self.cellComments not in _VALID_CELL_COMMENTS:
            raise ValueError(
                f"cellComments must be one of {_VALID_CELL_COMMENTS}, got {self.cellComments!r}"
            )
        if self.errors is not None and self.errors not in _VALID_ERRORS:
            raise ValueError(
                f"errors must be one of {_VALID_ERRORS}, got {self.errors!r}"
            )
        if self.scale is not None and not (10 <= self.scale <= 400):
            raise ValueError(f"scale must be between 10 and 400, got {self.scale}")

    def to_rust_dict(self) -> dict[str, Any]:
        """Emit the §10 ``page_setup`` dict for the PyO3 boundary."""
        return {
            "orientation": self.orientation if self.orientation != "default" else None,
            "paper_size": self.paperSize,
            "fit_to_width": self.fitToWidth,
            "fit_to_height": self.fitToHeight,
            "scale": self.scale,
            "first_page_number": self.firstPageNumber,
            "horizontal_dpi": self.horizontalDpi,
            "vertical_dpi": self.verticalDpi,
            "cell_comments": self.cellComments,
            "errors": self.errors,
            "use_first_page_number": self.useFirstPageNumber,
            "use_printer_defaults": self.usePrinterDefaults,
            "black_and_white": self.blackAndWhite,
            "draft": self.draft,
        }

    def is_default(self) -> bool:
        """True iff this PageSetup is at its construction defaults."""
        return self == PageSetup()


@dataclass
class PageMargins:
    """Page margins in inches (CT_PageMargins, ECMA-376 §18.3.1.49)."""

    left: float = 0.7
    right: float = 0.7
    top: float = 0.75
    bottom: float = 0.75
    header: float = 0.3
    footer: float = 0.3

    def to_rust_dict(self) -> dict[str, float]:
        return {
            "top": float(self.top),
            "bottom": float(self.bottom),
            "left": float(self.left),
            "right": float(self.right),
            "header": float(self.header),
            "footer": float(self.footer),
        }

    def is_default(self) -> bool:
        return self == PageMargins()


@dataclass
class PrintOptions:
    """`<printOptions>` toggles (CT_PrintOptions, ECMA-376 §18.3.1.70).

    Pod 2 re-exports this under ``wolfxl.worksheet.page.PrintOptions``.
    """

    horizontalCentered: bool = False  # noqa: N815
    verticalCentered: bool = False  # noqa: N815
    headings: bool = False
    gridLines: bool = False  # noqa: N815
    gridLinesSet: bool = True  # noqa: N815

    @property
    def horizontal_centered(self) -> bool:
        return self.horizontalCentered

    @horizontal_centered.setter
    def horizontal_centered(self, value: bool) -> None:
        self.horizontalCentered = bool(value)

    @property
    def vertical_centered(self) -> bool:
        return self.verticalCentered

    @vertical_centered.setter
    def vertical_centered(self, value: bool) -> None:
        self.verticalCentered = bool(value)


@dataclass
class PrintPageSetup:
    """Compatibility alias for openpyxl's ``PrintPageSetup``.

    openpyxl exposes this name in some module paths; the underlying
    object is the same as ``PageSetup``. Re-export for source
    compatibility — Pod 2 wires the import shim.
    """

    page_setup: PageSetup = field(default_factory=PageSetup)


__all__ = [
    "PageSetup",
    "PageMargins",
    "PrintOptions",
    "PrintPageSetup",
]
