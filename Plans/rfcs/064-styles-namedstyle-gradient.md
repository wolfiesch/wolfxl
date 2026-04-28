# RFC-064 — Styles: NamedStyle + GradientFill + Protection + DifferentialStyle + Fill (Sprint Π Pod Π-γ)

> Closes 5 stubs: `NamedStyle`, `Protection`, `GradientFill`,
> `DifferentialStyle`, `Fill`.
>
> **Highest-risk pod in Sprint Π** — touches `xl/styles.xml` emit,
> dxf table allocation, and Cell.style lookup.

## 1. Goal

Land Python construction + Rust emit + save() integration for the 5
style-table classes. Most-impactful is `NamedStyle` (registry on
Workbook + cellStyleXfs emit + `Cell.style = "Heading 1"` lookup).

## 2. Python class layer

### 2.1 `python/wolfxl/styles/_named_style.py` (NEW canonical impl)

```python
from dataclasses import dataclass, field
from typing import Any


@dataclass
class NamedStyle:
    """Named style (CT_CellStyle §18.8.7).

    A NamedStyle bundles Font + Fill + Border + Alignment +
    Protection + number_format under a single name. Workbook
    registry: ``wb.named_styles``.

    Cell assignment: ``cell.style = "Heading 1"``.
    """
    name: str = ""
    font: Any = None
    fill: Any = None
    border: Any = None
    alignment: Any = None
    protection: Any = None
    number_format: str = "General"
    builtinId: int | None = None  # noqa: N815  # 0 = Normal, 1 = Heading 1, ...
    customBuiltin: bool = False    # noqa: N815
    hidden: bool = False
    xfId: int | None = None         # noqa: N815  # set at registration time

    def to_rust_dict(self) -> dict[str, Any]: ...

    @property
    def is_builtin(self) -> bool:
        return self.builtinId is not None and not self.customBuiltin


@dataclass
class _NamedStyleList:
    """Registry on Workbook. ``wb.named_styles`` returns this."""
    _styles: list[NamedStyle] = field(default_factory=list)
    _by_name: dict[str, NamedStyle] = field(default_factory=dict)

    def append(self, ns: NamedStyle) -> None: ...
    def __getitem__(self, name: str) -> NamedStyle: ...
    def __contains__(self, name: str) -> bool: ...
    def __iter__(self): ...
    def names(self) -> list[str]: ...
```

Re-export:
- `python/wolfxl/styles/__init__.py` flips `NamedStyle = _make_stub(...)` to `from wolfxl.styles._named_style import NamedStyle`
- `python/wolfxl/styles/named_styles.py` re-exports `NamedStyle` (already exists, currently stub)

### 2.2 `python/wolfxl/styles/_protection.py` (NEW canonical impl)

```python
@dataclass
class Protection:
    """Cell protection (CT_CellProtection §18.8.33).

    Backs ``cell.protection``. Only meaningful when the sheet
    itself is protected (RFC-055 SheetProtection)."""
    locked: bool = True
    hidden: bool = False

    def to_rust_dict(self) -> dict[str, bool]:
        return {"locked": self.locked, "hidden": self.hidden}
```

`Cell.protection` property added. `_format_dirty` flushes through
existing dxf path.

### 2.3 `python/wolfxl/styles/_gradient.py` (NEW canonical impl)

```python
@dataclass
class Stop:
    """Single gradient color stop (CT_GradientStop §18.8.24)."""
    position: float = 0.0     # 0.0 - 1.0
    color: str = "FFFFFFFF"   # ARGB hex


@dataclass
class GradientFill:
    """Gradient fill (CT_GradientFill §18.8.24).

    Two flavours via ``type``:
    - ``"linear"``: degree controls angle
    - ``"path"``: left/right/top/bottom rectangles control radial focus
    """
    type: str = "linear"
    degree: float = 0.0
    left: float = 0.0
    right: float = 0.0
    top: float = 0.0
    bottom: float = 0.0
    stop: list[Stop] = field(default_factory=list)

    def to_rust_dict(self) -> dict[str, Any]: ...
```

Wires into the existing fills table. `Cell.fill = GradientFill(...)`
must work alongside `PatternFill`.

### 2.4 `python/wolfxl/styles/differential.py` (replaces stub)

```python
@dataclass
class DifferentialStyle:
    """A dxf entry (CT_Dxf §18.8.14).

    Already partially used by Sprint-Ο Pod-3 conditional formatting
    code via the ``Format`` model. Pod Π-γ promotes ``Format`` to a
    public alias of ``DifferentialStyle`` for openpyxl compatibility.
    """
    font: Any = None
    fill: Any = None
    border: Any = None
    alignment: Any = None
    number_format: str | None = None
    protection: Any = None

    def to_rust_dict(self) -> dict[str, Any]: ...
```

### 2.5 `python/wolfxl/styles/fills.py` (replaces stub)

```python
class Fill:
    """Abstract base for ``PatternFill`` / ``GradientFill``.

    Exists so user code that does ``isinstance(f, Fill)`` works for
    both subclasses. Wolfxl's existing PatternFill / GradientFill
    are made to subclass this at import time via __init_subclass__.
    """
```

## 3. Workbook integration

```python
class Workbook:
    @property
    def named_styles(self) -> _NamedStyleList:
        if self._named_styles is None:
            self._named_styles = _NamedStyleList()
            # Seed the 47 builtin named styles (Normal, Heading 1, etc.)
            self._seed_builtin_styles()
        return self._named_styles

    def _flush_pending_named_styles_to_patcher(self) -> None:
        """Drain user-defined NamedStyles into the cellStyleXfs table.
        Sequenced AFTER pivots / sheet-setup, BEFORE Phase 3 (cells)
        so cell.style="..." lookups resolve to allocated xfIds."""
```

New `__slots__`: `_named_styles`.

## 4. Cell.style lookup

```python
class Cell:
    @property
    def style(self) -> str | None:
        """Name of the NamedStyle applied to this cell, or None."""
        return self._named_style_name

    @style.setter
    def style(self, name: str | None) -> None:
        if name is not None:
            ns = self._ws._workbook.named_styles[name]
            # Apply font/fill/border/alignment/protection/number_format
            self._font = ns.font
            self._fill = ns.fill
            self._border = ns.border
            self._alignment = ns.alignment
            self._protection = ns.protection
            self._number_format = ns.number_format
            self._format_dirty = True
        self._named_style_name = name
```

New `__slots__` on Cell: `_named_style_name`, `_protection`.

## 5. Rust emit changes

`crates/wolfxl-writer/src/emit/styles.rs`:

- Extend `<fills>` table to handle GradientFill entries (XML differs from PatternFill)
- Add `<cellStyleXfs>` table emit (currently just `<cellXfs>`)
- Add `<cellStyles>` table emit (NamedStyle name → xfId mapping)
- Add `<dxfs>` extension for DifferentialStyle entries (coordinate with Pod-3 styling — RFC-061 §10.7)

`crates/wolfxl-writer/src/parse/named_styles.rs` (NEW): parses
`<cellStyleXfs>` + `<cellStyles>` from existing workbook for modify-mode preservation.

## 6. PyO3 bindings

`src/wolfxl/named_styles.rs` (NEW):
- `serialize_named_style_dict(py, dict) -> bytes`
- `serialize_gradient_fill_dict(py, dict) -> bytes`
- `serialize_protection_dict(py, dict) -> bytes`
- `serialize_dxf_dict(py, dict) -> bytes` (or extend Pod-3's dxf serialization)

## 7. Patcher Phase 2.5s

NEW phase between sheet-setup (2.5n) and slicers (2.5p) for named-style
drainage. (Page breaks at 2.5r, named styles at 2.5s.)

Final order: `2.5m → 2.5n → 2.5r → 2.5s → 2.5p → 2.5o → 2.5q → 2.5h`.

## 8. §10 contracts

```python
NamedStyle.to_rust_dict() = {
    "name": "Heading 1",
    "font": {...},          # Font dict (matches existing font_to_format_dict)
    "fill": {...},
    "border": {...},
    "alignment": {...},
    "protection": {"locked": True, "hidden": False},
    "number_format": "General",
    "builtin_id": 1,
    "custom_builtin": False,
    "hidden": False,
    "xf_id": null,           # allocated by Rust at flush time
}

GradientFill.to_rust_dict() = {
    "type": "linear",
    "degree": 90.0,
    "left": 0.0, "right": 0.0, "top": 0.0, "bottom": 0.0,
    "stops": [
        {"position": 0.0, "color": "FFFF0000"},
        {"position": 1.0, "color": "FF0000FF"},
    ],
}

DifferentialStyle.to_rust_dict() = {
    "font": {...} | None,
    "fill": {...} | None,
    "border": {...} | None,
    "alignment": {...} | None,
    "number_format": "0.00" | None,
    "protection": {"locked": True, "hidden": False} | None,
}
```

## 9. Tests

- `tests/test_named_style.py` (~50) — construction, registry, builtin lookup, `cell.style = "..."`
- `tests/test_gradient_fill.py` (~25) — linear + path variants
- `tests/test_cell_protection.py` (~15)
- `tests/test_differential_style.py` (~15)
- `tests/test_fill_isinstance.py` (~5) — `isinstance(PatternFill(...), Fill)` parity
- `tests/diffwriter/test_named_style_bytes.py` (`WOLFXL_TEST_EPOCH=0`)
- `tests/parity/test_named_style_parity.py` — round-trip via openpyxl
- `tests/parity/test_gradient_fill_parity.py`

## 10. Tolerable pre-existing failures

None.

## 11. Coordination with Sprint-Ο Pod-3

Pod 3's `Format` model in conditional-formatting / pivot styling
(`bff240f` / `831ea5e`) overlaps with `DifferentialStyle`. RFC-064
§2.4 makes `DifferentialStyle` the canonical name; `Format` becomes
an alias:

```python
# python/wolfxl/styles/differential.py
DifferentialStyle = ...  # canonical
Format = DifferentialStyle  # alias for Pod-3 compat
```

Pod-3's existing dxf-id allocation in patcher Phase 2.5p.5 stays put —
Pod Π-γ's NamedStyle drainage runs in Phase 2.5s and uses a separate
xfId allocator.
