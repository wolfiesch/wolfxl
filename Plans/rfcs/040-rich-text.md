# RFC-040: Rich-text reads (`Cell.rich_text` + `CellRichText`)

Status: Shipped <!-- TBD: flip to Shipped + add Pod-α merge SHA when Sprint Ι Pod-α lands -->
Owner: Sprint Ι Pod-α
Phase: Read-side parity (1.3)
Estimate: M
Depends-on: RFC-013 (patcher infra, indirectly via reader), RFC-021 (no — defined names is independent)
Unblocks: T2 rich-text writes (post-1.3), full SynthGL ingest fidelity for templates that carry inline run-formatted strings.

> **S** = ≤2 days; **M** = 3-5 days; **L** = 1-2 weeks; **XL** = 2+ weeks.

## 1. Problem Statement

KNOWN_GAPS.md Phase 3:

> `Cell.value` when backing is `CellRichText` — currently wolfxl
> flattens rich text to plain. Add `Cell.rich_text` property
> (iter-compatible with `openpyxl.cell.rich_text.CellRichText`).

User code:

```python
wb = wolfxl.load_workbook("template.xlsx")
ws = wb["Notes"]
cell = ws["B7"]                     # has bold "WARNING: " + plain "do not delete"
print(cell.value)                   # → "WARNING: do not delete"  (flat string)
print(cell.rich_text)               # AttributeError today
```

openpyxl returns a `CellRichText` instance — a list-like of
`TextBlock(font: InlineFont, text: str)` plus bare `str` runs — that
preserves per-run formatting. SynthGL (and any downstream consumer
that re-emits rich text without losing the run boundaries) needs the
same shape.

**Target behaviour**: when an `<is>` block in the cell or the matching
`<si>` entry in `xl/sharedStrings.xml` carries `<r><rPr>...</rPr>...</r>`
runs, expose them as `Cell.rich_text` (a `CellRichText`-shaped iterable).
`Cell.value` keeps its current "flatten to plain str" contract for
backwards compatibility — but only when there's no rich text;
breaking-change details in §7.

## 2. OOXML Spec Surface

| Spec section | Element | Notes |
|---|---|---|
| §18.4.8 CT_Rst | `<is>` (inline string) | Rich-text-bearing cells use `<c t="inlineStr">` and embed `<is><r><rPr>…</rPr><t>…</t></r></is>`. |
| §18.4.8 CT_Rst | `<si>` (shared string) | Same shape, but indexed by `<c t="s"><v>N</v></c>`. RFC-040 reads both encodings via the same parser. |
| §18.4.7 CT_RElt | `<r>` | One run; child `<rPr>` plus `<t>`. Concatenation of all `<t>` is the flat string `Cell.value` returns today. |
| §18.4.5 CT_RPrElt | `<rPr>` | Per-run formatting — `<b/>`, `<i/>`, `<sz val="…"/>`, `<color rgb="…"/>`, `<rFont val="…"/>`, etc. Mirrors `Font` but **inline** (no styles.xml dxf indirection). |
| §18.4.12 CT_PhoneticRun | `<rPh>` | East-Asian phonetic-guide runs. **Out of scope** (§10). |

Schema-ordering: `<r>` may interleave with `<rPh>` and bare `<t>`
fragments inside an `<si>`/`<is>`; the parser MUST tolerate any
mix and concatenate `<t>` text in document order.

## 3. openpyxl Reference

`openpyxl/cell/rich_text.py:1-150`:

* `InlineFont` — frozen dataclass mirroring `<rPr>` attrs.
* `TextBlock(font: InlineFont, text: str)` — one styled run.
* `CellRichText(list)` — iterable of `TextBlock | str`. Bare strings
  are unstyled (no `<rPr>`).
* `Cell.value` returns either a `str` OR a `CellRichText` depending on
  whether the cell has any styled runs. Plain-`<t>`-only cells still
  return `str`.

Edge cases openpyxl handles:

* Empty rich text (`<is><t/></is>` → `""`) — returns `str`, not `CellRichText`.
* Rich text with only one `<r>` and no `<rPr>` — flattens to `str`.
* `xml:space="preserve"` on `<t>` — leading/trailing whitespace honored.

We **do not** copy openpyxl's rich-text WRITE path (that's a future
RFC; see §10).

## 4. WolfXL Surface Area

### 4.1 Python coordinator

`python/wolfxl/cell/rich_text.py` (new file):

```python
@dataclass(frozen=True)
class InlineFont:
    name: str | None = None
    size: float | None = None
    bold: bool = False
    italic: bool = False
    underline: str | None = None
    strike: bool = False
    color: str | None = None  # ARGB hex
    family: int | None = None
    scheme: str | None = None

@dataclass(frozen=True)
class TextBlock:
    font: InlineFont
    text: str

class CellRichText(list):
    """``list`` subclass of ``TextBlock | str``. Stringification
    concatenates all run text in document order so ``str(rt) == cell.value_str``."""
    def __str__(self) -> str:
        return "".join(b.text if isinstance(b, TextBlock) else b for b in self)
```

`python/wolfxl/_cell.py`:

* New `Cell.rich_text` `@property` — returns `CellRichText | None`.
  `None` when the cell has no rich text (matches openpyxl's "either
  str or CellRichText" contract by routing the rich-text branch only
  when runs are present).
* `Cell.value` keeps its current flatten-to-str contract for inline
  strings WITHOUT runs. Cells WITH runs change behaviour — see §7.

### 4.2 Reader (calamine-styled)

`src/calamine_styled_backend.rs` already parses `<is>`/`<si>` for
flat text. RFC-040 adds a parallel parse path that produces a
`Vec<RichRun>` instead of a flat `String`, where:

```rust
struct RichRun {
    font: Option<InlineFontPayload>,  // None == bare <t>
    text: String,
}
struct InlineFontPayload { /* same shape as Python InlineFont */ }
```

The PyO3 boundary exposes `read_rich_text(sheet: &str, coord: &str)
-> Option<Vec<(Option<InlineFontPayload>, String)>>`. `None` means
"no rich text on this cell" (callers fall back to `Cell.value`).

### 4.3 Native writer / patcher

Out of scope for RFC-040. Both write paths continue to round-trip
rich text bytes when modifying via `XlsxPatcher` (the cell is
preserved verbatim) but offer no Python API to construct or mutate
rich text. Tracked as a future RFC; see §10.

## 5. Implementation Sketch

1. **Reader path**. Extend `calamine_styled_backend.rs` to walk the
   `<r>` children inside an `<is>` (or the `<si>` row of
   sharedStrings) and emit `RichRun` per child. `<rPr>` parsing
   reuses the existing inline-style helpers — no new dependencies.

2. **PyO3 surface**. Add `read_rich_text` to `CalamineStyledBook`. Returns
   `Option<Vec<…>>`; `None` is the no-runs case.

3. **Python wrapper**. `Cell.rich_text` calls into the Rust API once
   per access (no caching needed — modify-mode rich-text writes are
   future work, so the return is morally read-only).

4. **Empty-runs invariant**. A cell whose `<is>` has only one `<t>`
   without `<rPr>` returns `None` from `read_rich_text` — no spurious
   `CellRichText([str])` allocation.

5. **`str(CellRichText)` parity**. `CellRichText.__str__` matches
   `Cell.value` exactly for cells where both surfaces resolve, so
   `str(cell.rich_text) == cell.value` is an invariant the test
   harness pins.

## 6. Verification Matrix

1. **Rust unit tests** — new fixtures in
   `crates/wolfxl-reader/tests/rich_text.rs` covering `<is>`-encoded,
   `<si>`-encoded, single-run-no-rPr, multi-run, and phonetic-run
   ignored cases.
2. **Python round-trip** — `tests/parity/test_rich_text_parity.py`
   reads fixtures authored by openpyxl; asserts `cell.rich_text`
   matches the source `CellRichText` element-by-element.
3. **openpyxl parity** — `tests/parity/test_rich_text_parity.py`
   compares `Cell.value` (flat) and `Cell.rich_text` (list) against
   openpyxl on the same fixture.
4. **Cross-mode** — modify-mode reads rich text the same way; tested
   on a modify-mode loaded workbook.
5. **Regression fixtures** — generated in-test by
   `tests/parity/test_rich_text_parity.py` and
   `tests/test_rich_text_read.py`.
6. **LibreOffice cross-renderer** — N/A (read-only RFC).

## 7. Cross-Mode Asymmetries (BREAKING)

`Cell.value` semantics on rich-text-bearing cells:

* **Pre-1.3** (current): returns flat `str` — runs are silently
  flattened.
* **Post-1.3** (Pod-α): returns flat `str` (unchanged) for
  callers who want the existing behaviour, BUT the official path
  for "I want the runs" is `Cell.rich_text`. The release notes
  call this out — see `docs/release-notes-1.3.md`.

A new `Cell.value_str` accessor MAY be added to give callers a
guaranteed-flat-`str` API independent of how `Cell.value`
evolves. Decision deferred — see §11 OQ-1.

## 8. Risks

| # | Risk | Likelihood | Impact | Mitigation |
|---|------|------------|--------|------------|
| 1 | Calamine's `<rPr>` parser drops attrs we care about | low | medium | Add a fallback "raw bytes" path; expose unknown-attr passthrough on `InlineFont` |
| 2 | Memory bloat on large rich-text-heavy sheets | medium | low | Lazy-decode: `Cell.rich_text` parses on access, not on workbook open |
| 3 | `str(CellRichText)` diverges from `Cell.value` after a future write-side change | medium | medium | Pin both via `tests/parity/test_rich_text_parity.py`; add a no-divergence assert in CI |

## 9. Effort Breakdown

| Slice | Estimate | Notes |
|---|---|---|
| Reader (Rust) | 1.5 days | Extend `<is>`/`<si>` parser |
| PyO3 surface | 0.5 day | `read_rich_text` |
| Python wrapper + dataclasses | 0.5 day | `cell/rich_text.py` |
| Tests | 1 day | Round-trip + parity + fixture |
| Docs + release notes | 0.5 day | Migration guide |
| **Total** | **~4 days** | M-bucket |

## 10. Out of Scope

* **Writing** rich text (constructing `CellRichText` and saving). Requires
  matching the patcher's modify-mode flush AND the native writer's
  `<sheetData>` emit path. Tracked as a future RFC (post-1.3).
* `<rPh>` phonetic runs — drop on read (matches openpyxl).
* Rich text in `<headerFooter>` and `<comment>` text bodies — separate
  paths; not part of RFC-040.

## 11. Open Questions

| OQ | Question | Status |
|---|---|---|
| OQ-1 | Add `Cell.value_str` for callers who explicitly want flattened text? | Pending — Pod-α to call |
| OQ-2 | Should `cell.value = "plain"` on a rich-text cell clear `cell.rich_text`? Modify-mode only. | Defer — write path is out of scope |

## Acceptance

(Filled in after Pod-α merges.)

- Commit: <!-- TBD: Pod-α commit sha when integrated -->
- Verification: `python scripts/verify_rfc.py --rfc 040` GREEN at <!-- TBD -->
- Date: <!-- TBD -->
