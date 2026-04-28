# wolfxl 1.7.0 — public-launch slice (no pivot tables)

_Date: 2026-04-27_

WolfXL 1.7.0 is the **openpyxl-replacement launch slice**. After 23
RFCs across 8 sprints (Δ → Ξ) the construction-side surface is
exhaustively shipped, with one explicit exception: pivot tables are
preserved on round-trip but not yet constructible (that's v2.0.0 /
Sprint Ν).

## TL;DR

- ✅ **Production-ready.** PyPI classifier promoted from `4 - Beta`
  to `5 - Production/Stable`.
- ✅ **`pyproject.toml` + `Cargo.toml` versions sync to `1.7.0`.**
  The package version had drifted from the git tag since v0.5;
  `wolfxl.__version__` now reports `1.7.0` correctly.
- ✅ **`Worksheet.remove_chart(chart)` + `Worksheet.replace_chart(old, new)`.**
  Closes the v1.6.1 release-notes deferral.
- ✅ **`chart.title = RichText(...)` works.** The v1.6.1 known-gap
  `xfail` flips to passing.
- ✅ **`docs/migration/` rewritten for v1.7.** Compatibility Matrix
  + openpyxl-migration walkthroughs cover all 16 chart families,
  images, encryption, streaming reads, structural ops, modify-mode
  mutations, rich text.
- ✅ **`docs/performance/` refreshed.** Read / write / modify-mode /
  chart-construction speedup tables on a 1k / 10k / 100k row
  matrix.
- ✅ **`Plans/launch-posts.md` materialised** — HN, Twitter/X,
  r/Python, dev.to, GitHub Discussions drafts.

## What's new

### `Worksheet.remove_chart(chart)`

Mirrors the openpyxl idiom `ws._charts.remove(chart)`:

```python
import wolfxl
from wolfxl.chart import BarChart, Reference

wb = wolfxl.Workbook()
ws = wb.active
ws.append(["Region", "Q1", "Q2"])
ws.append(["NA", 100, 110])

chart = BarChart()
data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=2)
chart.add_data(data, titles_from_data=True)
ws.add_chart(chart, "E2")

# Decide we don't want it after all:
ws.remove_chart(chart)
wb.save("no_chart.xlsx")
```

**Scope (v1.7)**: removes a chart that was added via `add_chart`
and has not yet been flushed to disk. Removal of charts that
survive from the source workbook (modify mode) is a v1.8 follow-up
(needs the patcher to expose a `queue_chart_remove` alongside the
existing `queue_chart_add`). If you call `remove_chart` on a chart
that wasn't added via `add_chart`, you get a clear `ValueError`
pointing at the v1.8 follow-up.

### `Worksheet.replace_chart(old, new)`

Convenience for "swap the chart at this anchor":

```python
old_chart = BarChart()
ws.add_chart(old_chart, "G2")

# ... later, with a redesigned chart ...
new_chart = LineChart()
ws.replace_chart(old_chart, new_chart)
# new_chart inherits old_chart's anchor ("G2")
# unless new_chart._anchor was set explicitly first.
```

The replacement chart inherits the old chart's anchor (so
deterministic chart-ID allocation matches the pre-replace layout)
unless the caller already set `new._anchor` explicitly. Wrong
type for `new` raises `TypeError`; missing `old` raises
`ValueError`.

### `chart.title = RichText(...)`

In v1.6.1 the only `xfail` left in `tests/test_charts_write.py`
was `test_line_chart_title_rich_text`: `chart.title=` rejected
openpyxl-style `RichText(bodyPr=..., p=[...])` with a `TypeError`.
Sprint Ξ closes the gap:

```python
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    CharacterProperties,
    Paragraph,
    ParagraphProperties,
    Run as _OpenpyxlRun,    # 3.2.x; or RegularTextRun on 3.1.x
)

bold_run = _OpenpyxlRun(t="Bold", rPr=CharacterProperties(b=True))
red_run  = _OpenpyxlRun(
    t=" Red",
    rPr=CharacterProperties(solidFill=ColorChoice(srgbClr="FF0000")),
)
para = Paragraph(pPr=ParagraphProperties(), r=[bold_run, red_run])

chart = LineChart()
chart.title = RichText(p=[para])     # accepted as of v1.7
```

`TitleDescriptor.__set__` now accepts wolfxl-typed `RichText`,
openpyxl-typed `RichText` via duck typing, plus the existing `str`
and `Title` paths. `Title.to_dict()` additionally coerces
openpyxl's `ColorChoice`-typed `solidFill` into the hex string the
Rust emitter expects.

### Version sync

`pyproject.toml` and `Cargo.toml` are bumped to `1.7.0`. The Rust
side is the canonical source — `src/lib.rs:41` re-exports
`env!("CARGO_PKG_VERSION")` as `wolfxl.__version__`, so the bump
is a one-line edit in each `Cargo.toml` package and pyproject. From
v1.7 forward, the version stays tracked through both files.

### Production / Stable classifier

The PyPI classifier ships as `Development Status :: 5 - Production/Stable`.
This isn't just a vanity move:

- The **read-side surface** has been frozen since v1.4 (Sprint Κ —
  `.xls`/`.xlsb` reads) with no deprecations since.
- The **write-side surface** has been frozen since v1.6 (Sprint Μ —
  chart construction) with no deprecations since.
- The **modify-mode surface** has been frozen since v1.1 (Sprint Ζ —
  RFC-035 `copy_worksheet`).
- Every parity claim is pinned by
  `tests/parity/openpyxl_surface.py` + the strict-xfail ratchet.

The v1.7 release ships with **zero strict-xfail markers** outside
of optional-dependency-not-installed cases.

### Documentation refresh

- **`docs/migration/openpyxl-migration.md`** — full rewrite. The
  pre-v1.7 file was 38 lines covering read-side basics; the v1.7
  rewrite covers all 16 chart families, images, encryption (read
  + write), streaming reads, `.xlsb`/`.xls` reads, structural ops,
  modify-mode mutations, and rich text. Includes a "what to
  validate during migration" + "edge cases worth knowing" +
  "when to keep openpyxl alongside" section.
- **`docs/migration/compatibility-matrix.md`** — full rewrite.
  Exhaustive 50+ row table of every openpyxl symbol with v1.7
  status, plus ecosystem comparison vs openpyxl 3.1.5 / XlsxWriter
  / pandas / fastexcel / python-calamine / FastXLSX /
  rustpy-xlsxwriter.
- **`docs/performance/benchmark-results.md`** — v1.7 numbers on a
  1k / 10k / 100k row matrix vs openpyxl 3.1.5 (read / write /
  modify-mode / chart-construction). Headline:
  - Read 100k rows: **11×** speedup
  - Write 100k rows: **10×** speedup
  - Touch 1 cell + save (100k-row workbook): **124×** speedup
  - copy_worksheet (10k-row + table + DV + CF): **18×** speedup
- **`docs/performance/methodology.md`** — extended with
  construction-side benchmark guidance (chart cache rebuild
  semantics, image media reuse on copy_worksheet).
- **`docs/performance/run-on-your-files.md`** — extended with a
  chart-construction microbenchmark + copy_worksheet harness.

### Launch-post drafts

`Plans/launch-posts.md` is materialised with drafts for:

| Channel | Tone | Length |
|---|---|---|
| HN ("Show HN") | Technical | Medium |
| Twitter / X | Direct, hooky | 8-tweet thread |
| Reddit r/Python | Detailed, comparative | Medium-long |
| dev.to | Long-form narrative | Long (outline) |
| GitHub Discussions | Announcement, supportive | Medium |

Plus a pre-launch + post-launch monitoring checklist.

## RFC-046 §13 — legacy chart-dict key sunset

Sprint Μ-prime's integrator finalize taught us that the
`parse_graphical_properties` Rust parser carries a
backwards-compatibility shim that accepts BOTH the §10 form
(`solid_fill` + nested `ln: {solid_fill, w_emu, prst_dash}`) AND
the legacy form (`fill_color`, `line_color`, `line_dash`,
`line_width_emu`).

The Python emitter has used the §10 form exclusively since v1.6.1.
The legacy form survives only as a Rust-side accept-also for any
out-of-tree caller that builds a chart dict by hand and passes it
to the (private) `wolfxl._rust.serialize_chart_dict` helper.

**Sunset timeline**:

| Version | Status |
|---|---|
| v1.7.0 | Both forms accepted; documentation-only deprecation. |
| v2.0.0 | Removal — only §10.9 form supported. |

No runtime warning is emitted (avoids spurious noise for any
caller outside our test suite). v2.0.0 will remove the shim in
the same commit that adds the pivot-table chart parser path.

## Out of scope (deferred)

- **Pivot table construction** — Sprint Ν / v2.0.0. Pivot tables
  are preserved on modify-mode round-trip; they're just not
  constructible from scratch in Python yet.
- **Pivot-chart linkage** (`<c:pivotSource>` referencing a pivot
  cache definition) — Sprint Ν / v2.0.0.
- **Combination charts** (multi-plot — bar + line on shared
  axes) — not on the v1.7 roadmap; v1.7.x if requested.
- **`Worksheet.delete_chart_persisted`** (modify-mode removal of
  charts that survive from the source workbook) — v1.8 follow-up.
- **OpenDocument (`.ods`)** — out of scope; not on the roadmap.

## Acknowledgements

To everyone who file-bugged the v1.6 chart contract gap, the
streaming-datetime drift, the RFC-035 cross-RFC composition bugs,
and the 100 little things that surface at scale.

## Stats (post-1.7.0)

- `python -c "import wolfxl; print(wolfxl.__version__)"` → `1.7.0`
- `pytest tests/test_charts_write.py tests/test_charts_remove.py
  tests/test_charts_3d.py` → **82 passed**, 1 skipped (vs v1.6.1's
  75 passed, 1 skipped, 1 xfailed)
- `tests/test_charts_remove.py` adds 7 new tests
- `tests/test_charts_write.py::test_line_chart_title_rich_text`
  flips xfail → pass

## Pods that landed 1.7.0

- **Pod-α** — `pyproject.toml` / `Cargo.toml` version bump,
  `Worksheet.remove_chart` / `replace_chart` (`tests/test_charts_remove.py`),
  RichText title support (`python/wolfxl/chart/title.py`), RFC-046
  §13 / §14 / §15 documentation.
- **Pod-β** — `docs/migration/openpyxl-migration.md` +
  `docs/migration/compatibility-matrix.md` rewrites.
- **Pod-γ** — `docs/performance/` refresh.
- **Pod-δ** — `Plans/launch-posts.md` drafts.

Sprint Ξ ran without the worktree-fanout pattern (lessons #8-12
from prior sprints made the integrator-inline path safer for
small-LOC, mostly-docs slices). All four pods authored on
`feat/native-writer` directly.

## RFCs

- `Plans/rfcs/046-chart-construction.md` §13 / §14 / §15 — Sprint
  Ξ additions.
- `Plans/sprint-xi.md` — Sprint Ξ tracking doc.

## Next

**Sprint Ν / v2.0.0** — pivot tables + pivot caches + pivot charts.
This is the last construction-side milestone on the openpyxl-parity
roadmap. After it ships, the marketing claim becomes "full openpyxl
replacement, period."
