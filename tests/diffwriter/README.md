# Differential Writer Harness

Runs oracle (`rust_xlsxwriter`) and native (`wolfxl-writer`) backends
against the same build plan and diffs the output across four layers.

## Layers

| Layer | What it checks | Gate |
|-------|---------------|------|
| 1 — byte | SHA-256 per-part after XML canonicalization + fuzzy strip | gold-star target (≥80%) |
| 2 — XML structural | `lxml.etree` tree diff after sorting where spec permits | ship gate |
| 3 — semantic | `tests/parity/_scoring.py` HARD/SOFT/INFO ratchet | ship gate |
| 4 — LibreOffice smoke | `soffice --headless --convert-to xlsx` round-trip | gold-star (≥95%), opt-in |

## Running

```bash
# All cases, current backend routing from modules.toml
uv run pytest tests/diffwriter/

# Force dual-backend comparison on every case
WOLFXL_WRITER=both uv run pytest tests/diffwriter/
```

## Layer 4 — LibreOffice smoke

Layer 4 is **opt-in**. Without `WOLFXL_RUN_LIBREOFFICE_SMOKE=1` the
`tests/diffwriter/soffice_smoke.py` tests skip cleanly, so the suite is
harmless on machines that don't have LibreOffice installed.

Install LibreOffice locally:

```bash
brew install --cask libreoffice    # macOS
apt-get install libreoffice        # Linux
```

Run:

```bash
WOLFXL_RUN_LIBREOFFICE_SMOKE=1 uv run pytest tests/diffwriter/soffice_smoke.py -v
```

Acceptance gate is ≥95% pass rate across the 25 hand-built cases + 15
SynthGL fixtures (40 round-trips total). A handful of expected failures
goes in `_SOFFICE_XFAIL_CASES` in `cases/__init__.py` so the layer
remains green-by-default while documenting known LO incompatibilities.

Layer 4 is gold-star: a failing case indicates an interop bug worth
chasing in a follow-up slice but does NOT block ship.

## Module status

See `modules.toml` — the source of truth for which modules route to
native vs oracle, and which cases pass on each layer.

## Directory layout

```
tests/diffwriter/
  README.md               # this file
  __init__.py
  modules.toml            # per-module per-case status (source of truth)
  fuzzy_elements.json     # Layer-1 allowlist (timestamps etc.)
  order_rules.py          # Layer-2 child-sort rules
  cases/                  # hand-written test cases (one file per module)
  replay/                 # auto-generated from existing parity fixtures
  reports/                # generated HTML diff reports (gitignored)
```

## Wave status

- **Wave 0**: harness skeleton ✔
- **Wave 1**: foundations (model, refs, zip+container) — in progress
- **Wave 2**: emitters (styles, sheet, SST+workbook)
- **Wave 3**: rich features (comments+VML, tables, CF+DV)
- **Wave 4**: NativeWorkbook pyclass + DualWorkbook + harness wiring
- **Wave 5**: rip-out rust_xlsxwriter
