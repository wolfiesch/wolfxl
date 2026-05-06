# WolfXL 2.0 RC Validation Log

Status: pre-release hardening. Do not tag, publish, or cut 2.0 from this
checkpoint.

## Green Gates

| Gate | Result |
|---|---:|
| `uv run maturin develop` | Pass |
| Full Python suite with optional deps | 2720 passed / 20 skipped |
| Openpyxl compat oracle | 80 / 80 |
| XLS/XLSB optional slice with `python-calamine` | Pass |
| External oracle fixture pack | Pass |
| `cargo test --workspace -q` | Pass |
| `cargo fmt --check` | Pass |
| `cargo check -q` | Pass |
| `ruff check` focused on touched Python files | Pass |

## Upstream Openpyxl Corpus

The upstream openpyxl 3.1.5 test slice now collects cleanly under the
`openpyxl -> wolfxl` shim and passes 65 / 67 tests.

The remaining failures are strict package-shape expectations for legacy VBA
fixtures:

| Upstream test | Current delta | Release decision |
|---|---|---|
| `test_save_with_vba` | WolfXL preserves `xl/sharedStrings.xml` and `xl/drawings/drawing1.xml`; openpyxl rewrites/prunes them. | Review before release. Do not call the corpus green until accepted or fixed. |
| `test_save_with_saved_comments` | WolfXL preserves `xl/sharedStrings.xml` and `xl/comments1.xml`; openpyxl rewrites comments to `xl/comments/comment1.xml`. | Review before release. Do not call the corpus green until accepted or fixed. |

These are not known read/write API failures: the saved packages are valid
ZIP/OOXML files and preserve more source content than openpyxl. They are still
release-blocking evidence until we decide whether exact openpyxl package
normalization is required for the 2.0 replacement claim.

## Compatibility Hardening Landed In This RC Pass

- Exposed `wolfxl.xml`, `wolfxl.reader.excel`, `wolfxl.styles.styleable`,
  `wolfxl.cell.read_only`, `wolfxl.worksheet._read_only`,
  `wolfxl.packaging.manifest`, and `wolfxl.workbook._writer` import paths used
  by the upstream corpus.
- Added `LXML` / `DEFUSEDXML` top-level flags and `wolfxl.open` alias.
- Accepted `load_workbook(..., keep_vba=...)`; `keep_vba=True` routes through
  modify mode so macro parts remain available for preservation.
- Added binary file-like save support.
- Normalized nonstandard workbook-part names such as `xl/workbook10.xml` to a
  temporary standard workbook path before handing the package to the Rust
  reader.
- Matched openpyxl bounded streaming semantics for sparse rows while
  preserving merged-subordinate number-format behavior.

## Stop Criteria Before Release

Before cutting 2.0, make an explicit call on the two VBA package-shape deltas:

1. Fix WolfXL to match openpyxl's pruning/renaming behavior for these legacy
   macro fixtures, then rerun the upstream corpus to 67 / 67.
2. Or document and allowlist them as intentional source-preservation behavior,
   with a functional Excel/openpyxl round-trip proof for both fixture outputs.

Do not publish release collateral until this decision is recorded.
