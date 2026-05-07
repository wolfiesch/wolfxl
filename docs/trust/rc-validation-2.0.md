# WolfXL 2.0 RC Validation Log

Status: pre-release hardening. Do not tag, publish, or cut 2.0 from this
checkpoint.

## Green Gates

| Gate | Result |
|---|---:|
| `uv run maturin develop` | Pass |
| Full Python suite with optional deps (`openpyxl==3.1.5`, `pillow`) | 2658 passed / 42 skipped |
| Openpyxl compat oracle | 79 / 79 |
| XLS/XLSB optional slice with `python-calamine` | Pass |
| External oracle fixture pack | Pass |
| `cargo test --workspace -q` | Pass |
| `cargo fmt --check` | Pass |
| `cargo check -q` | Pass |
| `ruff check` focused on touched Python files | Pass |

## Upstream Openpyxl VBA Corpus

The upstream openpyxl 3.1.5 `test_vba.py` file now passes under the
`openpyxl -> wolfxl` shim.

| Upstream test | Result | Resolution |
|---|---:|---|
| `test_save_with_vba` | Pass | WolfXL now prunes unused shared strings and legacy form-control drawing parts to match openpyxl's saved package shape. |
| `test_save_with_saved_comments` | Pass | WolfXL now relocates legacy `xl/commentsN.xml` parts to `xl/comments/commentN.xml` and rewrites rel/content-type references. |
| `test_content_types` | Pass | Content-type overrides remain unique after normalization. |
| `test_save_without_vba` | Pass | Read-mode macro saves continue to satisfy openpyxl's macro-removal expectation. |

The previous strict package-shape blockers are fixed rather than allowlisted.

## Compatibility Hardening Landed In This RC Pass

- Exposed `wolfxl.xml`, `wolfxl.reader.excel`, `wolfxl.styles.styleable`,
  `wolfxl.cell.read_only`, `wolfxl.worksheet._read_only`,
  `wolfxl.packaging.manifest`, and `wolfxl.workbook._writer` import paths used
  by the upstream corpus.
- Added `LXML` / `DEFUSEDXML` top-level flags and `wolfxl.open` alias.
- Accepted `load_workbook(..., keep_vba=...)`; `keep_vba=True` routes through
  modify mode so macro parts remain available for preservation.
- Added openpyxl-compatible VBA package-shape normalization for source-backed
  saves, including shared-string pruning, legacy control drawing pruning, and
  saved-comment part relocation.
- Added binary file-like save support.
- Normalized nonstandard workbook-part names such as `xl/workbook10.xml` to a
  temporary standard workbook path before handing the package to the Rust
  reader.
- Matched openpyxl bounded streaming semantics for sparse rows while
  preserving merged-subordinate number-format behavior.

## Stop Criteria Before Release

The VBA package-shape decision has been made in favor of exact openpyxl
normalization and is now covered by local regression tests plus the upstream
`test_vba.py` shim run. Before tagging 2.0, rerun the full release gate suite
from a clean checkout and keep this document's counts current.
