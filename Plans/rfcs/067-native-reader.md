# RFC-067: Native WolfXL reader stack

Status: Draft
Owner: Native reader sprint
Phase: Pre-release dependency removal
Estimate: XL
Depends-on: existing read parity harness, committed XLSX/XLSB/XLS fixtures
Unblocks: removing `calamine-styles` from runtime dependencies and making an honest native-reader launch claim.

## Summary

WolfXL will replace its `calamine-styles` read dependency with native readers
for `.xlsx` / `.xlsm`, `.xlsb`, and `.xls`. The rollout stays opt-in until
native parity is proven. The final release gate is: native readers are default,
the `calamine-styles` dependency is gone, and public docs no longer describe
calamine as part of the runtime read stack.

## Public contract

The Python API stays stable:

- `wolfxl.load_workbook(...)` continues to accept paths, bytes-like objects,
  memoryviews, and binary file-like objects.
- `.xlsx` / `.xlsm` keep the full current read surface: values, formulas,
  styles, rich text, merged cells, dimensions, sheet features, metadata,
  streaming reads, and modify-mode compatibility.
- `.xlsb` and `.xls` preserve the current value-only contract: values,
  cached formula results, sheet ordering, dimensions, bounds, and bulk reads.
  Style access, modify mode, streaming, and password reads remain unsupported
  for those formats unless a later RFC expands scope.
- `.ods` remains unsupported and should keep raising a clear error.

## Implementation

### Phase 1: Reader seam

- Add `crates/wolfxl-reader` as the native reader foundation with shared
  workbook, sheet, and cell models that do not expose `calamine_styles` types.
- Keep the current Python path on calamine while native support grows behind an
  environment flag and/or shadow tests.
- Move helper logic out of `calamine_*` modules only after equivalent native
  call sites exist, to avoid broad churn before parity.

### Phase 2: Native XLSX / XLSM

- Parse the ZIP package, workbook relationships, workbook metadata, sheet
  order, sheet visibility, shared strings, and worksheet cell values.
- Add style-table parsing for fonts, fills, borders, alignment, number formats,
  row/column dimensions, and date/time conversion.
- Reuse existing WolfXL OOXML feature readers for hyperlinks, comments, tables,
  validations, conditional formats, freeze panes, page setup, workbook
  security, and document properties.
- Point streaming reads at the same native shared-string/style/date helpers as
  eager reads.

### Phase 3: Native XLSB

- Implement a BIFF12 record reader over the XLSB ZIP package.
- Support workbook/sheet discovery, shared strings, sparse worksheet cells,
  booleans, errors, numbers, strings, cached formula values, dimensions, and
  A1 range windows.
- Preserve value-only public behavior and existing `NotImplementedError` walls
  for styles and unsupported modes.

### Phase 4: Native XLS

- Implement an OLE Compound File reader sufficient to locate and read the
  `Workbook` / `Book` stream.
- Implement the BIFF8 records needed for current behavior: workbook globals,
  BoundSheet8, SST/Continue, LabelSst/Label, Number, RK/MulRK, BoolErr,
  Blank/MulBlank, Formula cached values, XF/FORMAT date typing, Dimensions,
  BOF/EOF, and sheet ordering.
- Preserve current value-only behavior and unsupported-mode errors.

### Phase 5: Flip and remove calamine

- Run native in shadow mode until side-by-side parity is stable.
- Make native the default reader.
- Remove `calamine-styles` from `Cargo.toml`, lockfiles, imports, build-info
  labels, and runtime docs.
- Rename or retire `calamine_*` helper modules once no live code depends on the
  old backend naming.

## Test plan

- Add crate-level parser tests in `wolfxl-reader` for workbook topology, shared
  strings, worksheet values, formulas, errors, booleans, sparse cells, and
  path/bytes loading.
- Add side-by-side native-vs-current tests while calamine is still present.
- Keep the existing openpyxl parity tests as the `.xlsx` acceptance gate.
- Replace final `.xlsb` / `.xls` parity gates with committed expected sidecars
  generated from trusted fixtures, because pandas/calamine cannot remain the
  final oracle after dependency removal.
- Run before flipping native default:
  - `cargo fmt --all --check`
  - `cargo check --all-targets --workspace`
  - `cargo test --workspace`
  - `uv run maturin develop`
  - `uv run pytest tests/parity/test_read_parity.py tests/parity/test_streaming_parity.py tests/parity/test_xlsb_reads.py tests/parity/test_xls_reads.py tests/test_wolfxl_compat.py tests/test_format_dispatch.py -q`
  - Full `uv run pytest` before release.

## Acceptance

- `WOLFXL_NATIVE_READER=1` passes the full XLSX parity and compatibility suite.
- `.xlsb` and `.xls` pass the committed value-only fixture sidecars.
- `rg "calamine_styles|calamine-styles" src crates python Cargo.toml Cargo.lock`
  finds no live runtime dependency references.
- README, migration docs, trust docs, and launch copy describe the native
  reader accurately and disclose value-only binary-format support.
