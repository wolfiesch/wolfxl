# Native Reader Calamine Exit Audit

Date: 2026-04-30
Status: in progress

## Current dependency truth

WolfXL's public `.xlsx` read path now uses `NativeXlsxBook` for normal eager
reads, streaming bootstrap reads, permissive malformed-workbook recovery, and
modify-mode bootstrap reads. The legacy `CalamineStyledBook` path is no longer
exported from the Python extension.

The `calamine-styles` dependency is still a live runtime dependency. It is not
safe to remove yet because these surfaces still compile against it:

- `src/calamine_styled_backend.rs` and helper modules remain in the worktree
  as retired legacy code until the dirty local edits touching those files are
  parked or intentionally integrated.
- `src/calamine_xlsb_xls_backend.rs` for value-only `.xlsb` and `.xls` reads.
- `crates/wolfxl-core`, which still advertises and implements multi-format
  preview reads through calamine-backed `.xlsx` / `.xls` / `.xlsb` / `.ods`
  paths.
- `crates/wolfxl-classify`, which still uses calamine types for binary-format
  source validation.
- Build metadata in `build.rs` / `src/lib.rs`, which still reports the
  Calamine-backed binary compatibility version.

The current honest launch claim is therefore: WolfXL has a native default
reader for normal eager `.xlsx` loads, streaming bootstrap reads, permissive
topology recovery, and modify-mode bootstrap reads, with no exported legacy
styled `.xlsx` backend. WolfXL is not yet dependency-free from Calamine across
the full read stack.

## Surfaces that can be removed after native parity

These should be removable after the dirty retired-backend edits are parked or
intentionally integrated:

- The styled `.xlsx` backend modules:
  - `src/calamine_styled_backend.rs`
  - `src/calamine_format_helpers.rs`
  - `src/calamine_record_format.rs`
  - `src/calamine_sheet_records.rs`
  - `src/calamine_style_dicts.rs`
  - `src/calamine_styled_array_formulas.rs`
  - `src/calamine_value_helpers.rs`
- README, compatibility, trust, release-note, and launch-copy references that
  describe `.xlsx` reading as calamine-backed.

## Surfaces that must remain unless scope changes

These cannot be removed without either building native binary readers or
intentionally dropping support:

- `CalamineXlsbBook` and `CalamineXlsBook` runtime bindings.
- `.xlsb` and `.xls` path/bytes loaders in `python/wolfxl/_workbook_sources.py`.
- `.xlsb` / `.xls` smoke and parity tests.
- `wolfxl-core` calamine-backed multi-format preview support.
- `wolfxl-classify` calamine-backed binary-format validation.
- Documentation that discloses `.xlsb` and `.xls` are value-only and currently
  use calamine-backed paths.

If we want a clean "no Calamine dependency" release claim, the remaining
product choices are:

1. Build native `.xlsb` and `.xls` readers and update the binary-format parity
   gates to use committed sidecar expectations instead of pandas+calamine.
2. Move `.xlsb` and `.xls` support behind an optional extra/feature so the main
   package can be dependency-free while an explicit compatibility build keeps
   those formats.
3. Drop `.xlsb` and `.xls` support for the native-reader release and document
   the regression clearly. This is the least attractive option because those
   reads are already public behavior.

## Documentation that must change before launch

Before any public native-reader announcement, audit and update:

- `README.md`
- `docs/index.md`
- `docs/migration/compatibility-matrix.md`
- `docs/trust/limitations.md`
- `docs/release-notes-*.md`
- `CHANGELOG.md`
- `Plans/launch-posts.md`
- `Plans/rfcs/067-native-reader.md`
- `crates/wolfxl-core/README.md`
- `crates/wolfxl-cli/README.md`

The docs should separate three facts:

- eager, streaming-bootstrap, and modify-bootstrap `.xlsx` reads are native by
  default;
- the legacy styled `.xlsx` backend is not exported from the Python extension;
- `.xlsb` and `.xls` remain value-only and calamine-backed until native binary
  readers land or support is explicitly scoped differently.

## Recommended staged removal plan

### Stage 1: finish `.xlsx` native coverage

- Keep streaming `.xlsx` bootstrap reads on the native shared-string/style/date
  helpers.
- Keep permissive malformed-workbook recovery on the native workbook topology
  loader, including the self-closing `<sheets/>` rels-graph fallback.
- Keep modify-mode bootstrap reads on `NativeXlsxBook`, while the patcher
  itself stays unchanged.
- Keep full openpyxl parity, native reader, streaming, and modify-mode tests in
  the gate.

### Stage 2: retire legacy styled `.xlsx`

- Delete the styled `.xlsx` calamine modules once no tests or runtime code
  import them.
- Keep `build_info()` distinguishing native `.xlsx` from compatibility binary
  readers.
- Run `rg "CalamineStyledBook|WOLFXL_CALAMINE_READER|calamine_styled"`.

### Stage 3: decide binary-format strategy

- If preserving `.xlsb` / `.xls`, implement native BIFF12/BIFF8 readers in
  `crates/wolfxl-reader` and replace pandas+calamine parity with committed
  expected sidecars.
- If making Calamine optional, split binary-format bindings behind a Cargo
  feature and Python extra, then ensure the default wheel does not pull
  `calamine-styles`.
- If dropping binary support, remove loaders/tests/docs for `.xlsb` / `.xls`
  and call the change out in release notes.

### Stage 4: remove dependency

- Remove `calamine-styles` from workspace dependencies and all crate
  dependency lists.
- Remove build metadata for `WOLFXL_DEP_CALAMINE_VERSION`.
- Run `cargo update`, `cargo check --all-targets --workspace`, `cargo test
  --workspace`, and the full Python suite.
- Run the final release gate:

```bash
rg "calamine_styles|calamine-styles|CalamineStyledBook|CalamineXlsbBook|CalamineXlsBook|WOLFXL_CALAMINE_READER" \
  src crates python Cargo.toml Cargo.lock README.md docs Plans tests
```

The only acceptable hits after a full dependency-free release should be
historical changelog/RFC context or explicit migration notes.
