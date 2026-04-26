# RFC-020: T1.5 — Document Properties Mutation (modify mode)

Status: Shipped
Owner: pod-P2
Phase: 3
Estimate: S
Depends-on: 013 (file_adds for the optional-`docProps/core.xml` add case; supersedes the original "no deps" frontmatter — see Acceptance below)
Unblocks: (none)

## 1. Problem Statement

In modify mode (`load_workbook(path, modify=True)`), setting any document property and then calling `save()` raises `NotImplementedError`:

```python
# python/wolfxl/_workbook.py:286-290
if self._properties_dirty:
    raise NotImplementedError(
        "Rewriting document properties on an existing file is a "
        "T1.5 follow-up. Use ``Workbook()`` + save for new files."
    )
```

The properties are already captured correctly: the getter (`_workbook.py:223-248`) populates `_properties_cache` from `_rust_reader.read_doc_properties()`, and any subsequent attribute write on the `DocumentProperties` dataclass flips `_properties_dirty = True` via `__setattr__` (`python/wolfxl/packaging/core.py:49-53`). All the user-facing state is correct at save time — only the patcher path refuses to act on it.

**Target behavior**: when `_properties_dirty` is True in modify mode, regenerate `docProps/core.xml` (and optionally `docProps/app.xml`) and replace the corresponding ZIP entries in the output file. Every other entry is raw-copied unchanged, preserving charts, macros, images, VBA, and all other parts that wolfxl does not understand.

## 2. OOXML Spec Surface

### docProps/core.xml (primary target)

Governed by ECMA-376 Part 2 (Open Packaging Conventions), §11 (Core Properties). Relationship type:
`http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties`

Root element: `<cp:coreProperties>` with five namespace declarations:
- `xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"`
- `xmlns:dc="http://purl.org/dc/elements/1.1/"`
- `xmlns:dcterms="http://purl.org/dc/terms/"`
- `xmlns:dcmitype="http://purl.org/dc/dcmitype/"`
- `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"`

Elements (all optional except `dc:creator` and `dcterms:modified` which Excel expects):

| Element | Type | Notes |
|---|---|---|
| `dc:title` | string | |
| `dc:subject` | string | |
| `dc:creator` | string | required in practice |
| `dc:description` | string | |
| `dc:identifier` | string | |
| `dc:language` | string | |
| `cp:keywords` | string | |
| `cp:category` | string | |
| `cp:contentStatus` | string | must appear after `cp:category`, before `dcterms:created` |
| `cp:lastModifiedBy` | string | |
| `cp:lastPrinted` | dateTime | ISO-8601 with `xsi:type="dcterms:W3CDTF"` |
| `cp:revision` | string | integer-valued in practice |
| `cp:version` | string | |
| `dcterms:created` | dateTime | `xsi:type="dcterms:W3CDTF"` required |
| `dcterms:modified` | dateTime | `xsi:type="dcterms:W3CDTF"` required |

Schema order is strict for Excel compatibility. The existing emitter at `crates/wolfxl-writer/src/emit/doc_props.rs:68-133` already enforces the correct order and can be called directly.

### docProps/app.xml (secondary target)

Governed by ECMA-376 Part 1 §15.2.12.3 (Extended File Properties). Relationship type:
`http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties`

Contains `Application`, `AppVersion` (must be `XX.YYYY` per §22.2.2.3, not semver), `HeadingPairs`, `TitlesOfParts` (sheet name list), plus optional `Company` and `Manager`. Because `TitlesOfParts` mirrors the live sheet list, `app.xml` must be regenerated any time sheets are added or renamed too — but in modify mode neither operation is supported yet. For this RFC, `app.xml` is regenerated whenever `_properties_dirty` is True, using the sheet names already known to the patcher.

The existing emitter at `crates/wolfxl-writer/src/emit/doc_props.rs:136-192` already handles `Company` and `Manager`; wolfxl's `DocumentProperties` Python class (`python/wolfxl/packaging/core.py`) does NOT expose `company` or `manager` fields — those are `DocProperties` fields in the Rust model only. Therefore, `app.xml` mutation in this RFC is limited to regenerating the sheet-name vector correctly; Company/Manager pass-through from the source file is punted to §10.

## 3. openpyxl Reference

Source: `.venv/lib/python3.14/site-packages/openpyxl/packaging/core.py`

- `DocumentProperties` uses descriptor-driven serialization (`Serialisable`). Every field listed in `__elements__` is round-tripped through `NestedText` or `QualifiedDateTime`.
- openpyxl always stamps `modified = datetime.now(utc)` on save regardless of whether properties were touched (`core.py:100`, `modified = modified or now`). We match this behavior.
- `dcterms:created` defaults to `now()` on construction if not supplied (`core.py:108`). We preserve the original file's `created` timestamp when the user does not explicitly set it.
- `cp:lastModifiedBy` is not auto-populated by openpyxl; users set it manually. wolfxl follows suit.
- Fields NOT copied from openpyxl: the descriptor system, `Serialisable` base class, lxml Element tree. wolfxl uses the existing string-builder emitter.
- `ExtendedProperties` in `openpyxl/packaging/extended.py` exposes `Company`, `Manager`, `HyperlinkBase`, `AppVersion`. wolfxl's Python `DocumentProperties` does not expose these, so this RFC only touches the `app.xml` entries wolfxl controls (Application, AppVersion, sheet list). The others are preserved verbatim from the source ZIP (they live in `app.xml` which we rewrite — see §5 for the preservation strategy).

## 4. WolfXL Surface Area

### 4.1 Python coordinator

File: `python/wolfxl/_workbook.py`

Lines `286-290`: delete the `NotImplementedError` block. Replace with a call to a new patcher method:

```python
if self._properties_dirty:
    self._flush_properties_to_patcher()
```

New private method `_flush_properties_to_patcher(self)` (approximately `_workbook.py:310` region):
- Extracts `_properties_cache` fields into a `dict[str, str]`.
- Calls `self._rust_patcher.queue_properties(payload: dict)`.
- Resets `_properties_dirty = False`.

The `payload` dict key set matches what `_flush_workbook_writes` already sends to `writer.set_properties()` (`_workbook.py:318-329`) — same field names, same Python → Rust boundary.

### 4.2 Patcher (modify mode)

New module: `src/wolfxl/properties.rs`

No separate module is strictly needed — the logic is thin enough to inline into `src/wolfxl/mod.rs` — but a dedicated file follows the existing pattern (cf. `sheet_patcher.rs`, `styles.rs`).

**Why RFC-011 is NOT a dependency**: `docProps/core.xml` is a ~600-byte flat file. There is no existing XML to merge into — the correct approach is a full rewrite on every dirty save, which is exactly what the write-mode emitter already does. RFC-011's block-merger primitive is designed for splicing new XML blocks into large structured documents (like `xl/workbook.xml` or `xl/worksheets/sheetN.xml`). For a small flat metadata file, the overhead and complexity of that approach would exceed the benefit. The writer's `emit_core` function (`crates/wolfxl-writer/src/emit/doc_props.rs:68`) already produces correct output. This RFC reuses it directly.

**Cross-crate reuse**: `wolfxl-writer` is a separate crate, not a library dependency of the `src/` PyO3 cdylib. Two options:

1. **Extract to `wolfxl-core`**: move `emit_core` / `emit_app` into a new `wolfxl-core::emit::doc_props` submodule so both the patcher and the writer can call it. `wolfxl-core` has no PyO3 dependency (CLAUDE.md invariant).
2. **Duplicate the small function**: copy the ~80-line string-builder into `src/wolfxl/properties.rs`. Accept the duplication; the function is pure and simple, and both copies would be tested independently.

Recommendation: **Option 2 for now** (½-day scope), with a TODO to consolidate into `wolfxl-core` when Option 1 is needed for another RFC. The function is not large enough to justify the crate-dependency surgery in this RFC's S-estimate.

Public Rust API in `src/wolfxl/properties.rs`:

```rust
use crate::wolfxl::DocPropertiesPayload;

pub struct DocPropertiesPayload {
    pub title: Option<String>,
    pub subject: Option<String>,
    pub creator: Option<String>,
    pub keywords: Option<String>,
    pub description: Option<String>,
    pub last_modified_by: Option<String>,
    pub category: Option<String>,
    pub content_status: Option<String>,
    pub created_iso: Option<String>,  // preserve original; None → now()
    pub modified_iso: Option<String>, // usually None → now()
    pub sheet_names: Vec<String>,     // for app.xml TitlesOfParts
}

/// Emit docProps/core.xml bytes for the patcher.
pub fn rewrite_core_props(payload: &DocPropertiesPayload) -> Vec<u8>;

/// Emit docProps/app.xml bytes for the patcher.
pub fn rewrite_app_props(payload: &DocPropertiesPayload) -> Vec<u8>;
```

New `#[pymethods]` entry on `XlsxPatcher` in `src/wolfxl/mod.rs`:

```rust
fn queue_properties(&mut self, props: &Bound<'_, PyDict>) -> PyResult<()>;
```

This stores a `DocPropertiesPayload` on the `XlsxPatcher` struct (new field `queued_props: Option<DocPropertiesPayload>`). On `do_save`, if `queued_props.is_some()`, the two new bytes are inserted into `file_patches` under keys `"docProps/core.xml"` and `"docProps/app.xml"` before the ZIP rewrite loop at `src/wolfxl/mod.rs:343`.

ZIP parts touched:
- `docProps/core.xml` — full rewrite
- `docProps/app.xml` — full rewrite (sheet list re-derived from `sheet_paths.keys()`)

### 4.3 Native writer (write mode)

No changes. The native writer already emits both parts correctly via `emit_core` and `emit_app` (`crates/wolfxl-writer/src/emit/doc_props.rs`). This RFC is patcher-only.

## 5. Algorithm

```
save(filename):
  if _rust_patcher is not None:
    if _properties_dirty:
      payload = build_payload(_properties_cache, sheet_names)
      _rust_patcher.queue_properties(payload)
      _properties_dirty = False
    for ws in _sheets.values():
      ws._flush()
    _rust_patcher.save(filename)

build_payload(props, sheet_names):
  # Preserve original created timestamp if user did not explicitly set it.
  # modified always stamps to now() (matching openpyxl behavior).
  return {
    title: props.title,
    subject: props.subject,
    creator: props.creator or "wolfxl",
    keywords: props.keywords,
    description: props.description,
    last_modified_by: props.lastModifiedBy,
    category: props.category,
    content_status: props.contentStatus,
    created_iso: props.created.isoformat() if props.created else None,
    modified_iso: None,   # let Rust stamp now()
    sheet_names: list(sheet_names),
  }

do_save (Rust):
  if queued_props.is_some():
    core_bytes = rewrite_core_props(&queued_props)
    app_bytes  = rewrite_app_props(&queued_props)
    file_patches.insert("docProps/core.xml", core_bytes)
    file_patches.insert("docProps/app.xml",  app_bytes)
  # ... existing ZIP rewrite loop (mod.rs:343-359)
```

Idempotency: calling `save()` twice with no additional mutations does NOT regenerate the parts on the second call, because `_properties_dirty` is cleared after the first flush and `queue_properties` is never called again.

**app.xml Company/Manager preservation**: when rewriting `app.xml`, the patcher currently has no access to the original file's Company and Manager values (they live in the source ZIP's `app.xml` which we never parse). These fields are silently dropped on the first dirty save. This is documented as a known loss in §7 and punted to a follow-up.

## 6. Test Plan

Standard 6-layer matrix:

1. **Rust unit** (`src/wolfxl/properties.rs` tests): `rewrite_core_props` produces well-formed XML containing each field; XML escaping works; WOLFXL_TEST_EPOCH=0 freezes both timestamps.

2. **Golden round-trip** (`tests/diffwriter/`): modify a file with `title="Quarterly Report"` and `creator="Alice"`, save, re-open with `load_workbook`, assert `wb.properties.title == "Quarterly Report"` and `wb.properties.creator == "Alice"`. Use `WOLFXL_TEST_EPOCH=0`.

3. **openpyxl parity** (`tests/parity/`): open the same fixture with openpyxl, set the same properties, save, re-open with openpyxl, assert identical field values. Confirm wolfxl output matches.

4. **LibreOffice** (manual / CI smoke): open wolfxl-patched file in LibreOffice, check File > Properties. Not automated; noted in test plan.

5. **Cross-mode**: confirm that non-dirty modify saves do NOT touch `docProps/core.xml` (byte-for-byte identity of that entry in the output ZIP).

6. **Regression fixture**: add a file whose `docProps/core.xml` contains all 14 optional fields. Verify each survives a dirty-and-save round-trip without truncation.

RFC-specific cases:
- Setting only `wb.properties.title` in modify mode; confirm `creator` still populated from original file.
- `wb.properties.modified` explicitly set; confirm that value wins over `now()`.
- `WOLFXL_TEST_EPOCH=0` freezes `dcterms:modified` to `1970-01-01T00:00:00Z`.
- XML special characters in title (`"A & B < C"`) are escaped in output.

## 7. Migration / Compat Notes

- **openpyxl auto-stamps `modified`**: openpyxl always sets `modified` to `now()` on save, even for read-only opens. wolfxl matches this: `dcterms:modified` always reflects save time. This is a semantic parity point, not a divergence.
- **`created` preservation**: openpyxl also resets `created` to `now()` when the field was absent from the source file. wolfxl preserves the original `created` if the user did not explicitly set it. Minor divergence, intentionally better behavior.
- **Company/Manager fields**: on the first dirty save, these are lost from `app.xml` (they are only in the Rust model for write mode, not surfaced in Python's `DocumentProperties`). This is a known regression vs. openpyxl when Company/Manager were present in the source file. Mitigated by keeping them if `_properties_dirty` is False (no rewrite at all).
- **`WOLFXL_TEST_EPOCH=0`**: the epoch override already implemented in `crates/wolfxl-writer/src/zip/mod.rs` is checked in `current_timestamp_iso8601()`. The patcher's equivalent must implement the same check (or import the same utility) so golden tests are stable.
- No feature flag needed — this is a pure bug-fix; the old path raises, so there is no existing behavior to preserve.

## 8. Risks & Open Questions

1. **Company/Manager loss on first dirty save**: When a source file has Company/Manager in `app.xml` and the user sets any property, those values are dropped. Resolution: document the loss in the NotImplementedError-removal commit message; file a follow-up to parse and pass Company/Manager through Python's `DocumentProperties` (Phase 3 hardening, no new RFC needed).

2. **`app.xml` HeadingPairs/TitlesOfParts**: The existing `emit_app` in the writer crate derives the sheet list from `wb.sheets`. The patcher version must use `sheet_paths.keys()` from `XlsxPatcher`. The order of keys in `HashMap` is non-deterministic; a `BTreeMap` or sorted Vec must be used for the sheet list to match the original order. Proposed resolution: store an ordered `sheet_names: Vec<String>` on `XlsxPatcher` alongside `sheet_paths`, populated in insertion order during `open()`.

3. **`docProps/core.xml` might be missing from source file**: The OOXML spec says it's optional. If the source file has no such entry, the patcher should add it rather than skip. The existing ZIP rewrite loop at `src/wolfxl/mod.rs:323` iterates only existing entries — adding a new entry requires a separate write step after the loop. Resolution: after the loop, check whether `docProps/core.xml` was written; if not, write it as a new entry.

4. **RFC-011 dependency re-evaluation**: RFC-020 was listed as depending on RFC-011 in the INDEX. After reading the implementation, RFC-011 (XML block merger) is NOT needed here — full rewrite of a small flat file is correct and simpler. The dependency is dropped in this RFC's frontmatter. INDEX.md should be updated accordingly.

## 9. Effort Breakdown

| Task | LOC est. | Days |
|---|---|---|
| `src/wolfxl/properties.rs` (rewrite_core_props, rewrite_app_props) | ~120 | 0.25 |
| `src/wolfxl/mod.rs` (queue_properties method, queued_props field, do_save wiring) | ~40 | 0.25 |
| `python/wolfxl/_workbook.py` (remove raise, add _flush_properties_to_patcher) | ~25 | 0.25 |
| Tests (Rust unit + Python golden + parity) | ~80 | 0.5 |
| **Total** | **~265** | **~1.25** |

## 10. Out of Scope

- **Company / Manager round-trip in modify mode**: requires adding these fields to Python's `DocumentProperties` class and the `queue_properties` payload. Deferred to Phase 3 hardening.
- **`lastPrinted`, `revision`, `version`, `identifier`, `language`**: Python's `DocumentProperties` exposes them (`packaging/core.py:37-44`) but `_flush_workbook_writes` does not currently send them to the writer. Normalize both paths together in a follow-up.
- **`app.xml` HeadingPairs/TitlesOfParts for chartsheets or hidden sheets**: not supported in wolfxl at all yet.
- **RFC-011 XML-block-merger**: not needed for this RFC (see §4.2 and §8 item 4).
- **app.xml parsing from source** (to preserve Company/Manager on round-trip): future work.

## Acceptance

Shipped via 4-commit slice on `feat/native-writer` (commits `349a302..<this commit>`) on 2026-04-25, bundled with RFC-013 (commits `2f3d5a7..ee9c166`) which provided `file_adds` for the optional-`docProps/core.xml` add path (§8 risk #3).

**Live behavior**:
- `wb.properties.title = "X"; wb.save(path)` round-trips on existing files in modify mode.
- All 11 fields the writer round-trips (title, subject, creator, keywords, description, lastModifiedBy, category, contentStatus, created, modified) round-trip in modify mode too.
- `WOLFXL_TEST_EPOCH=0` produces deterministic `dcterms:modified` for golden-file tests.
- `dcterms:modified` re-stamps to save-time on dirty save (semantically correct: saving IS a modification). User-explicit `props.modified = ...` bypasses the re-stamp via the new `_user_set` per-field tracker on `DocumentProperties`.
- No-op modify-mode save remains byte-identical to source (regression-guarded by `test_no_dirty_save_is_byte_identical`).

**Implementation**:
- `src/wolfxl/properties.rs` — Lift-and-shift of writer's `emit_core` / `emit_app`, retargeted from `&Workbook` to `&DocPropertiesPayload`. Marked TODO for consolidation per §4.2 Option 2 once a third caller appears.
- `src/wolfxl/mod.rs` — `queued_props: Option<DocPropertiesPayload>` field, `queue_properties` PyO3 method, Phase 2.5d in `do_save` (after Phase 2.5c content-types aggregation). Routing: `file_patches` if source already has the entry, RFC-013 `file_adds` if not.
- `python/wolfxl/_workbook.py` — `_flush_properties_to_patcher` replaces the T1.5 `NotImplementedError`. Filters `None` values before crossing the PyO3 boundary; honors per-field user-set tracking for `modified`.
- `python/wolfxl/packaging/core.py` — `DocumentProperties._user_set: set[str]` populated by `__setattr__` after `_attach_workbook`, lets the flush distinguish hydrated-from-source vs user-mutated.

**Coverage**: `tests/test_modify_properties.py` adds 10 integration tests; existing T1.5-raise tests in `tests/test_workbook_writes_t1.py` and `tests/test_modify_mode_independence.py` updated accordingly.

**Known regressions** (§7 / §10): source `<Company>` and `<Manager>` from `app.xml` are dropped on dirty save (Python's `DocumentProperties` doesn't expose those fields). Pinned by `test_app_xml_drops_company_manager_known_loss` so an accidental fix surfaces here, not in user reports. Same for `lastPrinted`, `revision`, `version`, `identifier`, `language`.
