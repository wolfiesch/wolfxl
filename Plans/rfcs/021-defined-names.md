# RFC-021: T1.5 — Defined Names Mutation (modify mode)

Status: Shipped
Owner: pod-P2
Phase: 3
Estimate: M
Depends-on: RFC-011, RFC-012
Unblocks: RFC-030, RFC-031, RFC-034, RFC-035, RFC-036

## 1. Problem Statement

In modify mode, adding a new defined name (named range) and then saving raises `NotImplementedError`:

```python
# python/wolfxl/_workbook.py:291-294
if self._pending_defined_names:
    raise NotImplementedError(
        "Adding defined names to an existing file is a T1.5 "
        "follow-up. Use ``Workbook()`` + save for new files."
    )
```

The user-visible queue is already populated correctly: `defined_names.__setitem__` (via `DefinedNameDict`) routes new entries to `_pending_defined_names` (`python/wolfxl/workbook/defined_name.py`). In write mode, `_flush_workbook_writes` (`_workbook.py:332-352`) drains this queue into the Rust writer. In modify mode, the queue is simply refused.

**Target behavior**: when `_pending_defined_names` is non-empty in modify mode, parse the existing `<definedNames>` block in `xl/workbook.xml`, merge the new entries (upsert by name+localSheetId), and write the updated `xl/workbook.xml` back into the output ZIP. Existing entries not referenced in the pending queue are preserved verbatim.

This also covers the update (replace) case: if a name already exists in the file, the new formula replaces it.

## 2. OOXML Spec Surface

Governed by ECMA-376 Part 1 §18.2.5 (`CT_DefinedName`) and §18.2.27 (`CT_Workbook`).

### CT_DefinedName element

```xml
<definedName
    name="X"
    localSheetId="0"
    hidden="1"
    comment="..."
    customMenu="..."
    description="..."
    help="..."
    statusBar="..."
    shortcutKey="..."
    function="1"
    vbProcedure="1"
    xlm="1"
    functionGroupId="14"
    publishToServer="1"
    workbookParameter="1"
>Sheet1!$A$1:$B$5</definedName>
```

- `name` (required): unique per workbook scope or per-sheet scope. Must start with a letter or `_`, no spaces, max 255 chars.
- Text content: the formula/range expression (not prefixed with `=` in the XML).
- `localSheetId` (optional): 0-based sheet position index. Absent means workbook-scope.
- `hidden` (optional): `"1"` or `"0"`. Controls Name Manager visibility.
- All other attributes are rare; wolfxl passes them through unmodified from existing entries.

### CT_Workbook child ordering (§18.2.27)

The OOXML schema enforces a specific ordering of children within `<workbook>`. `<definedNames>` must appear after `<sheets>` and before `<calcPr>`. The existing patcher's `xl/workbook.xml` content respects this order; the RFC-021 merger preserves it.

### Built-in names (`_xlnm.` prefix)

Reserved names such as `_xlnm.Print_Area`, `_xlnm.Print_Titles`, `_xlnm._FilterDatabase` are serialized with their full `_xlnm.` prefix in the XML. RFC-021 does not generate builtins from user input; existing builtins are preserved verbatim on round-trip via the upsert key `(name, local_sheet_id)`.

## 3. openpyxl Reference

- `DefinedName` fields: `name`, `comment`, `customMenu`, `description`, `help`, `statusBar`, `localSheetId`, `hidden`, `function`, `vbProcedure`, `xlm`, `functionGroupId`, `shortcutKey`, `publishToServer`, `workbookParameter`, `attr_text`.
- **`attr_text` vs `value`**: openpyxl exposes `value` as an `Alias("attr_text")`. wolfxl's `DefinedName` accepts both `value=` and `attr_text=` keyword arguments and stores them identically.
- `localSheetId` is the sheet's **position index** (0-based), NOT its name. This matches openpyxl.

## 4. WolfXL Surface Area

### 4.1 Python coordinator

File: `python/wolfxl/_workbook.py`

The pre-shipping `NotImplementedError("T1.5")` block at line 291-294 is replaced by a single call to `self._flush_defined_names_to_patcher()`. The new method iterates `_pending_defined_names.items()`, builds a flat dict payload (`name`, `formula`, optional `local_sheet_id`/`hidden`/`comment`), and routes each through `self._rust_patcher.queue_defined_name(payload)`. `None`-valued optional fields are filtered before crossing the PyO3 boundary so the Rust extractors see a clean missing-key signal (matches the convention in `_flush_properties_to_patcher`).

### 4.2 Patcher (modify mode)

New module: `src/wolfxl/defined_names.rs`

**Design choice**: streaming splice over the workbook XML, NOT a full rewrite. We scan `xl/workbook.xml` once with `quick_xml::Reader` to capture three byte positions:

1. The byte offset right after `</sheets>` (the inject point if no `<definedNames>` block exists).
2. The outer byte range of any existing `<definedNames>...</definedNames>` block.
3. The inner byte range (between the wrapper tags).

Children of any existing block are extracted as verbatim byte slices keyed by `(name, local_sheet_id)`. Upserts in `names` either replace a matching child's bytes (with attribute preservation — see below) or are appended at the end of the merged inner block. Children not referenced by any upsert flow through untouched.

The output is constructed as `[..pre_block][merged_block][..post_block]` where `pre_block`/`post_block` are byte slices of the source. Every child of `<workbook>` outside the spliced region survives byte-for-byte.

**Attribute preservation on update**: when an upsert matches an existing entry, we re-emit the start tag with the original attribute order plus any overrides. This preserves rare attributes (`customMenu`, `description`, `help`, `statusBar`, `shortcutKey`, etc.) that the Python API does not expose.

**Public Rust API in `src/wolfxl/defined_names.rs`**:

```rust
pub struct DefinedNameMut {
    pub name: String,
    pub formula: String,
    pub local_sheet_id: Option<u32>,
    pub hidden: Option<bool>,
    pub comment: Option<String>,
}

pub fn merge_defined_names(
    workbook_xml: &[u8],
    names: &[DefinedNameMut],
) -> Result<Vec<u8>, String>;
```

**Integration point for RFC-012**: `DefinedNameMut::formula` is a plain string. RFC-036 (`move_sheet`) will call RFC-012's translator on each formula before invoking `merge_defined_names`. The merger never inspects formula contents — it just escapes the text and writes it through.

New field on `XlsxPatcher` (`src/wolfxl/mod.rs`):
```rust
queued_defined_names: Vec<DefinedNameMut>,
```

New `#[pymethods]` entry: `fn queue_defined_name(&mut self, payload: &Bound<'_, PyDict>) -> PyResult<()>`.

In `do_save`, Phase 2.5f drains the queue:

```rust
if !self.queued_defined_names.is_empty() {
    let wb_xml = ooxml_util::zip_read_to_string(&mut zip, "xl/workbook.xml")?;
    let updated = defined_names::merge_defined_names(wb_xml.as_bytes(), &self.queued_defined_names)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("defined-names merge: {e}")))?;
    file_patches.insert("xl/workbook.xml".to_string(), updated);
}
```

ZIP parts touched: `xl/workbook.xml` only.

### 4.3 Native writer (write mode)

No changes. The native writer continues to emit `<definedNames>` from the structured `Workbook` model.

## 5. Algorithm

```
# Python — save(filename) modify-mode branch
if _rust_patcher is not None:
    if _pending_defined_names:
        for name, dn in _pending_defined_names.items():
            payload = {"name": dn.name, "formula": dn.value}
            if dn.localSheetId is not None: payload["local_sheet_id"] = dn.localSheetId
            if dn.hidden:                   payload["hidden"]         = True
            if dn.comment is not None:      payload["comment"]        = dn.comment
            _rust_patcher.queue_defined_name(payload)
        _pending_defined_names.clear()
    # ... flush sheets, call patcher.save(filename)

# Rust — merge_defined_names(workbook_xml, names)
# 1. Scan workbook.xml: find </sheets> end and any existing <definedNames> range.
# 2. If names is empty → return source bytes verbatim (modify-mode no-op).
# 3. Extract each existing <definedName> child as raw bytes + (name, local_sheet_id) key.
# 4. For each upsert in names:
#    - if (name, local_sheet_id) matches an existing child → replace its bytes
#      via serialize_upsert_over_existing(raw, upsert) (preserves all source attrs).
#    - else → append a freshly serialized child at the tail.
# 5. Splice [..pre_block][<definedNames>...inner...</definedNames>][..post_block].
```

**Idempotency**: `merge_defined_names` is pure on its inputs. Calling `save()` twice with no additional mutations is safe — `queued_defined_names` is empty after each save and the second call short-circuits at the no-op guard.

**Upsert key**: `(name, local_sheet_id)` — two defined names with the same `name` but different scopes are distinct.

**XML escaping**: formula text is text-escaped; attribute values are attribute-escaped via local helpers (mirrors writer escape semantics).

## 6. Test Plan

1. **Rust unit** (`src/wolfxl/defined_names.rs#tests`): 11 cases — inject when no block, append when block exists, upsert preserves attrs, workbook-scope vs sheet-scope coexist, XML escaping, empty-names identity (with and without existing block), built-in `_xlnm.Print_Area` round-trip, `hidden="1"` emission, `localSheetId="N"` emission, missing `</sheets>` errors.
2. **Pytest** (`tests/test_defined_names_modify.py`): 14 cases across `attr_text`/`value` alias contract, add new, update existing, preserve unrelated names, built-in print-area round-trip, sheet-scope routing, no-op guard, `hidden=True`, mixed cell+name save, queue cleared after save.
3. **Parity**: existing parity sweep (`tests/parity/`) green.
4. **LibreOffice**: deferred to manual smoke per RFC-021 §6.
5. **Cross-mode**: write-mode path unchanged (`_flush_workbook_writes` still works).
6. **Regression**: `tests/test_defined_names_t1.py` continues to pass; the previous "raises NotImplementedError" test is upgraded to `test_defined_names_modify_mode_round_trip` covering the new behaviour.

## 7. Migration / Compat Notes

- `value` and `attr_text` keyword arguments accepted on `DefinedName(...)`. Conflicting values raise `TypeError`; missing both raises `TypeError`.
- `localSheetId` is position-based (0-based index). Same as openpyxl.
- Formula strings are stored verbatim; no normalization or validation.
- The previously-shipped `NotImplementedError("T1.5")` is gone. Callers that relied on catching it stop seeing it. Intended upgrade path.

## 8. Risks & Open Questions

1. **RFC-011 dependency**: the splice approach is bespoke (not via `wolfxl-merger`, which targets sheet XML). When RFC-011 grows a generic workbook-scope splice utility, RFC-021 should consolidate.
2. **RFC-012 integration seam**: documented in `defined_names.rs` module docs. RFC-036 will call RFC-012 before RFC-021.
3. **Built-in names from user code**: a user could add `DefinedName(name="_xlnm.Print_Area", ...)` and the merger emits the prefixed name verbatim. Excel reads this correctly.
4. **No delete op**: callers cannot remove a defined name in modify mode this slice. Follow-up: `queue_remove_defined_name`.

## 9. Effort Breakdown — Actuals

| Task | LOC actual |
|---|---|
| `src/wolfxl/defined_names.rs` | ~525 (incl. 11 tests) |
| `src/wolfxl/mod.rs` (queue_defined_name + Phase 2.5f) | ~80 |
| `python/wolfxl/_workbook.py` (_flush_defined_names_to_patcher + dispatch) | ~40 |
| `python/wolfxl/workbook/defined_name.py` (attr_text alias + __init__) | ~70 |
| `tests/test_defined_names_modify.py` | ~310 (14 cases) |
| `tests/test_workbook_writes_t1.py` (raises → round-trip swap) | ~10 |
| **Total** | **~1035 LOC** |

## 10. Out of Scope

- Formula reference translation (RFC-012).
- Auto-reindex of `localSheetId` on `move_sheet` (RFC-036).
- Delete operation in modify mode.
- Validation of name syntax (starts with letter/underscore, no spaces, max 255 chars).
- Rare attributes (`customMenu`, `description`, `help`, `statusBar`, `shortcutKey`, `function`, `vbProcedure`, `xlm`, `functionGroupId`, `publishToServer`, `workbookParameter`) settable from the Python API. They survive verbatim on update of an existing name; new names omit them.

## Acceptance

Shipped 2026-04-25 on branch `feat/rfc-021-defined-names`.

Verification matrix:

- Rust unit tests in `src/wolfxl/defined_names.rs` (11 cases) — green via `cargo test -p wolfxl-core -p wolfxl-writer -p wolfxl-rels -p wolfxl-merger`. The patcher's cdylib does not link standalone via `cargo test`; defined-names tests live inline in the module and compile under `cargo build` (same precedent as `properties.rs`, `hyperlinks.rs`).
- pytest `tests/test_defined_names_modify.py` (14 cases) — green.
- `tests/test_defined_names_t1.py` (8 cases) and the upgraded `test_defined_names_modify_mode_round_trip` in `tests/test_workbook_writes_t1.py` — green.
- `tests/parity/` (98 cases) — green; no regressions.

Commit SHAs: see git log for `feat/rfc-021-defined-names` branch.
