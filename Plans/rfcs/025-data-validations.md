# RFC-025: Sheet-Scoped Data Validations in Modify Mode

Status: Shipped
Owner: pod-P4
Phase: 3
Estimate: M
Depends-on: RFC-011
Unblocks: RFC-030, RFC-031, RFC-035

## Acceptance: shipped via 5-commit slice on `feat/native-writer` (commits 6f6525b → ba3de4b) on 2026-04-25

Activates RFC-011's `queued_blocks` plumbing end-to-end. `ws.data_validations.append(dv)` now works in modify mode: the patcher reads any existing `<dataValidations>` block out of the source sheet XML, prepends those rules verbatim (byte-slice capture preserves escaped content, self-closing forms, and unknown attributes), appends the queued patches (serialized to mirror the writer's attribute order), and hands the combined block to the merger as `SheetBlock::DataValidations`. Final test count: 881 pytest passed (+11), 28 diffwriter golden files unchanged.

## 1. Problem Statement

`python/wolfxl/worksheet/datavalidation.py:92-96` raises `NotImplementedError`
when a user calls `ws.data_validations.append(dv)` on a worksheet opened in
modify mode:

```python
if wb._rust_writer is None:
    raise NotImplementedError(
        "Appending data validations to existing files is a T1.5 follow-up. "
        "Write mode (Workbook() + save) is supported."
    )
```

The desired behavior: `ws.data_validations.append(dv)` works in modify mode.
The patcher reads the existing `<dataValidations>` block from the sheet XML (if
any), merges the new rules with the existing ones, and writes the combined block
back via RFC-011's block merger.

Data validations live entirely within the sheet XML - no separate ZIP parts, no
rels entries, no content-type changes. This makes the patcher path significantly
simpler than RFC-024 (tables).

Write mode already works: `python/wolfxl/worksheet/datavalidation.py:98` queues
new DVs onto `ws._pending_data_validations`, and the native writer in
`crates/wolfxl-writer/src/emit/sheet_xml.rs:698-806` serializes the
`<dataValidations>` block. The patcher path is the missing piece.

## 2. OOXML Spec Surface

ECMA-376 Part 1 §18.3.1.32 defines `CT_DataValidations`. The block is a direct
child of `CT_Worksheet`, positioned AFTER `<conditionalFormatting>` blocks and
BEFORE `<hyperlinks>`.

Example with two validations:

```xml
<dataValidations count="2">
  <dataValidation type="list" allowBlank="1" showErrorMessage="1"
                  sqref="B2:B100">
    <formula1>"Apple,Banana,Cherry"</formula1>
  </dataValidation>
  <dataValidation type="whole" operator="between" allowBlank="0"
                  showInputMessage="1" showErrorMessage="1"
                  errorTitle="Invalid" error="Enter 1-100"
                  promptTitle="Hint" prompt="Enter a number 1-100"
                  sqref="C2:C100">
    <formula1>1</formula1>
    <formula2>100</formula2>
  </dataValidation>
</dataValidations>
```

Key `<dataValidation>` attributes:

| Attribute | Values | Notes |
|-----------|--------|-------|
| `type` | none, whole, decimal, list, date, time, textLength, custom | Required |
| `operator` | between, notBetween, equal, notEqual, lessThan, lessThanOrEqual, greaterThan, greaterThanOrEqual | Omitted for list/custom |
| `allowBlank` | 0/1 | Default 0 |
| `showDropDown` | 0/1 | Note: capital D; 1 = HIDE dropdown (inverted) |
| `showInputMessage` | 0/1 | Default 0 |
| `showErrorMessage` | 0/1 | Default 0 |
| `errorStyle` | stop, warning, information | Default stop; omit when stop |
| `errorTitle`, `error` | strings | Error dialog title and body |
| `promptTitle`, `prompt` | strings | Input message title and body |
| `sqref` | space-separated A1 ranges | Required |

Children: `<formula1>` (required when type != "any"), `<formula2>` (for
between/notBetween operators only).

For `type="list"` with an inline value list, `formula1` contains a
double-quoted comma-separated string: `"Apple,Banana,Cherry"`. For a range
reference it contains an unquoted ref: `Sheet2!$A$1:$A$10`.

## 3. openpyxl Reference

Files:
- `.venv/lib/python3.14/site-packages/openpyxl/worksheet/datavalidation.py`
- `.venv/lib/python3.14/site-packages/openpyxl/descriptors/` (for field types)

Key behaviors:

- `DataValidation.__init__` defaults `type="none"` (no validation), `operator="between"`.
- `DataValidationList.append(dv)` also calls `dv.sqref = sqref` if the DV was constructed without one (e.g. `ws.data_validations.add("A1:A10", dv)`).
- openpyxl merges DVs by sqref when the same range appears twice: it consolidates rules rather than emitting duplicates.
- When reading, openpyxl parses `<dataValidations>` from the sheet XML into `DataValidation` objects. Those are accessible as `ws.data_validations.dataValidation` (the list).
- openpyxl handles `type="list"` with `formula1` that starts with `=` as a range reference and strips the leading `=` internally. wolfxl stores it as-is (following the keep-leading-`=` convention from `DataValidation.formula1`).
- The `sqref` attribute is a space-separated multi-range string. openpyxl stores it as `MultiCellRange`; wolfxl keeps it as a plain string.

What we will NOT copy:
- openpyxl's `DataValidationList.add(range, dv)` signature (our API already uses `.append(dv)` with `.sqref` on the DV object).
- Sqref consolidation logic. wolfxl appends each DV as a separate element even if the range overlaps an existing one. Excel handles overlapping DVs by applying the last one; this matches real-world usage where users add DVs programmatically without re-checking for overlap.

## 4. WolfXL Surface Area

### 4.1 Python coordinator

Files touched:

- `python/wolfxl/worksheet/datavalidation.py:92-96` - Remove the patcher guard. The `append()` method queues onto `ws._pending_data_validations` unconditionally (same as write mode path at line 98).

- `python/wolfxl/_workbook.py` (the `save()` method) - After flushing cell patches, call `flush_pending_data_validations(ws, pending)` for each sheet with non-empty `_pending_data_validations`. This is analogous to the existing write-mode DV flush.

### 4.2 Patcher (modify mode)

New module: `src/wolfxl/validations.rs`

Public Rust API:

```rust
pub struct DataValidationPatch {
    pub validation_type: String,         // "list", "whole", "decimal", etc.
    pub operator: Option<String>,        // "between", "greaterThan", etc.
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    pub sqref: String,                   // space-separated multi-range
    pub allow_blank: bool,
    pub show_dropdown: bool,             // note: 1 = HIDE in OOXML (inverted)
    pub show_input_message: bool,
    pub show_error_message: bool,
    pub error_style: Option<String>,     // "warning", "information"; None = stop
    pub error_title: Option<String>,
    pub error: Option<String>,
    pub prompt_title: Option<String>,
    pub prompt: Option<String>,
}

/// Parse the existing <dataValidations> block from sheet XML.
///
/// Returns the raw XML bytes of the existing block (everything from
/// `<dataValidations` through `</dataValidations>`), or None if no block
/// exists in the sheet. Used by the merger to preserve existing rules.
pub fn extract_existing_dv_block(sheet_xml: &str) -> Option<Vec<u8>>;

/// Build a complete <dataValidations> block from existing + new patches.
///
/// If `existing_block` is Some, parse the existing <dataValidation> children
/// and prepend them to the new patches (existing rules come first; new ones
/// are appended). The returned bytes are the full
/// `<dataValidations count="N">...</dataValidations>` XML, ready to hand to
/// RFC-011's block merger as SheetBlock::DataValidations.
pub fn build_data_validations_block(
    existing_block: Option<&[u8]>,
    patches: &[DataValidationPatch],
) -> Vec<u8>;
```

The serialization logic mirrors `crates/wolfxl-writer/src/emit/sheet_xml.rs:698-806`
(`emit_data_validations`). The new module can import and reuse that function's
logic or inline a simplified version that operates on `DataValidationPatch`
rather than the native writer's `DataValidation` model struct.

### 4.3 Native writer (write mode)

No changes needed. Write mode already serializes DVs via
`crates/wolfxl-writer/src/emit/sheet_xml.rs:698-806`. The patcher module
independently builds the same XML structure for modify mode.

## 5. Algorithm

```
modify_mode_add_data_validations(sheet_xml, patches):

  # 1. Extract existing <dataValidations> block from the sheet XML (if any).
  existing_block = extract_existing_dv_block(sheet_xml)

  # 2. Build combined block: existing DVs first, then new patches appended.
  combined_block = build_data_validations_block(existing_block, patches)

  # 3. Hand the block to RFC-011's merger with placement tag DataValidations.
  #    RFC-011 replaces the existing <dataValidations> if present, or inserts
  #    after <conditionalFormatting> (or after <mergeCells>, or before
  #    <hyperlinks>) if not.
  patched_sheet_xml = block_merger(sheet_xml, SheetBlock::DataValidations(combined_block))

  return patched_sheet_xml
```

**Merge strategy — Option B (append, preserve existing)**: The patcher READS
the existing `<dataValidations>` block, parses out the individual
`<dataValidation>` children as opaque XML strings, then builds a new block that
contains: (1) existing rules verbatim, (2) new patches serialized fresh. This
matches openpyxl's `load_workbook` + `append` + `save` workflow where existing
rules are preserved. Option A (replace entire block) is explicitly rejected
because it would silently discard rules that users didn't add via wolfxl (e.g.
rules from Excel or another tool).

**Parsing existing block**: The `extract_existing_dv_block` and subsequent child
extraction uses quick-xml's streaming reader. Each `<dataValidation ...>...</dataValidation>`
element (including its children `<formula1>` / `<formula2>`) is captured as a raw
byte range and re-emitted verbatim into the combined block. This avoids
round-tripping through a parsed model, which could silently alter formatting or
lose attributes wolfxl doesn't know about.

**Sqref overlap**: Two DVs can target overlapping ranges. wolfxl emits them as
separate `<dataValidation>` elements; Excel applies the last matching rule when
ranges overlap. No de-duplication is attempted.

**`showDropDown` inversion**: OOXML's `showDropDown="1"` means HIDE the dropdown
(counterintuitive). The `DataValidationPatch.show_dropdown` field maps to the
OOXML attribute directly (1 = hide). The Python `DataValidation.showDropDown`
field is also a direct passthrough (following openpyxl's convention). Document
this inversion clearly in the Python docstring.

## 6. Test Plan

| Test | What it checks |
|------|---------------|
| `test_add_dv_to_clean_file` | File with no existing DVs. After save, sheet XML contains `<dataValidations count="1">` with correct attributes. |
| `test_add_dv_preserves_existing` | File with one existing list DV. Add a whole-number DV. After save, `count="2"`, both DVs present, existing one unchanged. |
| `test_dv_list_inline_values` | `type="list"`, `formula1='"A,B,C"'`. Verify `<formula1>"A,B,C"</formula1>` in output. |
| `test_dv_list_range_ref` | `type="list"`, `formula1="Sheet2!$A$1:$A$5"`. Verify formula1 contains the range ref verbatim. |
| `test_dv_whole_between` | `type="whole"`, `operator="between"`, formula1=1, formula2=100. Verify `<formula2>100</formula2>` present. |
| `test_dv_custom_type` | `type="custom"`, `formula1="=LEN(A1)>5"`. Verify operator attribute is omitted. |
| `test_dv_error_style_warning` | `errorStyle` appears as `"warning"` in output; default `stop` is omitted. |
| `test_dv_multi_sqref` | `sqref="A1:A10 C1:C10"` (space-separated). Verify sqref attribute passed through unchanged. |
| `test_roundtrip_libreoffice` | LibreOffice opens the patched file; validation rules trigger correctly. |

LibreOffice note: `type="custom"` with formulas that reference defined names or
cross-sheet ranges may silently fail validation in LibreOffice without erroring
on open. Document in `KNOWN_GAPS.md` if encountered.

## 7. Migration / Compat Notes

- **Before this RFC**: `ws.data_validations.append(dv)` in modify mode raises `NotImplementedError`. After: it works.
- **Backward compat**: Write mode behavior is unchanged.
- **openpyxl divergence**: openpyxl's `DataValidationList.add(range, rule)` method takes the range as a separate argument and sets `dv.sqref` automatically. wolfxl's API requires `dv.sqref` to be set before calling `.append(dv)`. This divergence pre-exists this RFC; no change needed here.

## 8. Risks & Open Questions

1. **(MED) Existing block byte-extraction accuracy**: The `extract_existing_dv_block` function must handle edge cases: empty `<dataValidations/>` (self-closing), nested elements with escaped content, and `<dataValidation>` elements with `<formula1>` containing `<` or `&`. Using quick-xml's event-based reader and capturing the raw byte slice between start/end events handles these correctly without manual string scanning.

2. **(LOW) `count` attribute accuracy**: The combined block's `count` attribute must equal the total number of `<dataValidation>` children (existing + new). Miscount causes Excel to silently drop validations. The implementation must count precisely rather than relying on the existing block's `count` attribute value.

3. **(LOW) LibreOffice rejects `type="custom"` formulas it doesn't understand**: LibreOffice may silently ignore or reject `<dataValidation type="custom">` elements with formulas that use Excel-specific functions. This is a known gap between Excel and LibreOffice and is out of scope for this RFC. Document in `tests/parity/KNOWN_GAPS.md` if it surfaces in testing.

4. **(LOW) Position of `<dataValidations>` in CT_Worksheet**: RFC-011's merger must place the block after `<conditionalFormatting>` and before `<hyperlinks>`. This RFC supplies the bytes; RFC-011 owns the placement.

## 9. Effort Breakdown

| Task | LOC est. | Days |
|------|----------|------|
| `src/wolfxl/validations.rs` — extract + build functions | ~200 | 1.5 |
| Python coordinator wiring in `_workbook.py` + remove guard at `datavalidation.py:92` | ~40 | 0.5 |
| RFC-011 `SheetBlock::DataValidations` support (if not already landed) | ~60 | 0.5 |
| Tests (Rust unit + pytest integration) | ~200 | 1.0 |
| **Total** | **~500** | **3.5** |

## 10. Out of Scope

- Reading data validations in modify mode (already works via CalamineStyledBook).
- Deleting or modifying existing data validations (no use case yet).
- `type="none"` (any value) validations are supported as a passthrough but are a no-op in Excel.
- IME mode attributes (`imeMode`) on `<dataValidation>`.
- Sqref de-duplication / consolidation.
