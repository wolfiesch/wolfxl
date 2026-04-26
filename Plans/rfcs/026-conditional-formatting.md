# RFC-026: Sheet-Scoped Conditional Formatting in Modify Mode

Status: Researched
Owner: pod-P4
Phase: 3
Estimate: M
Depends-on: RFC-011
Unblocks: RFC-030, RFC-031, RFC-035

## 1. Problem Statement

`python/wolfxl/formatting/__init__.py:72-76` raises `NotImplementedError` when a
user calls `ws.conditional_formatting.add(range, rule)` on a worksheet opened
in modify mode:

```python
if wb._rust_writer is None:
    raise NotImplementedError(
        "Adding conditional formatting rules to existing files is a T1.5 follow-up. "
        "Write mode (Workbook() + save) is supported."
    )
```

The desired behavior: `ws.conditional_formatting.add(range, rule)` works in
modify mode. The patcher reads the existing `<conditionalFormatting>` blocks
from the sheet XML (if any), appends the new blocks, updates the `<dxfs>`
collection in `xl/styles.xml` when the new rules carry formatting, and writes
both the sheet XML and styles XML back into the ZIP.

Conditional formatting lives in the sheet XML (the `<conditionalFormatting>`
blocks) with a dependency on `xl/styles.xml` for rules that reference
differential formats (`dxfId`). This makes CF the most complex of the three
sheet-collection RFCs, but still simpler than tables because it requires no new
ZIP parts or rels entries.

Write mode already works: `python/wolfxl/formatting/__init__.py:84` queues CF
onto `ws._pending_conditional_formats`, and the native writer in
`crates/wolfxl-writer/src/emit/sheet_xml.rs:523-696` serializes the
`<conditionalFormatting>` blocks. The patcher path is the missing piece.

## 2. OOXML Spec Surface

ECMA-376 Part 1 §18.3.1.18 defines `CT_ConditionalFormatting` and §18.3.1.10
defines `CT_CfRule`. The blocks are direct children of `CT_Worksheet`,
positioned BEFORE `<dataValidations>` and after `<mergeCells>`.

A typical CF block with a cellIs rule:

```xml
<conditionalFormatting sqref="A1:A10">
  <cfRule type="cellIs" priority="1" operator="greaterThan" dxfId="0">
    <formula>5</formula>
  </cfRule>
</conditionalFormatting>
```

A colorScale rule (no dxfId - these rules carry their own color inline):

```xml
<conditionalFormatting sqref="B1:B20">
  <cfRule type="colorScale" priority="2">
    <colorScale>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FFF8696B"/>
      <color rgb="FF63BE7B"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>
```

A dataBar rule:

```xml
<conditionalFormatting sqref="C1:C20">
  <cfRule type="dataBar" priority="3">
    <dataBar>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FF638EC6"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>
```

Key `<cfRule>` attributes:

| Attribute | Notes |
|-----------|-------|
| `type` | cellIs, expression, colorScale, dataBar, iconSet, top10, uniqueValues, duplicateValues, containsText, notContainsText, beginsWith, endsWith, containsBlanks, notContainsBlanks, containsErrors, notContainsErrors, timePeriod, aboveAverage |
| `priority` | Positive integer; lower = higher priority. Sheet-wide ordering. |
| `dxfId` | Index into `<dxfs>` in `xl/styles.xml`. Required for rules that apply cell formatting (cellIs, expression, text-based). Absent for colorScale, dataBar, iconSet. |
| `operator` | For cellIs: equal, notEqual, lessThan, lessThanOrEqual, greaterThan, greaterThanOrEqual, between, notBetween |
| `stopIfTrue` | 0/1; stops evaluating lower-priority rules when this one matches |

**Critical**: `dxfId` is a 0-based index into the `<dxfs>` element of
`xl/styles.xml`. Adding a new rule with formatting requires appending a new
`<dxf>` entry to `<dxfs>` in styles.xml and using the resulting index as
`dxfId`. Rules without formatting (colorScale, dataBar, iconSet) set no
`dxfId`.

**Priority numbering**: ECMA-376 requires that priority values are unique
across ALL `<cfRule>` elements in a sheet (not just within one
`<conditionalFormatting>` block). Priority 1 wins over priority 2.

## 3. openpyxl Reference

Files:
- `.venv/lib/python3.14/site-packages/openpyxl/formatting/rule.py`
- `.venv/lib/python3.14/site-packages/openpyxl/formatting/formatting.py`
- `.venv/lib/python3.14/site-packages/openpyxl/styles/differential.py`

Key behaviors:

- `openpyxl/formatting/rule.py` defines `Rule` (base), `ColorScaleRule`,
  `DataBarRule`, `IconSetRule`, `FormulaRule`, and `CellIsRule`. Each is a
  `Serialisable` descriptor class.
- openpyxl tracks a `ConditionalFormattingList` at `ws.conditional_formatting`.
  Calling `ws.conditional_formatting.add("A1:A10", rule)` adds to an internal
  `_cf_rules` dict keyed by sqref.
- openpyxl serializes CF rules in insertion order within each sqref group.
- Priority assignment: openpyxl numbers priorities starting at 1 and increments
  globally across all rules on the sheet.
- `DifferentialStyle` (in `styles/differential.py`) carries `font`, `fill`,
  `border`, `number_format` — the same components as a regular cell style but
  only the fields the CF rule overrides. This maps to a `<dxf>` element.
- For colorScale / dataBar / iconSet, openpyxl sets `dxf=None` and omits the
  `dxfId` attribute.
- openpyxl reads existing CF rules from the sheet XML into its model on
  `load_workbook`; calling `.add()` extends the in-memory list. On save, the
  full merged list (old + new) is serialized.

What we will NOT copy:
- openpyxl's descriptor-based `Serialisable` system. The wolfxl `Rule` class at
  `python/wolfxl/formatting/rule.py` is a simpler dataclass; we extend it as
  needed.
- openpyxl's `timePeriod` rule type. It requires special date logic; stub it as
  passthrough-only (preserve existing, refuse to add new via the patcher).

## 4. WolfXL Surface Area

### 4.1 Python coordinator

Files touched:

- `python/wolfxl/formatting/__init__.py:72-76` - Remove the patcher guard. The
  `add()` method queues onto `ws._pending_conditional_formats` unconditionally.

- `python/wolfxl/_workbook.py` (the `save()` method) - After flushing cell
  patches but before closing the ZIP, call
  `flush_pending_conditional_formats(ws, pending)` for each sheet with
  non-empty `_pending_conditional_formats`. This requires access to styles.xml
  (to allocate dxf IDs and update the dxfs section).

- `python/wolfxl/formatting/rule.py` - The `Rule` dataclass may need a
  `DifferentialStyle` field (or equivalent) to carry font/fill/border
  information for rules that set cell formatting. Currently the `Rule` class
  stores the styling via a `dxf` field; verify the field is present and
  accessible via PyO3.

### 4.2 Patcher (modify mode)

New module: `src/wolfxl/conditional_formatting.rs`

Public Rust API:

```rust
/// A differential format (font/fill/border override) for a CF rule.
/// Only the fields being overridden need to be populated.
pub struct DxfPatch {
    pub font_bold: Option<bool>,
    pub font_italic: Option<bool>,
    pub font_color_rgb: Option<String>,      // "FFRRGGBB"
    pub fill_pattern_type: Option<String>,   // "solid", etc.
    pub fill_fg_color_rgb: Option<String>,   // "FFRRGGBB"
    pub border_top_style: Option<String>,
    pub border_bottom_style: Option<String>,
    pub border_left_style: Option<String>,
    pub border_right_style: Option<String>,
}

/// A cfvo threshold for colorScale / dataBar / iconSet.
pub struct CfvoPatch {
    pub cfvo_type: String,  // "min", "max", "num", "percent", "percentile", "formula"
    pub val: Option<String>,
}

/// A color-scale stop.
pub struct ColorScaleStop {
    pub cfvo: CfvoPatch,
    pub color_rgb: String,
}

pub enum CfRuleKind {
    CellIs {
        operator: String,
        formula_a: String,
        formula_b: Option<String>,
    },
    Expression {
        formula: String,
    },
    ColorScale {
        stops: Vec<ColorScaleStop>,
    },
    DataBar {
        min: CfvoPatch,
        max: CfvoPatch,
        color_rgb: String,
    },
}

pub struct CfRulePatch {
    pub kind: CfRuleKind,
    /// Differential format for rules that set cell styling. None for
    /// colorScale / dataBar (those carry their own color inline).
    pub dxf: Option<DxfPatch>,
    pub stop_if_true: bool,
}

pub struct ConditionalFormattingPatch {
    pub sqref: String,              // space-separated multi-range
    pub rules: Vec<CfRulePatch>,
}

/// Output of the CF builder.
pub struct CfResult {
    /// The serialized XML blocks: one <conditionalFormatting sqref="...">
    /// per patch, concatenated. Hand directly to RFC-011's block merger
    /// as SheetBlock::ConditionalFormattings (appended after existing blocks).
    pub block_bytes: Vec<u8>,
    /// New <dxf> entries to append to xl/styles.xml's <dxfs> collection.
    /// Ordered to match the dxfId values already baked into block_bytes.
    pub new_dxfs: Vec<DxfPatch>,
}

/// Build the CF result for one worksheet's pending patches.
///
/// `existing_priority_max` is the highest priority value already present
/// in the sheet's existing CF rules. The new rules are numbered starting
/// at `existing_priority_max + 1`. The caller obtains this value by
/// scanning the existing sheet XML for the highest `priority` attribute
/// across all <cfRule> elements.
///
/// `existing_dxf_count` is the current count of <dxf> entries in
/// xl/styles.xml's <dxfs> section. New dxfId values start at this count.
pub fn build_cf_blocks(
    patches: &[ConditionalFormattingPatch],
    existing_priority_max: u32,
    existing_dxf_count: u32,
) -> CfResult;

/// Scan a sheet XML string and return the maximum priority value found
/// across all <cfRule> elements. Returns 0 if no CF rules exist.
pub fn scan_max_cf_priority(sheet_xml: &str) -> u32;

/// Count the number of <dxf> entries in xl/styles.xml.
/// Used to determine the starting dxfId for new rules.
pub fn count_dxfs(styles_xml: &str) -> u32;
```

**DXF seam to `src/wolfxl/styles.rs`**: The existing `styles.rs` module exposes
`inject_into_section(xml, "dxfs", dxf_xml)` which can be used directly to append
a `<dxf>` entry. However, `styles.rs` currently has no `dxfs` serializer (it
handles fonts, fills, borders, cellXfs, not differential formats). This RFC adds
a `dxf_to_xml(patch: &DxfPatch) -> String` function in
`src/wolfxl/conditional_formatting.rs` and calls `inject_into_section` from
`styles.rs` to append it. If `<dxfs>` does not exist in the styles file (common
in simple workbooks), the code must create the section before injecting.

See §5 for the creation logic when `<dxfs>` is absent.

### 4.3 Native writer (write mode)

No changes needed. Write mode serializes CF via
`crates/wolfxl-writer/src/emit/sheet_xml.rs:523-696`. The patcher module
builds the same XML structure for modify mode with the additional concern of
priority and dxfId allocation relative to existing content.

Note that the native writer currently stubs several CF rule types (ContainsText,
NotContainsText, BeginsWith, EndsWith, Duplicate, Unique, Top10, AboveAverage,
IconSet) with a dropped-rule warning. The patcher will ALSO stub these types
initially; patcher expansion follows the writer expansion wave.

## 5. Algorithm

```
modify_mode_add_cf(sheet_xml, styles_xml, patches):

  # 1. Scan existing sheet XML for max priority and existing sqrefs.
  priority_max = scan_max_cf_priority(sheet_xml)
  # (sqref overlap is not checked; overlapping rules co-exist)

  # 2. Count existing dxfs in styles.xml.
  dxf_count = count_dxfs(styles_xml)

  # 3. Build the new CF blocks and required dxf entries.
  result = build_cf_blocks(patches, priority_max, dxf_count)

  # 4. Update styles.xml: append each new dxf.
  #    If <dxfs> section is absent, create it.
  for dxf in result.new_dxfs:
      dxf_xml = dxf_to_xml(dxf)
      if styles_xml.contains("<dxfs"):
          (styles_xml, _) = inject_into_section(styles_xml, "dxfs", dxf_xml)
      else:
          # Insert <dxfs count="N">...</dxfs> before </styleSheet>
          styles_xml = create_dxfs_section(styles_xml, [dxf])

  # 5. Patch sheet XML.
  #    NOTE: As shipped, RFC-011's merger uses REPLACE-ALL CF semantics
  #    (see crates/wolfxl-merger/src/lib.rs Q4): supplying any
  #    SheetBlock::ConditionalFormatting drops every existing
  #    <conditionalFormatting> element from the source. The patcher
  #    therefore captures the source's existing blocks via
  #    extract_existing_cf_blocks and re-emits them VERBATIM at the head
  #    of result.block_bytes. The *effect* is the same as the original
  #    "leave existing untouched" wording — existing priorities and
  #    dxfIds remain correct relative to each other — but the
  #    *mechanism* is byte-slice capture-and-re-include in the patcher,
  #    not opt-out at the merger.
  patched_sheet_xml = block_merger(
      sheet_xml,
      SheetBlock::ConditionalFormatting(result.block_bytes)
      # result.block_bytes = existing_blocks_verbatim || new_blocks
  )

  return (patched_sheet_xml, styles_xml)
```

**Priority allocation**: New rules are numbered `existing_priority_max + 1`,
`existing_priority_max + 2`, etc. Within a single `ConditionalFormattingPatch`,
rules are numbered in order. Across multiple patches, numbering is sequential
in the order `patches` is passed. This ensures uniqueness across the sheet.

**dxfId allocation**: New dxf entries start at `existing_dxf_count`. The first
new rule with a `dxf` gets `dxfId = existing_dxf_count`, the second gets
`existing_dxf_count + 1`, etc. Rules without a dxf (colorScale, dataBar) skip
the counter.

**Existing CF blocks are preserved verbatim**: Unlike DVs (which parse existing
elements into a model), the CF patcher appends new blocks WITHOUT touching
existing ones. This is safe because:
- Existing priority values remain valid (new rules are numbered above them).
- Existing dxfId values remain valid (new dxfs are appended after existing ones).
- The sqref of existing blocks is not modified.

**Creating `<dxfs>` when absent**: Simple workbooks (no conditional formatting)
do not have a `<dxfs>` section in styles.xml. The implementation inserts
`<dxfs count="N">...</dxfs>` immediately before `</styleSheet>`. After
insertion, subsequent calls use `inject_into_section` as usual.

**Partial rule support - stubbed kinds**: ContainsText, NotContainsText,
BeginsWith, EndsWith, Duplicate, Unique, Top10, AboveAverage, and IconSet are
not serialized in this RFC (consistent with the native writer's current state).
Calling `.add()` with one of these rule types raises `NotImplementedError` at
the Python layer with a message pointing to the expansion wave.

## 6. Test Plan

| Test | What it checks |
|------|---------------|
| `test_add_cellis_rule_to_clean_file` | File with no existing CF. After save, sheet XML contains one `<conditionalFormatting>` block with `priority="1"`, styles.xml has `<dxfs count="1">`. |
| `test_add_cf_preserves_existing` | File with one existing CF rule (`priority="1"`, `dxfId="0"`). Add new rule. New rule gets `priority="2"`, `dxfId="1"`. Existing rule unchanged. |
| `test_dxf_id_monotonic_across_rules` | Two new CellIs rules in the same patch. First gets `dxfId=N`, second gets `dxfId=N+1`. |
| `test_priority_monotonic_across_sqrefs` | Two patches with different sqrefs. Second patch's rules have higher priority numbers than first patch's rules. |
| `test_colorscale_no_dxf` | ColorScale rule: no `dxfId` attribute on `<cfRule>`, styles.xml's `<dxfs>` unchanged. |
| `test_databar_no_dxf` | DataBar rule: same pattern as colorScale. |
| `test_expression_rule` | `type="expression"`, `formula="A1>B1"`. Verify `<formula>A1&gt;B1</formula>` in output (XML escaping). |
| `test_cellis_between` | `operator="between"`, two formulas. Verify both `<formula>` children present. |
| `test_create_dxfs_section_when_absent` | Source file has no `<dxfs>` in styles.xml. After adding a CellIs rule, `<dxfs count="1">` is present. |
| `test_roundtrip_libreoffice` | LibreOffice opens the patched file; CF rules visually apply to correct ranges. |

## 7. Migration / Compat Notes

- **Before this RFC**: `ws.conditional_formatting.add(range, rule)` in modify mode raises `NotImplementedError`. After: works for CellIs, Expression, ColorScale, DataBar.
- **Stubbed rule types**: ContainsText, BeginsWith, EndsWith, Duplicate, Unique, Top10, AboveAverage, IconSet continue to raise `NotImplementedError` in both modify mode (this RFC) and write mode (pending native writer expansion). The error message is updated to cite the CF expansion wave rather than T1.5.
- **Backward compat**: Write mode behavior is unchanged.
- **dxfId gap risk**: If the source file's `<dxfs>` count attribute is out of sync with the actual number of `<dxf>` elements (which Excel has been known to produce), `count_dxfs` must count actual child elements rather than trusting the `count` attribute. The implementation uses element counting.

## 8. Risks & Open Questions

1. **(HIGH) dxfId collision with existing rules**: If the source file's `<dxfs>` section already has N entries, new rules must start at dxfId=N. An off-by-one error here silently applies the wrong formatting. The `count_dxfs` function must count actual `<dxf>` child elements (not trust the `count` attribute) because some Excel versions emit incorrect count values. Resolution: scan and count children in `count_dxfs`.

2. **(HIGH) Priority collision with existing rules**: If the scanner misses an existing `<cfRule priority>` value, new rules could get duplicate priorities. Duplicate priorities cause Excel to apply rules non-deterministically. Resolution: `scan_max_cf_priority` uses quick-xml to read all `priority` attributes across all `<cfRule>` elements in the sheet and returns the maximum. This is O(sheet_xml.len()) but acceptable.

3. **(MED) styles.xml mutation synchronization**: The patcher reads styles.xml from the ZIP once, applies all dxf additions for all sheets, and writes it back once. If multiple sheets add CF rules in the same `save()` call, the running styles.xml string must be threaded through all sheet CF flushes. The coordinator in `_workbook.py` must handle this by accumulating dxf additions across sheets before the final write.

4. **(MED) Does RFC-026 require a new `styles.rs` extension?** Re-evaluation: `src/wolfxl/styles.rs` already exposes `inject_into_section(xml, section_tag, new_element)` which is exactly what's needed to append a `<dxf>` to `<dxfs>`. The only missing piece is (a) a `dxf_to_xml(DxfPatch)` serializer and (b) the `<dxfs>` section creation fallback when the section is absent. Both of these are small additions (< 50 LOC) that belong in `src/wolfxl/conditional_formatting.rs` rather than `styles.rs`. **Conclusion: no new public API needed in `styles.rs`; reuse `inject_into_section` as-is.**

5. **(LOW) Empty `<conditionalFormatting>` blocks are invalid**: Excel will "repair" a file containing `<conditionalFormatting sqref="A1:A10"></conditionalFormatting>` with no `<cfRule>` children. The native writer already guards against this (see `sheet_xml.rs:541-544`). The patcher must apply the same guard: skip emission of a `<conditionalFormatting>` wrapper if the rules buffer is empty after handling stubbed kinds.

## 9. Effort Breakdown

| Task | LOC est. | Days |
|------|----------|------|
| `src/wolfxl/conditional_formatting.rs` — build_cf_blocks, dxf_to_xml, scan helpers | ~300 | 2.0 |
| `styles.rs` additions: `<dxfs>` section creation fallback | ~50 | 0.25 |
| Python coordinator wiring in `_workbook.py` + remove guard at `formatting/__init__.py:72` | ~60 | 0.5 |
| RFC-011 `SheetBlock::ConditionalFormattings` support (append semantics) | ~60 | 0.5 |
| Tests (Rust unit + pytest integration) | ~250 | 1.0 |
| **Total** | **~720** | **4.25** |

## 10. Out of Scope

- Pivot table CF (uses a different mechanism, separate XML namespace extLst).
- Sparkline CF (in `extLst`, not supported by the current model).
- ContainsText, BeginsWith, EndsWith, Duplicate, Unique, Top10, AboveAverage, IconSet rule types (pending CF expansion wave after native writer ships these).
- Modifying or deleting existing CF rules.
- Reading CF rules in modify mode (already works via CalamineStyledBook reader returning `ConditionalFormatting` objects).
- `timePeriod` rule type (requires date serial logic).

## Acceptance: shipped via 5-commit slice on `feat/native-writer` (commits `3e40530..<this commit>`) on 2026-04-25

- Activates `ws.conditional_formatting.add(range, rule)` in modify mode for `cellIs`, `expression`, `colorScale`, `dataBar`. Stubbed kinds (ContainsText, IconSet, Top10, AboveAverage, etc.) raise `NotImplementedError` with a §10 pointer.
- Phase-2.5b in `XlsxPatcher::do_save` walks queued CF patches in deterministic sorted-sheet-name order, threading a workbook-wide `running_dxf_count` so cross-sheet `dxfId` allocations are reproducible and globally unique. One combined `xl/styles.xml` mutation per save.
- Replace-all merger semantics handled in the patcher: `extract_existing_cf_blocks` captures source byte ranges, re-emitted verbatim at the head of `build_cf_blocks` output (see §5 mechanism clarification above).
- Headline gates (in `tests/test_modify_conditional_formatting.py`, all green): `test_add_cf_preserves_existing` (byte-preservation under replace-all) and `test_dxf_id_monotonic_across_sheets` (cross-sheet allocator). Plus 11 RFC-§6 cases. No-op save (`test_cf_no_pending_no_op`) byte-identical via the extended `do_save` short-circuit predicate.
- 14 Rust unit tests in `src/wolfxl/conditional_formatting.rs` cover scans, byte-range capture, dxfId/priority allocation, and `<dxfs>` section creation when absent.
- `RFC-011`'s `SheetBlock::ConditionalFormatting` variant is no longer dead code — joins `DataValidations` (RFC-025) as a live patcher block kind.
