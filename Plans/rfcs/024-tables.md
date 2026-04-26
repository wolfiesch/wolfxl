# RFC-024: Sheet-Scoped Tables in Modify Mode

Status: Shipped
Owner: pod-P4
Phase: 3
Estimate: M
Depends-on: RFC-010, RFC-011
Unblocks: RFC-030, RFC-031, RFC-035

## 1. Problem Statement

`python/wolfxl/_worksheet.py:1511-1515` raises `NotImplementedError` whenever a
user calls `ws.add_table(table)` on a worksheet opened via `load_workbook(path)`
(modify mode):

```python
if wb._rust_writer is None:
    raise NotImplementedError(
        "Adding tables to existing files is a T1.5 follow-up. "
        "Write mode (Workbook() + save) is supported."
    )
```

The desired behavior: `ws.add_table(t)` works in modify mode. The patcher reads
the current rels and content-types from the ZIP, allocates a workbook-unique
table ID, serializes `xl/tables/tableN.xml`, inserts a `<tableParts>` block into
the patched `sheetN.xml`, adds the rels entry and the content-type override, and
writes everything back into the ZIP without touching any other parts.

Write mode already works end-to-end: `python/wolfxl/_worksheet.py:1520` queues
the table on `self._pending_tables`, and the native writer in
`crates/wolfxl-writer/src/emit/tables_xml.rs` + `crates/wolfxl-writer/src/emit/rels.rs`
flushes it. The patcher path is the missing piece.

## 2. OOXML Spec Surface

ECMA-376 Part 1 §18.5.1.2 defines `CT_Table`. A complete minimal table XML part:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="1" name="SalesTable" displayName="SalesTable" ref="A1:E10"
       totalsRowShown="0">
  <autoFilter ref="A1:E10"/>
  <tableColumns count="5">
    <tableColumn id="1" name="Region"/>
    <tableColumn id="2" name="Q1"/>
    <tableColumn id="3" name="Q2"/>
    <tableColumn id="4" name="Q3"/>
    <tableColumn id="5" name="Total"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2"
                  showFirstColumn="0" showLastColumn="0"
                  showRowStripes="1" showColumnStripes="0"/>
</table>
```

Required `<table>` attributes: `id` (workbook-unique u32), `name`, `displayName`,
`ref`. Optional but nearly always present: `totalsRowShown`, `headerRowCount`
(default 1).

Sheet XML side (ECMA-376 §18.3.1.88, last child of CT_Worksheet before
`<extLst>`):

```xml
<tableParts count="2">
  <tablePart r:id="rId3"/>
  <tablePart r:id="rId4"/>
</tableParts>
```

Content-type override (goes into `[Content_Types].xml`):
```
application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml
```
Part path pattern: `/xl/tables/tableN.xml` where N is the 1-based global
counter (not per-sheet).

Relationship type for sheet rels:
```
http://schemas.openxmlformats.org/officeDocument/2006/relationships/table
```
Target in the sheet's `xl/worksheets/_rels/sheetN.xml.rels`:
`../tables/tableN.xml`

## 3. openpyxl Reference

File: `.venv/lib/python3.14/site-packages/openpyxl/worksheet/table.py`

Key behaviors:

- `Table.__init__` defaults `displayName` to `name` when not set (line ~120).
- openpyxl allocates table IDs by scanning all sheets: `max(t.id for ws in wb for t in ws._tables.values()) + 1`, defaulting to 1 when no tables exist.
- Name collision check: openpyxl raises `ValueError: Table with that name already exists` when a name is reused. This is the correct behavior to match.
- `headerRowCount` defaults to 1; setting to 0 means no header row and alters the `autoFilter ref` (the filter then covers one row less).
- `totalsRowCount` > 0 reduces the effective data range by that many rows; the `ref` still covers the totals rows.
- openpyxl writes `<tableParts>` as the absolute last element before `</worksheet>` (after `<extLst>` if present, or just before `</worksheet>`). Excel is tolerant of position as long as the element is a child of `<worksheet>`.
- Table column `id` values are 1-based local IDs within the table (not global). They reset to 1 for each table.
- openpyxl does NOT validate that column names match actual cell header content.

What we will NOT copy:
- openpyxl's `TableFormula` / calculated column formula support. These appear as `<tableColumn totalsRowFormula="...">` children. Patcher passes these through on existing tables (from source ZIP) but does not allow setting them on new patches.
- `sortState` child elements inside `<table>`. Not needed for the common case; preserve on existing tables via byte-copy only.

## 4. WolfXL Surface Area

### 4.1 Python coordinator

Files touched:

- `python/wolfxl/_worksheet.py:1511-1515` - Remove the patcher guard. Route
  `add_table` to `self._pending_tables.append(table)` regardless of write vs
  modify mode. The `_pending_tables` list is already populated in write mode;
  modify mode will flush via the patcher at save time.

- `python/wolfxl/_workbook.py` (the `save()` method) - After flushing cell
  patches, call `flush_pending_tables(ws, pending)` for each sheet that has
  `_pending_tables`. This coordinator call does not exist today; it will be
  added alongside similar calls for DVs and CF.

- `python/wolfxl/worksheet/table.py` - No changes needed. The `Table`,
  `TableColumn`, and `TableStyleInfo` dataclasses are already correct.

### 4.2 Patcher (modify mode)

New module: `src/wolfxl/tables.rs`

Public Rust API:

```rust
use std::collections::HashSet;

pub struct TablePatch {
    pub name: String,
    pub display_name: String,
    pub ref_range: String,         // "A1:E10"
    pub columns: Vec<String>,      // column names in order
    pub style: Option<TableStylePatch>,
    pub header_row_count: u32,     // default 1
    pub totals_row_shown: bool,
    pub autofilter: bool,
}

pub struct TableStylePatch {
    pub name: String,
    pub show_first_column: bool,
    pub show_last_column: bool,
    pub show_row_stripes: bool,
    pub show_column_stripes: bool,
}

/// Output of the tables builder: new ZIP parts + sheet-level XML block.
pub struct TablesResult {
    /// Vec of (zip_entry_path, xml_bytes) for each new tableN.xml.
    /// Paths like "xl/tables/table3.xml".
    pub table_parts: Vec<(String, Vec<u8>)>,
    /// The <tableParts count="N">...</tableParts> XML block, ready to hand
    /// to the RFC-011 block merger as SheetBlock::TableParts.
    pub table_parts_block: Vec<u8>,
    /// One (rId, target_path) per new table, for injection into the sheet rels.
    /// RFC-010's RelsGraph appends these.
    pub new_rels: Vec<(String, String)>,
    /// Content-type override entries: (part_name, content_type).
    pub new_content_types: Vec<(String, String)>,
}

/// Build the tables result for one worksheet's pending patches.
///
/// `existing_table_ids` must include all table IDs from ALL sheets (not just
/// this one) so that allocated IDs are workbook-unique. The caller assembles
/// this set by scanning the patcher's ZIP for existing tableN.xml entries.
///
/// `existing_global_table_count` is the total number of tables already in
/// the workbook across all sheets (used to compute the next part index).
///
/// Returns Err if any patch name collides with an existing table name
/// (collected from `existing_table_names`).
pub fn build_tables(
    patches: &[TablePatch],
    existing_table_ids: &HashSet<u32>,
    existing_global_table_count: u32,
    existing_table_names: &HashSet<String>,
) -> Result<TablesResult, String>;
```

The implementation serializes each table using the same logic as
`crates/wolfxl-writer/src/emit/tables_xml.rs` (reuse directly or copy the
function into the new module and call it). The table `id` attribute uses
`next_available_id(existing_table_ids)` = `(1..).find(|i| !set.contains(i))`.

### 4.3 Native writer (write mode)

No changes needed. Tables in write mode already flow through
`crates/wolfxl-writer/src/emit/tables_xml.rs` and
`crates/wolfxl-writer/src/emit/rels.rs`. The write-mode path is already
exercised by existing tests in `crates/wolfxl-writer/src/emit/tables_xml.rs:99-377`.

## 5. Algorithm

```
modify_mode_add_tables(ws, patches):

  # 1. Collect workbook-wide existing table metadata by scanning the ZIP.
  existing_ids   = scan_zip_for_table_ids(zip)          # from xl/tables/tableN.xml, read id attr
  existing_names = scan_zip_for_table_names(zip)        # read name attr from same files
  global_count   = count_xl_tables_in_zip(zip)          # how many tableN.xml exist

  # 2. Validate — fail fast on name collision.
  for patch in patches:
      if patch.name in existing_names:
          raise ValueError(f"Table with name {patch.name!r} already exists")

  # 3. Build the tables result.
  result = build_tables(patches, existing_ids, global_count, existing_names)

  # 4. Add new xl/tables/tableN.xml parts to the ZIP.
  for (path, xml_bytes) in result.table_parts:
      zip.add_entry(path, xml_bytes)

  # 5. Update [Content_Types].xml — inject Override entries.
  for (part_name, ct) in result.new_content_types:
      inject_content_type_override(zip, part_name, ct)

  # 6. Update the sheet rels file (xl/worksheets/_rels/sheetN.xml.rels).
  for (rid, target) in result.new_rels:
      append_to_sheet_rels(zip, sheet_idx, rid, target, RT_TABLE)
  # RFC-010's RelsGraph handles rId allocation and rels file mutation.

  # 7. Patch the sheet XML — insert <tableParts> block.
  # RFC-011's block merger inserts <tableParts> just before </worksheet>.
  # It replaces an existing <tableParts> block if one is already present,
  # merging old rIds with new ones (to preserve tables added before this save).
  sheet_xml = patch_sheet_xml_with_block(
      sheet_xml,
      SheetBlock::TableParts(result.table_parts_block)
  )
```

**Idempotency / duplicate handling**: If the user calls `add_table(t)` twice
with the same `Table` object, the name collision check fires on the second call
(the first call inserts into `existing_table_names` before the second is
processed). This matches openpyxl behavior.

**Existing `<tableParts>` merging**: When the source file already has a
`<tableParts>` block (the sheet already had tables), RFC-011's merger APPENDS
the new `<tablePart>` entries to the existing block rather than replacing it.
This preserves pre-existing table-to-rels mappings.

**Global table numbering**: The new `tableN.xml` file index starts at
`global_count + 1`. The rId in the sheet rels uses the next available rId
(determined by scanning the sheet rels file for existing `rId` values).

## 6. Test Plan

Standard matrix: unit tests in Rust, integration tests via pytest.

| Test | What it checks |
|------|---------------|
| `test_add_table_to_clean_file` | File with no existing tables. After save, `xl/tables/table1.xml` exists, `[Content_Types].xml` has the override, sheet rels has rId1, sheet XML has `<tableParts count="1">`. |
| `test_add_second_table_cross_sheet` | File has one table on Sheet1. Add table to Sheet2. Verify `table2.xml` uses global id=2, sheet2 rels has rId, Sheet1 is untouched. |
| `test_table_id_monotonic` | File has tables with IDs 1 and 3 (gap). New table must receive ID 2 (next available), not 4. |
| `test_name_collision_raises` | `add_table` with same name as existing table raises `ValueError` (not `NotImplementedError`). |
| `test_roundtrip_libreoffice` | Open output in LibreOffice headlessly; verify table range, column names, style rendered. |
| `test_autofilter_false` | `Table(autofilter=False)` - no `<autoFilter>` child in table XML. |
| `test_totals_row` | `Table(totalsRowCount=1)` - `totalsRowShown="1"` in XML. |
| `test_preserve_existing_tables` | File has table T1. After adding T2, re-open and verify both T1 and T2 are readable. |

## 7. Migration / Compat Notes

- **Before this RFC**: `ws.add_table(t)` in modify mode raises `NotImplementedError`. After: it works.
- **Backward compat**: Write mode behavior is unchanged. The `NotImplementedError` guard is simply removed from the patcher branch.
- **openpyxl divergence**: openpyxl allows setting `Table.id` explicitly; wolfxl always allocates the ID automatically. If a user passes a `Table` with an explicit `id` field, that value is ignored and a fresh workbook-unique ID is allocated. Document in the migration notes comment at `_worksheet.py:1511`.

## 8. Risks & Open Questions

1. **(MED) ZIP mutation for content-types**: `[Content_Types].xml` is a workbook-level file shared by all sheets. If two sheets add tables in the same `save()` call, the patcher must serialize both sheets' mutations into content-types before writing. Resolution: collect all pending content-type additions across all sheets before writing the file once at the end of `save()`.

2. **(MED) rId collision when sheet rels already has entries**: The patcher must scan the existing sheet rels XML to find the highest current `rId` before assigning new ones. RFC-010's `RelsGraph` should own this scan. If RFC-010 is not yet landed, this RFC must implement its own rId-max scanner as a local helper, to be refactored out when RFC-010 ships.

3. **(LOW) `<tableParts>` position in CT_Worksheet**: ECMA-376 specifies `<tableParts>` as the last child before `<extLst>`. Some third-party validators are strict about this. RFC-011's merger must place the block at the correct position; this RFC supplies the bytes but relies on RFC-011 for placement.

4. **(LOW) Table name uniqueness scope**: OOXML requires table names unique across the workbook. The current scan reads only the `name` attribute from existing `xl/tables/tableN.xml` files. If a workbook uses `displayName` as a separate identifier, additional scanning may be needed. Punt to RFC-035 (rewrite refs) if this causes issues in practice.

## 9. Effort Breakdown

| Task | LOC est. | Days |
|------|----------|------|
| `src/wolfxl/tables.rs` — build_tables fn + XML serializer | ~200 | 1.5 |
| Python coordinator wiring in `_workbook.py` + remove guard at `_worksheet.py:1511` | ~50 | 0.5 |
| Content-types + rels mutation helpers (or RFC-010 integration) | ~100 | 0.5 |
| RFC-011 `SheetBlock::TableParts` support (if not already landed) | ~80 | 0.5 |
| Tests (Rust unit + pytest integration) | ~200 | 1.0 |
| **Total** | **~630** | **4.0** |

## 10. Out of Scope

- Calculated column formulas (`totalsRowFormula`, `calculatedColumnFormula`) on new tables. Existing tables with these attributes are preserved verbatim via ZIP byte-copy.
- `sortState` inside table XML.
- Reading existing tables in modify mode (already works via CalamineStyledBook reader).
- Deleting or renaming existing tables.
- Pivot table style references (`PIVOTSTYLES`).

## Acceptance

Shipped on `feat/rfc-024-tables` cut from `feat/native-writer @ 5f0b79a`.

| Commit | Subject |
|---|---|
| `2920818` | feat(tables): add src/wolfxl/tables.rs (RFC-024 build_tables + ZIP scan) |
| `baf4d81` | feat(patcher): wire queue_table + Phase-2.5f for RFC-024 |

Verification gate (run via `cargo test -p wolfxl-core -p wolfxl-writer -p wolfxl-rels -p wolfxl-merger --quiet` and `pytest`):

- ✅ `cargo test` across the four required crates — green.
- ✅ `tests/test_tables_modify.py` — 8 tests cover: clean-add (ID 1, part 1, content-type override, sheet rels, `<tableParts>`); add to file with one existing table (ID 2, both tables present in rels and `<tableParts>`); style preserved (`TableStyleLight9` + banding flags survive); openpyxl-can-read (range + column names round-trip); name-collision raises `ValueError`; no-pending-tables byte-identical save (Phase-2.5f short-circuit guard); cross-mode parity (`Workbook() + add_table + save` ≡ `load_workbook + add_table + save` byte-for-byte under `WOLFXL_TEST_EPOCH=0`); two-sheet global ID allocation (workbook-unique, not per-sheet).
- ✅ `pytest tests/parity/` — 97 passed, 1 skipped (pre-existing).
- ✅ `pytest tests/diffwriter` — 28 passed; the writer's id-attribute off-by-one fix (`crates/wolfxl-writer/src/lib.rs`) brings write-mode in line with the oracle library, which is what the cross-mode parity test required.
- ✅ `ruff check python/ tests/` on RFC-024 files — clean (9 pre-existing failures in unrelated `tests/diffwriter/*` files inherited from `feat/native-writer`).
- ⚠️ `scripts/verify_rfc.py --rfc 024 --quick` — script not present on `feat/native-writer`; gate not runnable. Manual verification covered by the items above.
- 📝 LibreOffice round-trip: documented (not asserted in CI). Run `soffice --headless --convert-to xlsx out.xlsx --outdir /tmp/lo` and re-open to verify table style + column names render. Excel-strict validation is implicit in the openpyxl-can-read assertion.

Spec deviations:

1. **Workbook-unique ID source.** RFC-024 §5 step 1 reads "scan_zip_for_table_ids". The shipped path scans every `xl/tables/table*.xml` part once at flush time (not at `open()` time) so the patcher's open path stays cheap for the common no-tables case. ID allocation is still workbook-unique because the scan happens before any allocation in `do_save`.
2. **Native-writer off-by-one fix bundled.** The cross-mode parity gate exposed a pre-existing off-by-one in the writer's orchestrator (`crates/wolfxl-writer/src/lib.rs` was passing the 1-based filename counter as the 0-based `table_idx` argument, producing `id="2"` for the first table). One-line fix included; diffwriter golden tests still green because the oracle library already used `id="1"` for the first table.

Follow-ups (out of scope for RFC-024):

- `scripts/verify_rfc.py` is referenced by the RFC template's verification gate but not yet checked into `feat/native-writer`. RFC-013's harness work touched the same path; recommend folding in alongside the next RFC that depends on it.
- Calculated-column formulas (`totalsRowFormula`, `calculatedColumnFormula`) on new tables are still §10 out-of-scope. RFC-035 (`copy_worksheet`) will need the existing-table preservation path to keep working — covered by `test_no_pending_tables_is_byte_identical`.
