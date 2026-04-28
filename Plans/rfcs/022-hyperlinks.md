# RFC-022: T1.5 — Hyperlinks (modify mode)

Status: Shipped
Owner: pod-P3
Phase: 3
Estimate: M
Depends-on: RFC-010, RFC-011
Unblocks: RFC-030, RFC-031, RFC-035 (rewrite refs)

## 1. Problem Statement

`python/wolfxl/_cell.py:181-185` raises today:

```python
raise NotImplementedError(
    "Setting cell.hyperlink on an existing file is a T1.5 follow-up. "
    "Write mode (Workbook() + save) is supported - open the file via "
    "Workbook() rather than load_workbook()."
)
```

Users running ETL or batch-edit pipelines that openpyxl handles in three lines
(`wb = load_workbook(p); ws["A1"].hyperlink = "https://..."; wb.save(p)`) hit
this error and have to fall back to openpyxl, which defeats the wolfxl
performance story for any workflow that combines bulk numeric writes with even
a single hyperlink. After this RFC, the same call in modify mode flushes a new
`<hyperlink>` entry into the worksheet's `<hyperlinks>` block and (for external
URLs) appends a `<Relationship>` to `xl/worksheets/_rels/sheet{N}.xml.rels` —
without re-serializing any other part. Setting `cell.hyperlink = None` removes
both the entry and the rels relationship.

Setter plumbing already lands the data in `_pending_hyperlinks` on the
worksheet (`python/wolfxl/_cell.py:186-201`). Read-side dispatch already pulls
existing hyperlinks via `Worksheet._get_hyperlinks_map`
(`python/wolfxl/_worksheet.py:1564-1597`). The write-mode flush in
`_flush_compat_properties` (`python/wolfxl/_worksheet.py:1888-1904`) translates
`Hyperlink` dataclasses into the writer's dict shape. **This RFC ports that
flush path to the patcher** by invoking a new `XlsxPatcher.queue_hyperlink()`
method and emitting the merged sheet XML + rels at save time.

## 2. OOXML Spec Surface

**Element**: `CT_Hyperlink` — ECMA-376 Part 1, §18.3.1.34 (within
`CT_Hyperlinks`, §18.3.1.35; within `CT_Worksheet`, §18.3.1.99).

**Schema namespace**: `http://schemas.openxmlformats.org/spreadsheetml/2006/main`
(default for the worksheet doc).

**Relationship namespace**: `http://schemas.openxmlformats.org/officeDocument/2006/relationships`
(prefix conventionally `r`; declared on the `<worksheet>` root in every
real-world file).

**Attributes** on `<hyperlink>`:

| Attribute  | Required | Type   | Notes |
|---|---|---|---|
| `ref`      | yes      | ST_Ref | A1 reference or range (`B2`, `B2:C5`). |
| `r:id`     | one-of   | ST_Rid | Relationship ID — used for **external** targets only. |
| `location` | one-of   | string | Internal target — `'Sheet2'!A1`, no leading `#`. |
| `tooltip`  | no       | string | Hover text. |
| `display`  | no       | string | Display text override. Cell value still wins for rendering; this is metadata. |

`r:id` and `location` are mutually exclusive in practice (Excel prefers `r:id`
when both are present). Per §18.3.1.34, at least one of the two must be
present for the element to be meaningful, but the schema does not enforce it.

**Container ordering inside `CT_Worksheet` (§18.3.1.99)**: `<hyperlinks>`
appears **after** `<mergeCells>` and `<conditionalFormatting>` /
`<dataValidations>`, **before** `<printOptions>`, `<pageMargins>`,
`<pageSetup>`, `<headerFooter>`, `<rowBreaks>`, `<colBreaks>`,
`<customProperties>`, `<cellWatches>`, `<ignoredErrors>`, `<smartTags>`,
`<drawing>`, `<legacyDrawing>`, `<legacyDrawingHF>`, `<picture>`,
`<oleObjects>`, `<controls>`, `<webPublishItems>`, `<tableParts>`, `<extLst>`.
RFC-011's block merger uses this ordering table; the new `Hyperlinks` block
slots in at the documented position.

**Relationship type URI** (for external links):
`http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink`

**External URLs require `TargetMode="External"`** on the `<Relationship>` —
internal-to-archive parts omit it. See the live fixture at
`tests/fixtures/tier2/13_hyperlinks.xlsx` (`xl/worksheets/_rels/sheet2.xml.rels`)
for canonical encoding, including `&amp;` escaping in URL targets.

## 3. openpyxl Reference

**Source files**:

- `.venv/lib/python3.14/site-packages/openpyxl/worksheet/hyperlink.py:9-46` —
  `Hyperlink` and `HyperlinkList` Serialisable classes.
- `.venv/lib/python3.14/site-packages/openpyxl/cell/cell.py:231-246` —
  `Cell.hyperlink` setter. Wraps a bare string in
  `Hyperlink(ref="", target=val)`, sets `ref` from the cell coordinate, and
  side-effects: if `cell.value` is `None`, populates it with `target` or
  `location`. Setting to `None` clears the link.
- `openpyxl.worksheet._writer.WorksheetWriter.write_hyperlinks` (in
  `worksheet/_writer.py`) emits the block. For each hyperlink with a
  non-None `target`, openpyxl creates a Relationship with the URL and assigns
  the rId back to `link.id` before serialisation.

**Algorithm summary**:

1. On `cell.hyperlink = X`, `Cell` instance stores the link locally
   (`self._hyperlink`).
2. On `Workbook.save()`, `WorksheetWriter` walks `ws._cells` collecting
   non-None `_hyperlink` attributes into a `HyperlinkList`.
3. For each link: if `link.target` is set, register a Relationship of type
   `hyperlink` with `Target=link.target`, `TargetMode="External"`, capture the
   rId, and write it back to `link.id`.
4. If `link.location` is set (and `target` is not), emit `location="..."`
   instead of `r:id`.
5. Serialize `<hyperlinks>` containing one `<hyperlink>` per link, in cell
   insertion order (no canonicalization).
6. The relationships file is regenerated from scratch — openpyxl always
   re-allocates rIds on save, so old rId numbers are not preserved.

**Edge cases openpyxl handles**:

- Bare string target → wrapped in `Hyperlink(target=val)` (cell.py:241-242).
- `cell.value is None` after assignment → auto-populated with the URL or
  location (cell.py:245-246). **wolfxl already mirrors this** via the
  read-side cache update in `_cell.py:199-201`.
- Range refs (`B2:C5`) on `ref` — accepted as a string, no special handling.

**Edge cases openpyxl mishandles** (verified by reading the source):

- URL fragments (`https://x.com/page#section`) are emitted verbatim in
  `Target="…"` but Excel itself sometimes splits the fragment off into a
  `location` attribute on the `<hyperlink>`. openpyxl never does this, so
  fragment-bearing URLs round-trip but show different click behaviour in Excel
  vs a fresh openpyxl-saved file. **wolfxl will match openpyxl**: keep
  fragments in the rels target, do not split.
- Whitespace in URLs is not URL-encoded by openpyxl. wolfxl will likewise pass
  the string through verbatim — the live fixture (`13_hyperlinks.xlsx` rId3,
  `https://example.com/search?q=excel%20bench&sort=desc#section-2`) shows the
  user is expected to pre-encode.
- Multiple hyperlinks on the same cell: openpyxl silently overwrites — only
  the last `cell.hyperlink = X` survives. wolfxl's `_pending_hyperlinks` dict
  has the same overwrite semantics; pin in tests.

**What we will NOT copy**:

- openpyxl's full re-serialisation of every relationship on save. The patcher
  preserves existing rId numbers (RFC-010) so unrelated rels (printerSettings,
  drawings, vmlDrawingHF) keep their original IDs and any other part of the
  file referencing them still resolves correctly.

## 4. WolfXL Surface Area

### 4.1 Python coordinator

**Files touched**:

| File | Lines | Change |
|------|-------|--------|
| `python/wolfxl/_cell.py` | 175-201 | Replace `if writer is None: raise` guard. Both `writer` and `patcher` paths land in `_pending_hyperlinks`. |
| `python/wolfxl/_worksheet.py` | 1821-1857 (`_flush_to_patcher`) | After per-cell value/format flush, iterate `_pending_hyperlinks` and call `patcher.queue_hyperlink(...)` / `patcher.queue_hyperlink_delete(...)`. Mirrors the writer-side block at lines 1888-1904. |

**Public Python API unchanged**: `cell.hyperlink = "https://..."`,
`cell.hyperlink = Hyperlink(target=..., location=..., tooltip=..., display=...)`,
`cell.hyperlink = None` all keep their current signatures and side-effects. The
read-side cache update at `_cell.py:199-201` and the dataclass at
`python/wolfxl/worksheet/hyperlink.py:12-28` are unchanged.

**Setter dispatch (new logic at `_cell.py:175-185`)**:

```python
@hyperlink.setter
def hyperlink(self, value: Any) -> None:
    ws = self._ws
    wb = ws._workbook
    if wb._rust_writer is None and wb._rust_patcher is None:
        raise RuntimeError("workbook has no backend; this should not happen")
    # Both writer and patcher consume the same _pending_hyperlinks dict.
    if value is None:
        ws._pending_hyperlinks[self.coordinate] = None  # delete sentinel
        if ws._hyperlinks_cache is not None:
            ws._hyperlinks_cache.pop(self.coordinate, None)
        return
    if isinstance(value, str):
        from wolfxl.worksheet.hyperlink import Hyperlink as _Hl
        value = _Hl(target=value)
    ws._pending_hyperlinks[self.coordinate] = value
    if ws._hyperlinks_cache is None:
        ws._hyperlinks_cache = {}
    ws._hyperlinks_cache[self.coordinate] = value
```

The `None` value becomes a delete sentinel (was `pop()` before — see
`_cell.py:186-191`). The patcher needs to learn about deletions explicitly so
it can also drop the rels entry; this is the single behavioural change in
Python.

### 4.2 Patcher (modify mode)

**New module**: `src/wolfxl/hyperlinks.rs`.

**PyO3-exposed methods on `XlsxPatcher`** (new, registered in
`src/wolfxl/mod.rs`):

```rust
fn queue_hyperlink(&mut self, sheet: &str, cell: &str, payload: &Bound<'_, PyDict>) -> PyResult<()>;
fn queue_hyperlink_delete(&mut self, sheet: &str, cell: &str) -> PyResult<()>;
```

**Internal types**:

```rust
#[derive(Debug, Clone)]
pub struct HyperlinkPatch {
    pub coordinate: String,    // "B2"
    pub target: String,        // URL or internal ref ("'Sheet2'!A1")
    pub is_internal: bool,
    pub tooltip: Option<String>,
    pub display: Option<String>,
}

#[derive(Debug, Clone)]
pub enum HyperlinkOp {
    Set(HyperlinkPatch),
    Delete,
}

/// Build the merged `<hyperlinks>` block bytes for a sheet.
///
/// `existing` is parsed from the live sheet XML by [`extract_hyperlinks`].
/// `ops` is the queued patches. `rels` is the per-sheet rels graph from
/// RFC-010, mutated to add new external rIds and remove orphaned ones.
///
/// Returns `(block_bytes, deleted_rids)`. Block bytes are passed to
/// RFC-011's block merger as `SheetBlock::Hyperlinks(bytes)`. Deleted
/// rIds are removed from `rels` after computing the new block (so they
/// don't get reused for fresh patches).
pub fn build_hyperlinks_block(
    existing: BTreeMap<String, ExistingHyperlink>,
    ops: &BTreeMap<String, HyperlinkOp>,
    rels: &mut RelsGraph,
) -> (Vec<u8>, Vec<RelId>);

/// Parse the existing `<hyperlinks>` block out of a sheet XML, resolving
/// `r:id` attributes against the sheet's rels graph so we get URLs not
/// rIds.
pub fn extract_hyperlinks(sheet_xml: &[u8], rels: &RelsGraph) -> BTreeMap<String, ExistingHyperlink>;
```

**ZIP parts read** (per save):

- `xl/worksheets/sheet{N}.xml` (already read by `do_save` for cell patches)
- `xl/worksheets/_rels/sheet{N}.xml.rels` (new — may be absent for sheets
  with no rels; RFC-010 returns an empty graph in that case)

**ZIP parts mutated**:

- `xl/worksheets/sheet{N}.xml` — RFC-011's merger replaces the `<hyperlinks>`
  block (or inserts one at the spec-defined position).
- `xl/worksheets/_rels/sheet{N}.xml.rels` — written if non-empty after edits;
  if all rels were deleted (unusual — would imply no comments, no tables, no
  external links remain) the file is deleted from the ZIP.

**ZIP parts unchanged**: `[Content_Types].xml` (hyperlinks share the
worksheet's already-listed content type — no `<Override>` change needed),
`xl/_rels/workbook.xml.rels`, all other sheets, all comments/drawings/etc.

### 4.3 Native writer (write mode)

Already wired. `crates/wolfxl-writer/src/emit/sheet_xml.rs:451-491`
(`emit_hyperlinks`) emits the block. `crates/wolfxl-writer/src/emit/rels.rs:128-196`
(`emit_sheet`) emits external-link relationships. The native writer rebuilds
both files from scratch so no merger is needed there. **No native-writer
changes are required for this RFC** — it's purely the patcher seam.

## 5. Algorithm

### 5.1 High level

```
on save:
  for each sheet with non-empty _pending_hyperlinks:
    rels = RFC-010::parse(read("xl/worksheets/_rels/sheet{N}.xml.rels"))
    existing = extract_hyperlinks(read("xl/worksheets/sheet{N}.xml"), rels)

    # Apply ops in coordinate order so behaviour is deterministic
    merged = existing.clone()
    for (coord, op) in pending.iter():
        match op {
            Set(patch)  => merged.insert(coord, patch.into())
            Delete      => merged.remove(coord)  // also flag rid for removal
        }

    # Build new rels: keep rIds for unchanged links, allocate new for added,
    # mark orphans for delete.
    deleted_rids = []
    for (coord, hl) in &merged:
        if hl is external && hl.rid is None:
            hl.rid = rels.add_external(hyperlink_rt, hl.target)
    for (coord, removed) in existing - merged:
        if removed.rid is set:
            deleted_rids.push(removed.rid)
    rels.remove_many(deleted_rids)

    block = serialize(merged)  // RFC-011 SheetBlock::Hyperlinks(block)
    new_sheet_xml = RFC-011::merge_blocks(sheet_xml, [(BlockKind::Hyperlinks, block)])
    new_rels_xml  = rels.serialize()

    file_patches[sheet_path]      = new_sheet_xml
    file_patches[rels_path]       = new_rels_xml
```

### 5.2 Idempotency

`build_hyperlinks_block` is a pure function of `(existing, ops, rels)`. Calling
`save()` twice with no further mutations produces byte-identical output (modulo
the `WOLFXL_TEST_EPOCH=0` guard for ZIP mtimes). The algorithm preserves
existing rId numbers for unchanged external links, so files that already had
hyperlinks will only see new rIds for **newly added** external links, never
renumbering of unrelated ones.

### 5.3 Deletion semantics

When `cell.hyperlink = None` is queued:

1. Python places `None` in `_pending_hyperlinks[coord]` (delete sentinel).
2. Patcher's `queue_hyperlink_delete(sheet, coord)` records `HyperlinkOp::Delete`.
3. At save: if the link was external, look up its rId in the existing rels
   graph. Add to `deleted_rids`. Remove from `rels` after merging.
4. If `merged` ends up empty, RFC-011 removes the entire `<hyperlinks>` block
   (not just emits an empty one).
5. If the rels file becomes empty (no comments, no tables, no external links
   remain), the ZIP entry is deleted entirely. RFC-010 owns this decision via
   `RelsGraph::is_empty()`.

### 5.4 Internal vs external classification

The classification is a property of the `Hyperlink` Python dataclass: `target`
field set → external; `location` field set → internal. The setter never
auto-detects from string content (no "starts with `#`" heuristic). This
matches the field already documented at
`python/wolfxl/worksheet/hyperlink.py:17-19` and the Rust model at
`crates/wolfxl-writer/src/emit/sheet_xml.rs:462-470` ("Source of truth is the
field, NOT a string prefix on `target`").

### 5.5 rId allocation seam

RFC-010 owns the rId graph. The hyperlink patcher only calls
`rels.add_external(RT_HYPERLINK, target_url)` and `rels.remove(rid)`. It does
not assume any particular numbering scheme — the native writer's "comments
get rId1+rId2 then tables then hyperlinks" convention
(`sheet_xml.rs:454-457`) is irrelevant in the patcher because RFC-010
preserves the existing file's ordering. New rIds use the next available
number.

## 6. Test Plan

Standard verification matrix from the plan §Verification:

1. **Rust unit tests** — `src/wolfxl/hyperlinks.rs::tests`:
   - `extract_hyperlinks_resolves_rids` — round-trip the
     `13_hyperlinks.xlsx` fixture's sheet2 block, assert all four entries
     come out with correct external/internal classification.
   - `build_block_preserves_existing_rids` — start with two existing
     external links (rId1, rId2), add a third → result has rId1, rId2, rId3
     in that order; rId1/rId2 unchanged.
   - `delete_external_removes_rid` — start with rId1 external link, queue
     delete → block empty, rId1 in `deleted_rids`.
   - `internal_link_no_rid` — internal hyperlink emits `location=` with no
     `r:id=` and does not allocate from RFC-010.
   - `mixed_internal_external_in_one_block` — three links: ext, int, ext →
     correct attribute mix in serialised order.
   - `attr_escape_url_with_ampersand` — target containing `&` produces
     `&amp;` in the rels file (already exercised in `rels.rs::sheet_rels_all_three_kinds_coexist`).
   - `tooltip_and_display_emitted` — both attributes round-trip.

2. **Golden round-trip (`tests/diffwriter/`)** — new fixture
   `tests/diffwriter/cases/hyperlinks_modify.py` opens
   `13_hyperlinks.xlsx`, sets `ws["B6"].hyperlink = "https://new.example.com"`,
   saves, asserts byte-equal to a checked-in golden under
   `WOLFXL_TEST_EPOCH=0`. Generate the golden by running the script once and
   verifying with `unzip -l` + manual inspection.

3. **openpyxl parity (`tests/parity/`)** — new test
   `tests/parity/test_hyperlink_modify_parity.py`:
   - Open `13_hyperlinks.xlsx` with both `wolfxl.load_workbook(p, modify=True)`
     and `openpyxl.load_workbook(p)`.
   - Add the same external + internal hyperlink to the same cells.
   - Save both. Re-open both with `openpyxl`. Compare every
     `cell.hyperlink.target / location / tooltip / display / ref` for every
     hyperlinked cell across both files. Diffs go to ratchet.

4. **LibreOffice round-trip** — new fixture in `tests/parity/libreoffice/`:
   add 5 external links + 2 internal links to `13_hyperlinks.xlsx`,
   save, open with `soffice --headless --convert-to xlsx`, re-read with
   wolfxl, assert all 7 links present with correct targets. Bonus: open
   the file in LibreOffice GUI and click each link manually (one-time
   manual verification, captured as a screenshot in the PR).

5. **Cross-mode** — open a file in modify mode, add a hyperlink, save.
   Open the result in **read** mode, assert `cell.hyperlink.target` matches.
   Then re-open in **modify** mode, delete the hyperlink, save. Assert the
   `<hyperlinks>` block is gone and the rels file no longer references the
   URL.

6. **Regression fixture** — copy `13_hyperlinks.xlsx` to
   `tests/parity/fixtures/regress/rfc022_existing_hyperlinks.xlsx` and lock
   the byte-identity of "open + save with no changes" via `WOLFXL_TEST_EPOCH=0`.
   Catches future regressions where the patcher rewrites the block
   spuriously.

**RFC-specific cases**:

- **rId preservation across unrelated parts**: open a file that has
  hyperlinks AND a `printerSettings` rel (live fixture
  `time_series/ilpa_pe_fund_reporting_v1.1.xlsx` has this on sheet1 with
  `rId1=printerSettings, rId2=vmlDrawing`). Add a new hyperlink. Assert
  printerSettings/vmlDrawing rId numbers are unchanged.
- **Setting + deleting same cell in one session** (`ws["A1"].hyperlink = X;
  ws["A1"].hyperlink = None`): only the delete should reach the patcher.
  `_pending_hyperlinks["A1"]` is `None` at flush time; the cell never had
  a link in the original file → no-op.
- **Range refs**: `Hyperlink(ref="A1:C3", target="...")` round-trips
  with `ref="A1:C3"`. (openpyxl supports this; we should too.)
- **Same URL on multiple cells** allocates separate rIds (matches openpyxl
  and Excel) — no dedup of `Target` values in rels.

## 7. Migration / Compat Notes

**Behaviour diffs vs openpyxl**:

- openpyxl regenerates **all** rIds in the rels file on save. wolfxl
  preserves existing rIds. This means if a downstream tool depends on
  specific rId values it will see different numbers from openpyxl's output
  but stable numbers across wolfxl saves. This is strictly better but
  worth calling out.
- openpyxl's setter side-effect (auto-populate `cell.value = url` when
  `cell.value is None`) is **not** mirrored in modify mode currently —
  the read-side cache update at `_cell.py:199-201` reflects the link in
  `cell.hyperlink` but does not touch `cell.value`. **Decision**: add the
  side-effect to match openpyxl exactly; users coming from openpyxl
  expect it. Implement in `_cell.py:175-185` after the cache update.

**Backward-compat shims**: none required. The setter signature is
unchanged; only the previously-raised `NotImplementedError` is now
honoured.

**Feature flag during rollout**: `WOLFXL_DISABLE_HYPERLINK_PATCH=1` env
var bypasses the new path and re-raises the old `NotImplementedError`.
Lets us ship the feature dark for one release and flip the default once
the parity ratchet is green. Remove in the release after.

## 8. Risks & Open Questions

1. **rId collisions when an existing rels file has gaps** (e.g.,
   `rId1, rId3` with no `rId2`). Resolution: RFC-010's
   `RelsGraph::next_id()` walks `(max(existing) + 1)`, never reuses
   gaps. Verified by the live fixture `13_hyperlinks.xlsx` where rIds are
   contiguous; gap behaviour will be unit-tested in RFC-010.

2. **Internal hyperlinks pointing at sheet names with single quotes**
   (e.g., `'Joe''s Sheet'!A1` — Excel doubles internal quotes). Resolution:
   the `location` attribute value is XML-escaped as a normal string; the
   double-quote sheet-name encoding is the user's responsibility and
   matches openpyxl's behaviour. Add a test fixture with such a sheet.

3. **Hyperlink ref as a range vs single cell**: extraction needs to
   tolerate `ref="A1:C3"` and key the patch dict by the full range
   string (not just the top-left cell). Resolution: use the raw `ref`
   attribute as the dict key. This means `ws["A1"].hyperlink = X`
   followed by reading `ws["A2"].hyperlink` returns `None` even if the
   original file had a range link covering A1:C3 — but this is the same
   gap as openpyxl, so we punt to RFC-XXX (post-1.0).

4. **External link target with embedded fragment** (`https://x.com#frag`):
   we keep the fragment in the rels `Target=` attribute; do not split
   into `location`. Matches openpyxl. Document in §3 (already done).

5. **Concurrent modify-mode saves on the same path**: if two processes
   open the same xlsx in modify mode, last write wins; rId allocations
   are not coordinated. Punt — this is a general patcher concern, not
   hyperlink-specific.

6. **What if the workbook has no `xl/worksheets/_rels/` directory at all?**
   (Common for files openpyxl created without any external refs.)
   Resolution: RFC-010 returns an empty `RelsGraph`, the patcher writes
   a fresh rels file, and the ZIP rewriter (`mod.rs:do_save`) needs to
   know it's a *new* entry, not a replacement. The current `file_patches`
   HashMap is only checked against entries already in the ZIP
   (`mod.rs:343-350`); we need to extend the rewrite loop to also append
   new entries that weren't in the source. Tracked as an RFC-010
   sub-task.

## 9. Effort Breakdown

| Task | LOC est. | Days |
|---|---|---|
| `src/wolfxl/hyperlinks.rs` (extract + build + delete logic) | 350 | 2 |
| PyO3 wiring (`queue_hyperlink`, `queue_hyperlink_delete`) in `src/wolfxl/mod.rs` | 80 | 0.5 |
| `do_save` integration: read sheet rels, invoke RFC-010, call hyperlinks module, hand bytes to RFC-011, write rels back | 120 | 1 |
| Python: `_cell.py` setter dispatch + `_worksheet.py` `_flush_to_patcher` extension | 60 | 0.5 |
| Rust unit tests (7 cases) | 250 | 1 |
| Diffwriter golden + parity test + LibreOffice round-trip | 200 | 1 |
| Regression fixture lock + cross-mode test | 80 | 0.5 |
| **Total** | **1140** | **6.5 (≈ 1.5 calendar weeks with review)** |

## 10. Out of Scope

- **Range hyperlinks** (`ref="A1:C3"`): supported on round-trip but reads
  from individual cells inside the range return `None`. Full per-cell
  expansion punted to a post-1.0 RFC.
- **Auto-encoding of unencoded URLs** (spaces, unicode in path): we
  pass the user's string through verbatim. Mirrors openpyxl.
- **Hyperlink updates triggered by `move_range` / row insertion**: that's
  RFC-030 / RFC-031 / RFC-035 territory ("Unblocks: 030, 031, 035 rewrite refs").
  The plumbing here exposes `extract_hyperlinks` and
  `build_hyperlinks_block` as the seams those RFCs call into.
- **`<hyperlinks>` block-level attributes** beyond per-link `ref`/`r:id`/
  `location`/`tooltip`/`display`. ECMA defines no others.
- **Threaded comments referenced as hyperlinks**: not a real OOXML
  thing; only mentioning to confirm we're not inventing it.

## Acceptance

Shipped in commit range `2d54cf4..<C5>` on branch `feat/native-writer`.

### Live behavior

- `cell.hyperlink = "https://..."` round-trips in modify mode (clean
  file, fixture with prior hyperlinks, set+delete in one session).
- `cell.hyperlink = Hyperlink(location=...)` round-trips for internal
  links (no rId allocated, sheet rels untouched).
- `cell.hyperlink = None` is the explicit-delete sentinel (INDEX
  decision #5). Removes the rId from `xl/worksheets/_rels/sheetN.xml.rels`
  for external links; mutates only sheet XML for internal links.
- `tooltip` and `display` attributes round-trip through openpyxl.
- URLs with `&` (and other XML-special chars) escape correctly in the
  rels target; round-trip recovers the literal.
- openpyxl-parity side-effect: when a hyperlink is set on a cell with
  `value is None`, the cell value is populated with the
  `display` → `target` → `location` first-non-None.
- No-op modify-mode save remains byte-identical to source (short-circuit
  predicate includes `queued_hyperlinks.is_empty()`).

### Implementation surface

- `src/wolfxl/hyperlinks.rs` — `HyperlinkPatch`, `HyperlinkOp`,
  `ExistingHyperlink`, `extract_hyperlinks`, `build_hyperlinks_block`
  + 9 inline tests.
- `src/wolfxl/mod.rs` — `queued_hyperlinks` field, `queue_hyperlink`
  / `queue_hyperlink_delete` PyO3 setters, `sheet_rels_path_for` /
  `load_or_empty_rels` helpers, Phase 2.5e in `do_save`, three
  `_test_*` PyO3 hooks. Also: rels serialization now routes new rels
  paths to `file_adds` (clean-file path).
- `python/wolfxl/_cell.py` — hyperlink setter dispatches to both
  writer and patcher backends; openpyxl-parity value side-effect.
- `python/wolfxl/_workbook.py` — `_flush_pending_hyperlinks_to_patcher`
  drains `_pending_hyperlinks` per sheet; called BEFORE DV/CF flushes.
- `tests/test_modify_hyperlinks.py` — 8 integration tests.
- `tests/test_cell_read_t1.py` — `test_modify_mode_hyperlink_setter_raises`
  removed (T1.5 guard now obsolete).
- `tests/test_cell_write_t1.py` — `test_modify_mode_raises_with_t15_hint`
  narrowed to comments-only.

`AncillaryPartRegistry` (RFC-013) has its first live caller via Phase
2.5e's `populate_for_sheet` call.

### Deferred to hardening

- LibreOffice round-trip tests (per RFC-013 / RFC-020 / RFC-025 /
  RFC-026 precedent — none ship LibreOffice harness in the live slice).
- Full openpyxl-parity sweep across the cross-product of
  target/location/tooltip/display × set/delete/preserve. Byte-identical
  no-op gate already catches the worst-class regression.
- Range hyperlinks (`ref="A1:C3"`): supported on round-trip, but reads
  from individual cells inside the range still return `None` (RFC-022
  §10).
- `WOLFXL_DISABLE_HYPERLINK_PATCH` feature flag (RFC-022 §7): not
  implemented. The flag was a precaution before RFC-013 was a known
  quantity; with the registry shipping clean and `emit_hyperlinks`
  proven in write mode, the flag adds complexity for marginal value.
