# RFC-010: `*.rels` Graph (Parser + Mutable Graph + Serializer) for the Patcher

Status: Researched
Owner: pod-P1
Phase: 2
Estimate: M
Depends-on: RFC-001
Unblocks: RFC-022, RFC-023, RFC-024, RFC-035

## 1. Problem Statement

The patcher today never opens any `*.rels` part — it raw-copies them. See
`src/wolfxl/mod.rs:323-356` (the rewrite loop reads each ZIP entry and either
substitutes a value from `file_patches` or copies bytes verbatim) and
`src/wolfxl/sheet_patcher.rs` (no rels handling anywhere). The single
non-rels parse path the patcher has is `parse_relationship_targets` in
`src/ooxml_util.rs:80-109`, which is read-only and discards everything except
the `Id → Target` map.

That works for value patches (cell text never references a relationship), but
every modify-mode workflow Phase 3 needs to add carries new rels:

- **Hyperlinks** (RFC-022). External hyperlinks must register as
  `<Relationship Type=".../hyperlink" TargetMode="External" Target="https://..."/>`
  in `xl/worksheets/_rels/sheet{N}.xml.rels`. The `<hyperlink r:id="rIdN"/>`
  child of `<sheetData>` then resolves through that rels file. Today the
  rels file is raw-copied, so the new `r:id` resolves to nothing and Excel
  silently drops the link (or, worse, "Repaired" the file on open).
- **Comments** (RFC-023). Adding a single comment requires two new
  relationships — `.../comments` → `comments{N}.xml` and `.../vmlDrawing`
  → `vmlDrawing{N}.vml` — plus a `<legacyDrawing r:id="..."/>` element in
  the sheet XML. Same failure mode as hyperlinks.
- **Tables** (RFC-024). Each `<tablePart r:id="rIdK"/>` requires a
  `.../table` relationship pointing at `../tables/table{K}.xml`. Tables
  also use **global numbering** across the workbook (see
  `crates/wolfxl-writer/src/emit/rels.rs:128-181` for how the writer
  handles this), so the rels graph must know how to re-allocate `rId`s
  without colliding with already-present ones.
- **Defined-name `move_sheet`** (RFC-035). When the workbook-level
  `xl/_rels/workbook.xml.rels` is reordered (sheet rename across files),
  no rId allocation changes, but the file may still need a serialize
  pass if a sheet is added or removed.

Without a rels graph, every Phase-3 RFC has to either roll its own
mini-parser/serializer (duplication, drift) or stay broken in modify mode.

User-visible failure today, illustrative:

```python
import wolfxl
wb = wolfxl.load_workbook("report.xlsx", modify=True)
wb["Sheet1"]["A1"].hyperlink = "https://example.com"   # NotImplementedError today
wb.save("out.xlsx")
```

After RFC-022 lands, that call routes through `XlsxPatcher.queue_hyperlink(...)`
which then calls into the API specified here.

## 2. OOXML Spec Surface

ECMA-376 Part 1 §15.2 (Open Packaging Conventions — Relationships).
Also CT_Relationships in OPC §10.3.

**Namespace**

```
http://schemas.openxmlformats.org/package/2006/relationships
```

This is the **package** rels namespace and is what every `.rels` file uses
on its root element. It is **distinct** from
`http://schemas.openxmlformats.org/officeDocument/2006/relationships` (the
prefix bound to `r:` inside parts like `workbook.xml` for `r:id` attributes).
See `crates/wolfxl-writer/src/emit/rels.rs:42` (`RELS_NS` constant) for the
value the writer already emits.

**Root element**

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="…" Target="…" [TargetMode="External"]/>
  ...
</Relationships>
```

**`<Relationship>` attributes**

| Attribute    | Required | Notes |
|--------------|----------|-------|
| `Id`         | yes      | Token unique within this rels file. Convention: `rId<N>` with `N >= 1`. The numeric suffix is not required by spec but is universal in real files. |
| `Type`       | yes      | Absolute URI naming the relationship class. |
| `Target`     | yes      | Relative URI (resolved against the part owning the rels file) or an absolute URI when `TargetMode="External"`. |
| `TargetMode` | optional | `Internal` (default) or `External`. External must be set explicitly for hyperlinks, oleObject links, etc. |

**Relationship-type URIs we care about** (full list, not abbreviated — these
are referenced verbatim in code):

| Constant | URI |
|---|---|
| `RT_OFFICE_DOC`     | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument` |
| `RT_CORE_PROPS`     | `http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties` |
| `RT_EXT_PROPS`      | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties` |
| `RT_CUSTOM_PROPS`   | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties` |
| `RT_WORKSHEET`      | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet` |
| `RT_CHARTSHEET`     | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet` |
| `RT_STYLES`         | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles` |
| `RT_THEME`          | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme` |
| `RT_SHARED_STRINGS` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings` |
| `RT_HYPERLINK`      | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink` |
| `RT_COMMENTS`       | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments` |
| `RT_VML_DRAWING`    | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing` |
| `RT_DRAWING`        | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing` |
| `RT_IMAGE`          | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image` |
| `RT_TABLE`          | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/table` |
| `RT_PIVOT_TABLE`    | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable` |
| `RT_PIVOT_CACHE_DEF`| `http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition` |
| `RT_OLE_OBJECT`     | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject` |
| `RT_VBA_PROJECT`    | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/vbaProject` |
| `RT_PRINTER_SETTINGS` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings` |

The first nine are already defined as string constants in
`crates/wolfxl-writer/src/emit/rels.rs:22-40`; we will move them into the new
shared module so both the writer and the patcher consume one source of truth
(see §4.3).

**ID uniqueness rules** (ECMA-376 §15.2.4):
Ids must be unique within a single `*.rels` file; they have no meaning across
files. Real files in `tests/fixtures/tier2/13_hyperlinks.xlsx` use the
`rId<N>` form with monotonically increasing N. Excel does not require ids to
be contiguous — e.g. removing rId3 from a file with rId1..rId5 leaves a gap
and Excel still opens it cleanly.

## 3. openpyxl Reference

File: `.venv/lib/python3.14/site-packages/openpyxl/packaging/relationship.py`
(158 lines).

Algorithm summary:

1. `Relationship` class (`relationship.py:21-50`) is a `Serialisable`
   descriptor model with four attributes: `Id`, `Type`, `Target`,
   `TargetMode`. The descriptors give it XML round-trip for free via the
   metaclass, at the cost of pulling in
   `openpyxl.descriptors.serialisable.Serialisable` and the entire
   container/descriptor machinery.
2. `RelationshipList(ElementList)` (`relationship.py:53-91`) holds a list
   of `Relationship` and overrides `append` so a new relationship without
   an `Id` gets one auto-assigned as `f"rId{len(self)}"`. **This is buggy**
   for our use case: after a `remove`, `len(self)` shrinks and the next
   `append` reuses an id that points at a stale entry. Excel hasn't
   complained about this in practice because openpyxl rebuilds the entire
   workbook from a parsed model on save, so removed rIds are fully
   gone — but a *patcher* that mixes existing-on-disk rIds with new ones
   cannot do that. We must use a true monotonic counter.
3. `RelationshipList.find(content_type)` (`relationship.py:65-73`) is a
   linear scan yielding all matching relationships. Same shape we want.
4. `get_rels_path(path)` (`relationship.py:94-103`) computes the
   conventional rels path for a given part: `xl/workbook.xml` →
   `xl/_rels/workbook.xml.rels`. We replicate this as a free function.
5. `get_dependents(archive, filename)` (`relationship.py:106-130`) parses
   a rels file, then **rewrites every `Target` to be absolute** by joining
   it against the parent folder. We will NOT do this — the patcher needs
   to round-trip the file byte-identical, so `Target` stays relative.
6. `get_rel(archive, deps, ...)` (`relationship.py:133-158`) follows a
   relationship to load the dependent part. Out of scope — the patcher
   resolves rels-to-paths in a separate utility (`ooxml_util.rs:80`).

What we will NOT copy:

- **Descriptor metaclass machinery.** Three Rust struct fields with explicit
  parse/serialize is far cleaner than mirroring Python's descriptor
  framework. The PyO3 crate also can't bridge it without huge boilerplate.
- **`get_rels_path` heuristic for absolute-vs-relative target rewriting.**
  Round-trip determinism is more important than absolute paths.
- **`__init__`'s `type=` shortcut** (`relationship.py:34-50`) where the
  caller passes a short suffix and the constructor expands it against
  `REL_NS`. We always require the full URI — typo-resistant via the
  module-level constants in §2.
- **Auto-Id allocation by `len(self)`.** See bug above. We use a
  monotonic counter that survives removes.

## 4. WolfXL Surface Area

### 4.1 Python coordinator

No new public Python API. The rels graph is consumed only by patcher code
and never crosses the PyO3 boundary directly. Three Phase-3 user-facing
methods (added in their own RFCs, not this one) will internally invoke it:

- `Worksheet.set_hyperlink(cell, target)` (RFC-022)
- `Worksheet.add_comment(cell, text, author)` (RFC-023)
- `Worksheet.add_table(name, ref, ...)` (RFC-024)

### 4.2 Patcher (modify mode) — new module

New file: `src/wolfxl/rels.rs` (estimated ~280 LOC including tests).

Public API:

```rust
//! `.rels` (Open Packaging Conventions Relationships) parser, mutable graph,
//! and serializer for the surgical patcher.

use std::fmt;

/// Relationship Id (e.g. "rId7"). Newtype so we cannot accidentally mix it
/// with a Target string.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RelId(pub String);

impl RelId {
    /// Parse the numeric suffix of an "rId<N>" string. Returns None if the
    /// id does not follow the convention (some legacy files use other
    /// schemes; we preserve those untouched but cannot allocate next to
    /// them).
    pub fn numeric_suffix(&self) -> Option<u32> { ... }
}

impl fmt::Display for RelId { ... }

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TargetMode {
    Internal,
    External,
}

#[derive(Debug, Clone)]
pub struct Relationship {
    pub id: RelId,
    pub rel_type: String,    // full URI (use the RT_* constants below)
    pub target: String,      // verbatim from source; not normalized
    pub mode: TargetMode,
}

/// Parsed + mutable view of a single `*.rels` file.
pub struct RelsGraph {
    rels: Vec<Relationship>,   // preserves source order; see §5
    next_rid: u32,             // monotonic, never decreases (see §3 bug note)
}

impl RelsGraph {
    /// Create an empty graph (for new rels files emitted from scratch).
    pub fn new() -> Self { ... }

    /// Parse an existing rels file. Empty input → empty graph.
    /// Errors: malformed XML, missing required attributes.
    pub fn parse(xml: &[u8]) -> Result<Self, String>;

    /// Append a new relationship, allocating a fresh rId.
    /// Returns the allocated id.
    pub fn add(&mut self, rel_type: &str, target: &str, mode: TargetMode) -> RelId;

    /// Append a relationship with an explicit id (used when migrating an
    /// existing-on-disk rels graph and a caller needs to set rId by hand,
    /// e.g. rebuilding workbook.xml.rels). Panics if id already present.
    pub fn add_with_id(&mut self, id: RelId, rel_type: &str, target: &str, mode: TargetMode);

    /// Remove a relationship by id. No-op if not present.
    /// Does NOT renumber. Gaps are preserved. The next add() still
    /// returns a fresh id strictly greater than every id ever seen.
    pub fn remove(&mut self, id: &RelId);

    /// Look up by id. O(N) — rels files are tiny (< 50 entries in 99.9%
    /// of files we've sampled in `tests/fixtures/`).
    pub fn get(&self, id: &RelId) -> Option<&Relationship>;

    /// Return all relationships of the given type, in source order.
    pub fn find_by_type(&self, rel_type: &str) -> Vec<&Relationship>;

    /// Find by target — used to dedupe (same external URL hyperlinked from
    /// two cells should reuse one rId).
    pub fn find_by_target(&self, target: &str, mode: TargetMode) -> Option<&Relationship>;

    /// Iterate in source order.
    pub fn iter(&self) -> std::slice::Iter<'_, Relationship>;

    /// Serialize back to bytes. Always emits the canonical preamble
    /// (declaration + root open) and preserves source order. See §5 for
    /// formatting determinism requirements.
    pub fn serialize(&self) -> Vec<u8>;
}

/// Compute the conventional rels path for a part:
///   "xl/workbook.xml"           -> "xl/_rels/workbook.xml.rels"
///   "xl/worksheets/sheet1.xml"  -> "xl/worksheets/_rels/sheet1.xml.rels"
///   "[Content_Types].xml"       -> never has a rels file; returns None
pub fn rels_path_for(part_path: &str) -> Option<String>;

// ---------------------------------------------------------------------------
// Relationship-type URIs. Single source of truth; the writer's emit/rels.rs
// constants will be removed and re-exported from here.
// ---------------------------------------------------------------------------
pub mod rt {
    pub const OFFICE_DOC: &str       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    pub const CORE_PROPS: &str       = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    pub const EXT_PROPS: &str        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    pub const CUSTOM_PROPS: &str     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    pub const WORKSHEET: &str        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    pub const CHARTSHEET: &str       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    pub const STYLES: &str           = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const THEME: &str            = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    pub const SHARED_STRINGS: &str   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    pub const HYPERLINK: &str        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const COMMENTS: &str         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    pub const VML_DRAWING: &str      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
    pub const DRAWING: &str          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    pub const IMAGE: &str            = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    pub const TABLE: &str            = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    pub const PIVOT_TABLE: &str      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
    pub const PIVOT_CACHE_DEF: &str  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
    pub const OLE_OBJECT: &str       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
    pub const VBA_PROJECT: &str      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vbaProject";
    pub const PRINTER_SETTINGS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
}
```

ZIP parts read/mutated/emitted by patcher modules using this graph:

| ZIP part | Role | Writers (RFCs) |
|---|---|---|
| `_rels/.rels` | top-level (workbook + props) | rare; only RFC-035 sheet-add |
| `xl/_rels/workbook.xml.rels` | sheet/styles/SST table | RFC-035 |
| `xl/worksheets/_rels/sheet{N}.xml.rels` | per-sheet (hyperlinks/comments/tables) | RFC-022, RFC-023, RFC-024 |
| `xl/comments/_rels/...` | rare (comments → embedded objects) | not Phase 3 |

Wiring point in `src/wolfxl/mod.rs`:

The save loop at `mod.rs:323-356` will gain a third branch. Today:

```rust
let data = if let Some(patched) = file_patches.get(&name) {
    patched.clone()
} else {
    let mut buf = Vec::new();
    file.read_to_end(&mut buf)?;
    buf
};
```

After RFC-010+RFC-022 land, `file_patches` will also contain the serialized
bytes for any `*.rels` that got mutated. The patcher modules call
`RelsGraph::serialize()` and stuff the result into `file_patches` before the
loop runs.

### 4.3 Native writer (write mode)

The writer already emits rels via `crates/wolfxl-writer/src/emit/rels.rs`
(see line counts: 407 lines, well-tested). Migration plan:

1. Move the `RT_*` constants from `crates/wolfxl-writer/src/emit/rels.rs:22-40`
   into a new `wolfxl-rels` crate (or into `src/wolfxl/rels.rs` and re-export
   from the writer). Decision: put it in `src/wolfxl/rels.rs` because the
   `src/wolfxl/` tree already lives outside the wolfxl-writer crate boundary
   and is consumed by both the patcher (via `wolfxl::rels::*`) and (via a
   `pub use`) the writer's `emit/rels.rs`. This avoids spinning up yet
   another workspace crate.
2. Refactor `emit_root`/`emit_workbook`/`emit_sheet` (`emit/rels.rs:83-196`)
   to build a `RelsGraph` and call `serialize()` instead of hand-rolling
   the XML in `String`s. This is **mechanical**; the test suite at
   `emit/rels.rs:198-407` validates the output is structurally equivalent.
3. Net result: one `RelsGraph` implementation drives both modes. Drift
   between writer-emitted rels and patcher-emitted rels becomes
   structurally impossible.

No new state on `NativeWorkbook`. The graph is constructed locally inside
each `emit_*` call from the existing `Workbook` model.

## 5. Algorithm

### 5.1 Parsing

Streaming via `quick_xml::Reader`, mirroring the style already used in
`src/ooxml_util.rs:80-109` and `src/wolfxl/sheet_patcher.rs:75-245`.

```text
fn parse(xml: &[u8]) -> Result<RelsGraph, String>:
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut rels = Vec::new();
    let mut max_seen: u32 = 0;
    let mut buf = Vec::new();

    loop:
        match reader.read_event_into(&mut buf):
            Start(e) | Empty(e) if e.local_name() == b"Relationship":
                let id_str = attr_value(&e, b"Id")
                    .ok_or("Relationship missing Id")?;
                let rel_type = attr_value(&e, b"Type")
                    .ok_or("Relationship missing Type")?;
                let target = attr_value(&e, b"Target")
                    .ok_or("Relationship missing Target")?;
                let mode = match attr_value(&e, b"TargetMode").as_deref():
                    Some("External") => TargetMode::External,
                    _ => TargetMode::Internal,
                ;
                let id = RelId(id_str);
                if let Some(n) = id.numeric_suffix(): max_seen = max(max_seen, n);
                rels.push(Relationship { id, rel_type, target, mode });
            Eof: break
            Err(e): return Err(format!("XML parse error: {e}"))
            _: continue
        buf.clear()

    Ok(RelsGraph { rels, next_rid: max_seen + 1 })
```

Notes:

- We tolerate XML preamble variation (with or without BOM, with or without
  XML declaration).
- We tolerate nonconforming-but-real-world `Relationships` root elements
  (e.g. files that include extra attributes — Microsoft Office occasionally
  adds them). The parser ignores anything except `<Relationship>` children.
- We do **not** validate that `Target` is well-formed URI. ECMA-376 says
  it should be (RFC 3986), but real files contain unencoded Windows paths
  and the patcher must round-trip them.

### 5.2 rId allocation (`add()`)

```text
fn add(rel_type, target, mode) -> RelId:
    let id = RelId(format!("rId{}", self.next_rid));
    self.next_rid += 1;
    self.rels.push(Relationship { id: id.clone(), rel_type, target, mode });
    id
```

**Monotonic counter, never decreases.** This is the key correctness fix
over openpyxl's `f"rId{len(self)}"` approach (see §3 bullet 2). Concrete
failure mode it prevents:

```
Initial:  rId1 rId2 rId3
Remove:   rId2  →  rels = [rId1, rId3]
                   openpyxl: next_id = "rId{len=2}" = "rId2"  (BUG: collides
                   with no longer existing rId2 if downstream re-references)
                   wolfxl:   next_rid = 4 → "rId4"
```

Real-world impact: Excel doesn't crash, but a defined-name pointing at the
old rId2 silently re-aims at the wrong target. Since modify mode preserves
unread parts byte-for-byte, any such stale reference would be live.

### 5.3 Serialization

```text
fn serialize(&self) -> Vec<u8>:
    let mut out = String::with_capacity(256 + self.rels.len() * 220);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    out.push_str("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    for r in &self.rels:
        out.push_str("<Relationship Id=\"");
        push_xml_attr_escape(&mut out, &r.id.0);
        out.push_str("\" Type=\"");
        out.push_str(&r.rel_type);     // URIs from rt:: are pre-validated; no escape needed
        out.push_str("\" Target=\"");
        push_xml_attr_escape(&mut out, &r.target);
        if r.mode == TargetMode::External:
            out.push_str("\" TargetMode=\"External");
        out.push_str("\"/>");
    out.push_str("</Relationships>");
    out.into_bytes()
```

Reuses the `xml_attr_escape` already in
`crates/wolfxl-writer/src/emit/rels.rs:67-80` — when we move the constants,
we move the helper too (or re-export it).

**Order preservation:** parsing pushes into a `Vec` in document order;
serialize iterates that `Vec`. No sort. Excel and LibreOffice both preserve
order; reordering trips file-diff tools and breaks the modify-mode
"minimal diff" invariant we sell users on.

**Determinism:** the same `RelsGraph` always serializes to the same bytes.
Critical for `WOLFXL_TEST_EPOCH=0` golden files
(`tests/diffwriter/`) and for the cross-surface parity harness in
`tests/test_classifier_parity.py`.

### 5.4 Idempotency

`parse(serialize(graph)) == graph` (structural equality on
`(rels, next_rid)`, modulo source-order being canonical).

`serialize(parse(bytes)) ≅ bytes` for *attribute content*. The byte-level
diff is bounded to:

- Whitespace between `<Relationship>` elements (we emit none; some files
  insert newlines).
- Self-closing form (`<Relationship .../>`) — we always use it; some files
  use `<Relationship ...></Relationship>`.
- Attribute order on `<Relationship>` — we always emit `Id`, `Type`,
  `Target`, [`TargetMode`]; some files use other orders.
- Encoding declaration whitespace — we always emit
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`.

Tests use a normalizer (parse-and-diff via `RelsGraph::eq`) for "semantic
identity" plus a separate byte-identical test against re-emitted output of
our own serializer (idempotency).

## 6. Test Plan

Standard verification matrix from the master plan §Verification, plus:

**Unit tests in `src/wolfxl/rels.rs`:**

1. `parse_minimal_root` — `_rels/.rels` from
   `tests/fixtures/minimal.xlsx` parses to 3 relationships
   (workbook, core, app).
2. `parse_workbook_rels_with_styles_and_sst` —
   `xl/_rels/workbook.xml.rels` from
   `tests/fixtures/tier1/01_cell_values.xlsx` parses, `find_by_type(rt::WORKSHEET)`
   returns exactly N entries where N == sheet count.
3. `parse_external_hyperlink_marks_external` —
   `tests/fixtures/tier2/13_hyperlinks.xlsx` has `TargetMode="External"` on
   the URL relationships; parser sets `mode == External`.
4. `add_after_parse_uses_strictly_greater_rid` — parse a file with
   max rId=5, call `add(...)` twice, assert returned ids are `rId6` and
   `rId7`.
5. `remove_then_add_does_not_collide` — parse file with rId1..rId3,
   `remove(rId2)`, `add(...)` returns `rId4` not `rId2`. Demonstrates the
   openpyxl bug fix.
6. `find_by_target_dedupe` — add same external hyperlink target twice;
   first call to `find_by_target` returns Some with the original id,
   confirming caller can dedupe.
7. `serialize_round_trips_idempotent` — `parse(serialize(parse(bytes))) ==
   parse(bytes)` for every file matched by:
   ```
   find tests/fixtures -name '*.rels' -type f
   find tests/fixtures -name '*.xlsx' | xargs unzip-and-extract-rels
   ```
8. `serialize_external_xml_escape` — target containing `&<>"'` round-trips
   through `xml_attr_escape` correctly (covers
   `https://example.com/path?q=1&r=2` already used in
   `crates/wolfxl-writer/src/emit/rels.rs:392-405`).
9. `numeric_suffix_handles_legacy_ids` — relationship with id
   `myId123` (no `rId` prefix) parses, `numeric_suffix()` returns None,
   `add()` still works (ignores legacy id when picking next).

**Integration tests:**

10. `tests/integration_rels_round_trip.rs` (new) — for each `.xlsx` in
    `tests/fixtures/`, open with `XlsxPatcher::open`, list every
    `*.rels` ZIP entry, parse each through `RelsGraph::parse`, serialize
    back, and assert `parse(serialize(...)) == original_graph`.
11. `tests/integration_rels_dedup.rs` (new) — open
    `tier2/13_hyperlinks.xlsx`, count distinct external targets in
    sheet1.xml.rels, assert no duplicate (Target, Type) pairs after a
    no-op patch+save cycle.

**Property test:** `proptest` 1k iterations of
random-(rel_type, target, mode) triples → `add` → `serialize` → `parse`
→ assert structural equality. Catches escape-handling regressions.

**Cross-surface parity:** after the writer migrates to use `RelsGraph`
internally, `cargo test -p wolfxl-writer emit::rels::tests` must continue
to pass without modification (the current 8 tests at
`emit/rels.rs:229-407` are the contract).

## 7. Migration / Compat Notes

- **No public Python API change.** Pure-internal refactor.
- **Writer migration is non-breaking.** The writer crate gets a new
  dependency on `src/wolfxl/rels.rs` (or, if we prefer the writer to be
  pyo3-free, we extract `rels.rs` into a `wolfxl-rels` workspace crate
  and both depend on it — see §8 risk #1).
- **Feature flag during rollout.** Not needed. The patcher doesn't
  produce rels output today (raw-copies only); turning that on per-RFC
  is the rollout. RFC-022/023/024 each gate behind their own readiness.
- **Behavior diff vs openpyxl:** removed-then-readded relationships get
  a new rId in wolfxl, the same rId in openpyxl. Documented in §3 as a
  *correctness* improvement, not a breaking change. There is no public
  surface in either library that exposes raw rIds.
- **Documented invariant for downstream RFCs:** "`RelsGraph::add` returns
  a fresh `RelId` you may immediately store in a `r:id="…"` attribute in
  any sibling part. Do not parse or interpret the numeric suffix —
  treat it as opaque." This goes in the module doc comment.

## 8. Risks & Open Questions

1. **(MED) Crate boundary for `wolfxl-rels`.** The writer
   (`crates/wolfxl-writer/`) is intentionally PyO3-free
   (CLAUDE.md "wolfxl-core carries no PyO3 dependency"). The patcher
   (`src/wolfxl/`) is inside the PyO3 cdylib crate. Putting the rels
   graph in `src/wolfxl/rels.rs` means the writer would have to reach
   into the PyO3 crate to use it, breaking the layering rule.
   **Resolution:** extract to a new `crates/wolfxl-rels/` workspace
   crate. Both `src/wolfxl/rels.rs` (a 1-line `pub use wolfxl_rels::*;`
   re-export for ergonomics) and `crates/wolfxl-writer/src/emit/rels.rs`
   depend on it. This adds one crate but is consistent with the
   `wolfxl-core` boundary already established.

2. **(LOW) Legacy rId conventions in real files.** The relationship.py
   docstring suggests Microsoft uses non-`rId<N>` ids in some templates
   (e.g. user-named ids from PowerPoint). Our `numeric_suffix()` returns
   `None` for those; `next_rid` only tracks ids that match the
   convention. **Resolution:** parse and preserve them verbatim; never
   emit a fresh id with their style. The first `add` after parsing such
   a file allocates `rId1` (because `next_rid` starts at 1 if no
   numeric ids were seen), which collides if any legacy id happens to
   be `rId1` written non-conventionally. Detect-and-bump in `parse()`:
   if any id literally equals `format!("rId{n}", n)` for any candidate
   `n`, treat it as numeric for `max_seen` purposes.

3. **(LOW) Parser strictness on missing `Type` or `Target`.** Real-world
   corrupt files exist. Today's `parse_relationship_targets`
   (`ooxml_util.rs:80-109`) silently skips missing-attribute entries.
   **Resolution:** match that behavior in `RelsGraph::parse` — log a
   warning via `tracing::warn!` if either is missing, but skip the entry
   rather than failing the whole graph. Add a `parse_strict()` variant
   for tests that demands all attributes present.

4. **(LOW) Whitespace between `<Relationship>` elements in input.**
   Some files indent each Relationship on its own line. Our serializer
   emits one continuous line. The diff is cosmetic and doesn't affect
   Excel's parse, but breaks `git diff` cleanliness for users
   round-tripping a file. **Resolution:** punt to a future
   `pretty: bool` flag on `serialize()`. Default behavior remains
   compact for byte-stable output.

5. **(MED) Concurrent rels mutation.** If the patcher queues both a
   hyperlink (RFC-022) and a comment (RFC-023) on the same sheet, the
   per-sheet rels graph is mutated twice. Allocation order matters
   because rId values depend on what was allocated first. **Resolution:**
   the patcher's save loop owns the `RelsGraph` and passes a `&mut`
   reference to each sub-RFC's emit function in a deterministic order
   (alphabetical by RFC number — comments before hyperlinks before tables).
   Document this order in `src/wolfxl/mod.rs`.

## 9. Effort Breakdown

| Task | LOC est. | Days |
|---|---|---|
| `src/wolfxl/rels.rs` (or `crates/wolfxl-rels/`) — module + parser + serializer + RelsGraph | 280 | 1.0 |
| Move `RT_*` constants out of `wolfxl-writer/src/emit/rels.rs`, refactor emit_* to use `RelsGraph` | 150 (net -50) | 0.5 |
| Unit tests (9 above) | 220 | 0.5 |
| Integration tests (round-trip + dedup) | 120 | 0.3 |
| Property test (proptest) | 60 | 0.2 |
| Wiring into `src/wolfxl/mod.rs` save loop (no-op until RFC-022 lands) | 25 | 0.2 |
| **Total** | **~855** | **~2.7 days** |

Estimate band: M (≤ 3 days). No surprises expected — algorithm is well-trodden
and openpyxl already validates the spec interpretation.

## 10. Out of Scope

The following are NOT in this RFC and must not be implemented as a
side-effect:

- Rels target path **normalization** (relative-to-absolute resolution).
  Patcher round-trips `Target` verbatim; absolute resolution is the caller's
  problem. (See §3 "what we will NOT copy" bullet 2.)
- **`[Content_Types].xml`** is not a `.rels` file. RFC-013 (separate)
  handles content-type registry mutation.
- Rels validation against the registered relationship-type vocabulary.
  We accept any URI in `Type` even if it's not in our `rt::` constants.
- **External hyperlink dedup heuristics** beyond exact target+mode match.
  RFC-022 may layer a smarter dedupe (case-insensitive, trailing-slash-
  agnostic) on top of `find_by_target`.
- **Relationships in macro-enabled (.xlsm) `vbaProject` parts.** Modify
  mode preserves these byte-for-byte; we never mutate them.
- **Pretty-printing / indentation** of rels files. See risk #4.
- **Concurrent multi-thread mutation.** The graph is `!Sync` (it doesn't
  need to be — patcher work is single-threaded inside `do_save`).

## Acceptance: shipped via 6-commit slice on `feat/native-writer` (through 309554d) on 2026-04-25

Crate `crates/wolfxl-rels/` extracted as a workspace member; writer's
`emit/rels.rs`, the reader's `parse_relationship_targets`, and the patcher's
save loop all consume one implementation. 19 new tests in `wolfxl-rels` (15
unit + 2 dedup integration + 1 round-trip integration + 1 proptest);
9 existing emit tests in `wolfxl-writer` pass unchanged; 124 reader compat
tests + 12 modify-mode tests pass. The patcher's `rels_patches` field is
plumbed but dead code at HEAD — RFC-022/023/024 will populate it.
