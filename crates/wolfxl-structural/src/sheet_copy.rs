//! `plan_sheet_copy` — pure planner for `Workbook.copy_worksheet` (RFC-035
//! Phase 7.2).
//!
//! Lives in `wolfxl-structural` (PyO3-free) so its tests can run via
//! `cargo test -p wolfxl-structural` without touching the cdylib's link
//! issue. The patcher (`src/wolfxl/sheet_copy.rs`) is a thin re-export
//! over this module per the spec's API location requirement.
//!
//! See `Plans/rfcs/035-copy-worksheet.md` §4.2, §5.1–§5.7 for the
//! authoritative contract.
//!
//! # Responsibilities
//!
//! 1. Walk the source sheet's rels graph (one level for built-in part
//!    types; nested resolver for drawings → images).
//! 2. Allocate fresh suffixes for each cloneable ancillary part via the
//!    shared [`PartIdAllocator`].
//! 3. Build the `(old_rid → new_rid)` remap and apply it as a SINGLE
//!    pass over the cloned sheet XML (RFC-035 §8 risk #4).
//! 4. Clone each ancillary part's bytes (table parts get name dedup +
//!    re-emitted bytes; comments / VML / drawings flow through verbatim
//!    in this slice).
//! 5. Scan workbook.xml for sheet-scoped defined names whose
//!    `localSheetId == src_idx` and clone each with `localSheetId ==
//!    new_idx` (RFC-035 §5.4 / OQ-c default).
//! 6. Validate (RFC-035 §5.9): error if `dst_title` already exists or if
//!    `src_title` is missing from the sheet list.

use std::collections::{HashMap, HashSet};

use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

use wolfxl_rels::{
    rt, walk_sheet_subgraph_with_nested, PartIdAllocator, RelsGraph, TargetMode,
};

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// Inputs for one sheet-copy planning call. Borrows everything the
/// planner needs from the patcher's current view of the workbook —
/// including `source_zip_parts`, which the caller pre-loads (the
/// planner does not perform ZIP I/O).
pub struct SheetCopyInputs<'a> {
    /// Source sheet title (for error messages + workbook.xml lookup).
    pub src_title: String,
    /// Destination sheet title (for the new `<sheet>` entry). Must be
    /// pre-deduped by the caller.
    pub dst_title: String,
    /// Source sheet ZIP part path, e.g. `xl/worksheets/sheet3.xml`.
    pub src_sheet_path: String,
    /// Pre-loaded source ZIP parts: `(path, bytes)`. Includes the
    /// sheet itself, every table/comments/VML/drawing part it
    /// references, and any nested-rels parts (drawing's own rels file
    /// for the image case).
    pub source_zip_parts: &'a HashMap<String, Vec<u8>>,
    /// The source sheet's parsed rels graph
    /// (`xl/worksheets/_rels/sheet<src>.xml.rels`). May be empty.
    pub source_rels: &'a RelsGraph,
    /// `xl/workbook.xml` bytes — used to scan sheet-scoped
    /// `<definedName>` entries and to compute `src_idx`.
    pub workbook_xml: &'a [u8],
    /// Centralized part-suffix allocator. Mutated by the planner.
    pub allocator: &'a mut PartIdAllocator,
    /// Workbook-wide set of table `name`/`displayName` values that are
    /// already taken (for §5.5 dedup).
    pub existing_table_names: &'a HashSet<String>,
    /// Sprint Θ Pod-C2 — when `true`, drawings reachable from the
    /// source sheet have their referenced `xl/media/imageN.<ext>`
    /// targets DEEP-CLONED into freshly numbered media parts. The
    /// cloned drawing's nested rels file is rewritten to point at
    /// the new paths and the image bytes are added to
    /// `new_ancillary_parts`.
    ///
    /// The default (`false`) preserves the historical RFC-035 §5.3
    /// alias-by-target behaviour.
    pub deep_copy_images: bool,
}

/// One cloned defined-name (RFC-035 §5.4). The planner returns the
/// XML bytes ready to splice into `<definedNames>`; the caller
/// concatenates and merges.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DefinedNameClone {
    /// `name` attribute, verbatim from the source defined name.
    pub name: String,
    /// `localSheetId`, recomputed to point at the new sheet's
    /// position in the post-copy tab list.
    pub local_sheet_id: u32,
    /// Formula text (verbatim from source — sheet copies share
    /// coordinate space, so cell references resolve correctly without
    /// translation).
    pub formula: String,
    /// Pre-rendered `<definedName …>…</definedName>` element bytes
    /// for direct splicing.
    pub element_bytes: Vec<u8>,
}

/// All mutations the planner produces for one sheet copy. The caller
/// (Phase 2.7 of `XlsxPatcher::do_save`) applies these atomically.
#[derive(Debug, Clone, Default)]
pub struct SheetCopyMutations {
    /// New sheet's ZIP part path, e.g. `xl/worksheets/sheet5.xml`.
    pub new_sheet_path: String,
    /// New sheet's XML bytes (source bytes with the `r:id` remap
    /// applied).
    pub new_sheet_xml: Vec<u8>,
    /// `(zip_path, bytes)` for every cloned ancillary part: tables,
    /// comments, VML, drawings, drawing rels files (when present).
    /// Image media is NOT in this list — images are aliased per
    /// RFC-035 §5.3 / §8 risk #2.
    pub new_ancillary_parts: Vec<(String, Vec<u8>)>,
    /// `(part_path, content_type)` overrides to add to
    /// `[Content_Types].xml`. `part_path` is leading-slashed per OPC
    /// convention (`"/xl/tables/table7.xml"`).
    pub content_type_overrides_to_add: Vec<(String, String)>,
    /// `<sheet name="…" sheetId="…" r:id="…"/>` element bytes for the
    /// caller to splice into workbook.xml's `<sheets>` block. The
    /// caller fills in `r:id` AFTER allocating the workbook-rels
    /// entry (we don't know the workbook rels graph here).
    pub workbook_sheets_append: Vec<u8>,
    /// `(rId-placeholder, rel_type, target)` for the workbook rels
    /// entry the caller must add. The `rId-placeholder` is the same
    /// string used in `workbook_sheets_append` so the caller can do a
    /// single replace-all pass after allocating the actual rId.
    pub workbook_rels_to_add: Vec<(String, String, String)>,
    /// Sheet-scoped defined-name clones (RFC-035 §5.4).
    pub defined_names_to_add: Vec<DefinedNameClone>,
    /// Newly allocated table `name`/`displayName` values added to
    /// `existing_table_names` for collision detection across multiple
    /// in-flight copies. The caller folds these into the running set
    /// before the next `plan_sheet_copy` call.
    pub new_table_names: Vec<String>,
    /// `old_rid → new_rid` for the source sheet's local rels graph.
    /// Useful for callers that want to do additional rewrites on the
    /// cloned sheet bytes; the planner has already applied this remap
    /// to `new_sheet_xml`.
    pub rid_remap: HashMap<String, String>,
}

/// Errors `plan_sheet_copy` may return.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SheetCopyError {
    /// `dst_title` already names an existing sheet in workbook.xml.
    DuplicateDestinationTitle(String),
    /// `src_title` was not found in workbook.xml's `<sheets>` list.
    MissingSourceTitle(String),
    /// Source sheet bytes were not in `source_zip_parts`.
    MissingSourceSheetBytes(String),
    /// XML parse / serialization error.
    Xml(String),
}

impl std::fmt::Display for SheetCopyError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            SheetCopyError::DuplicateDestinationTitle(s) => {
                write!(f, "destination sheet title {s:?} already exists")
            }
            SheetCopyError::MissingSourceTitle(s) => {
                write!(f, "source sheet title {s:?} not found in workbook.xml")
            }
            SheetCopyError::MissingSourceSheetBytes(s) => {
                write!(f, "source sheet bytes for {s:?} not provided")
            }
            SheetCopyError::Xml(e) => write!(f, "xml error: {e}"),
        }
    }
}

impl std::error::Error for SheetCopyError {}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_TABLE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const CT_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";

// ---------------------------------------------------------------------------
// Planner entry point
// ---------------------------------------------------------------------------

/// Plan one sheet copy. Pure: does not mutate any caller-owned data
/// other than `inputs.allocator` (which is `&mut`).
pub fn plan_sheet_copy(
    inputs: SheetCopyInputs<'_>,
) -> Result<SheetCopyMutations, SheetCopyError> {
    // ---- Validate ----------------------------------------------------------
    let SheetMetadata {
        sheet_titles,
        defined_names,
    } = parse_workbook_metadata(inputs.workbook_xml)?;

    let src_idx = sheet_titles
        .iter()
        .position(|t| t == &inputs.src_title)
        .ok_or_else(|| SheetCopyError::MissingSourceTitle(inputs.src_title.clone()))?
        as u32;

    if sheet_titles.iter().any(|t| t == &inputs.dst_title) {
        return Err(SheetCopyError::DuplicateDestinationTitle(
            inputs.dst_title.clone(),
        ));
    }

    let new_idx = sheet_titles.len() as u32; // appended at end

    let src_sheet_xml = inputs
        .source_zip_parts
        .get(&inputs.src_sheet_path)
        .ok_or_else(|| SheetCopyError::MissingSourceSheetBytes(inputs.src_sheet_path.clone()))?;

    // ---- Walk the rels subgraph -------------------------------------------
    // Resolver: for any reachable part we hit, look up its `_rels` file (if
    // any) in source_zip_parts and parse it. This catches drawings → images
    // and chart parts → media.
    let zip_parts = inputs.source_zip_parts;
    let subgraph = walk_sheet_subgraph_with_nested(
        inputs.source_rels,
        &inputs.src_sheet_path,
        |part_path| {
            let rels_path = wolfxl_rels::rels_path_for(part_path)?;
            let bytes = zip_parts.get(&rels_path)?;
            RelsGraph::parse(bytes).ok()
        },
    );

    // ---- Allocate destination part suffixes + build rid_remap -------------
    let new_sheet_n = inputs.allocator.alloc_sheet();
    let new_sheet_path = format!("xl/worksheets/sheet{new_sheet_n}.xml");

    // Build the destination rels graph alongside the rid_remap. We assign
    // freshly monotonic rIds in the new graph so the remap is stable. The
    // destination rels graph is NOT serialized here — the caller (Phase 2.7)
    // wires it into the patcher's `rels_patches` map.
    let mut dest_rels = RelsGraph::new();
    let mut rid_remap: HashMap<String, String> = HashMap::new();

    let mut new_ancillary_parts: Vec<(String, Vec<u8>)> = Vec::new();
    let mut content_type_overrides_to_add: Vec<(String, String)> = Vec::new();
    let mut new_table_names: Vec<String> = Vec::new();

    // For tracking the "set already used in this session" to dedup tables
    // when one copy clones two tables with the same base name.
    let mut taken_names: HashSet<String> = inputs.existing_table_names.clone();

    // Cache of per-part rels graphs keyed by part path so we don't re-parse.
    let mut nested_rels_cache: HashMap<String, RelsGraph> = HashMap::new();
    for part_path in &subgraph.reachable_parts {
        if let Some(rp) = wolfxl_rels::rels_path_for(part_path) {
            if let Some(bytes) = zip_parts.get(&rp) {
                if let Ok(g) = RelsGraph::parse(bytes) {
                    nested_rels_cache.insert(part_path.clone(), g);
                }
            }
        }
    }

    // Walk the source sheet's edges in source order and clone each part.
    for source_rel in inputs.source_rels.iter() {
        let old_rid = source_rel.id.0.clone();
        match source_rel.rel_type.as_str() {
            // ------ Tables (need name dedup + bytes re-emit) ------
            t if t == rt::TABLE => {
                let resolved = resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
                let new_n = inputs.allocator.alloc_table();
                let new_part_path = format!("xl/tables/table{new_n}.xml");
                // Clone bytes with name + id dedup.
                let src_bytes = zip_parts.get(&resolved).cloned().unwrap_or_default();
                let (cloned_bytes, new_name) =
                    clone_table_part(&src_bytes, &mut taken_names, new_n)?;
                new_ancillary_parts.push((new_part_path.clone(), cloned_bytes));
                new_table_names.push(new_name.clone());
                taken_names.insert(new_name);
                content_type_overrides_to_add
                    .push((format!("/{}", new_part_path), CT_TABLE.into()));
                let new_target = format!("../tables/table{new_n}.xml");
                let new_rid = dest_rels.add(rt::TABLE, &new_target, TargetMode::Internal);
                rid_remap.insert(old_rid, new_rid.0);
            }
            // ------ Comments ------
            t if t == rt::COMMENTS => {
                let resolved = resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
                let new_n = inputs.allocator.alloc_comments();
                let new_part_path = format!("xl/comments{new_n}.xml");
                let src_bytes = zip_parts.get(&resolved).cloned().unwrap_or_default();
                new_ancillary_parts.push((new_part_path.clone(), src_bytes));
                content_type_overrides_to_add
                    .push((format!("/{}", new_part_path), CT_COMMENTS.into()));
                let new_target = format!("../comments{new_n}.xml");
                let new_rid = dest_rels.add(rt::COMMENTS, &new_target, TargetMode::Internal);
                rid_remap.insert(old_rid, new_rid.0);
            }
            // ------ VML drawings ------
            t if t == rt::VML_DRAWING => {
                let resolved = resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
                let new_n = inputs.allocator.alloc_vml_drawing();
                let new_part_path = format!("xl/drawings/vmlDrawing{new_n}.vml");
                let src_bytes = zip_parts.get(&resolved).cloned().unwrap_or_default();
                new_ancillary_parts.push((new_part_path.clone(), src_bytes));
                // VML uses a Default content-type by extension; not an
                // Override, but we still surface it so the caller can
                // ensure_default("vml", ...) at the content-types layer.
                let new_target = format!("../drawings/vmlDrawing{new_n}.vml");
                let new_rid = dest_rels.add(rt::VML_DRAWING, &new_target, TargetMode::Internal);
                rid_remap.insert(old_rid, new_rid.0);
            }
            // ------ DrawingML drawings (own rels file → image alias OR deep-clone) ------
            t if t == rt::DRAWING => {
                let resolved = resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
                let new_n = inputs.allocator.alloc_drawing();
                let new_part_path = format!("xl/drawings/drawing{new_n}.xml");
                let src_bytes = zip_parts.get(&resolved).cloned().unwrap_or_default();
                new_ancillary_parts.push((new_part_path.clone(), src_bytes));
                content_type_overrides_to_add
                    .push((format!("/{}", new_part_path), CT_DRAWING.into()));
                let new_target = format!("../drawings/drawing{new_n}.xml");
                let new_rid = dest_rels.add(rt::DRAWING, &new_target, TargetMode::Internal);
                rid_remap.insert(old_rid, new_rid.0);
                // Clone the drawing's own rels file IF present.
                //
                // - Default (alias) mode: the image target stays
                //   pointed at the same xl/media/imageN.{ext} path
                //   (RFC-035 §5.3).
                // - Deep-copy mode (Sprint Θ Pod-C2): allocate a
                //   fresh `xl/media/imageM.<ext>` per image rel,
                //   copy its bytes into `new_ancillary_parts`, and
                //   re-point the cloned drawing rel at the new
                //   target.
                if let Some(nested) = nested_rels_cache.get(&resolved) {
                    let mut cloned = RelsGraph::new();
                    for nrel in nested.iter() {
                        if inputs.deep_copy_images && nrel.rel_type == rt::IMAGE {
                            let resolved_image = resolve_relative(
                                parent_dir(&resolved),
                                &nrel.target,
                            );
                            // Determine the original extension; fall
                            // back to "png" if the original had none
                            // (rare in practice but keeps us safe).
                            let ext = resolved_image
                                .rsplit_once('.')
                                .map(|(_, e)| e.to_string())
                                .unwrap_or_else(|| "png".to_string());
                            let new_image_n = inputs.allocator.alloc_image();
                            let new_image_path =
                                format!("xl/media/image{new_image_n}.{ext}");
                            // Copy the original bytes if present.
                            if let Some(bytes) = zip_parts.get(&resolved_image).cloned() {
                                new_ancillary_parts.push((new_image_path.clone(), bytes));
                            }
                            // The rel's `target` is drawing-relative
                            // ("../media/imageN.<ext>"); we rewrite
                            // it to the new suffix.
                            let new_rel_target =
                                format!("../media/image{new_image_n}.{ext}");
                            cloned.add(&nrel.rel_type, &new_rel_target, nrel.mode);
                        } else {
                            cloned.add(&nrel.rel_type, &nrel.target, nrel.mode);
                        }
                    }
                    let nested_rels_path = format!(
                        "xl/drawings/_rels/drawing{new_n}.xml.rels"
                    );
                    new_ancillary_parts.push((nested_rels_path, cloned.serialize()));
                }
            }
            // ------ Charts (clone bytes verbatim, fresh suffix) ------
            t if t == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
                || t == rt::PIVOT_TABLE =>
            {
                // Out-of-scope for full re-pointing per §10. Clone-by-alias:
                // reuse the source target verbatim so the new sheet shares
                // the chart with the source. (Same reasoning as image
                // aliasing.) No new ancillary part emitted.
                let new_rid =
                    dest_rels.add(&source_rel.rel_type, &source_rel.target, source_rel.mode);
                rid_remap.insert(old_rid, new_rid.0);
            }
            // ------ Hyperlinks (External: alias by URL) ------
            t if t == rt::HYPERLINK => {
                let new_rid =
                    dest_rels.add(&source_rel.rel_type, &source_rel.target, source_rel.mode);
                rid_remap.insert(old_rid, new_rid.0);
            }
            // ------ Anything else (alias the target verbatim) ------
            _ => {
                let new_rid =
                    dest_rels.add(&source_rel.rel_type, &source_rel.target, source_rel.mode);
                rid_remap.insert(old_rid, new_rid.0);
            }
        }
    }

    // Emit the destination rels file alongside the cloned sheet so the
    // caller can drop both into file_adds in one shot.
    if !dest_rels.is_empty() {
        let dest_rels_path = wolfxl_rels::rels_path_for(&new_sheet_path)
            .unwrap_or_else(|| format!("_rels/{new_sheet_path}.rels"));
        new_ancillary_parts.push((dest_rels_path, dest_rels.serialize()));
    }

    // ---- Apply the rId remap to the cloned sheet XML in one pass ----------
    let new_sheet_xml = rewrite_rids(src_sheet_xml, &rid_remap)?;

    // ---- Sheet-scoped defined-name clones ---------------------------------
    let defined_names_to_add: Vec<DefinedNameClone> = defined_names
        .iter()
        .filter(|dn| dn.local_sheet_id == Some(src_idx))
        .map(|dn| {
            let element_bytes = render_defined_name(&dn.name, new_idx, &dn.formula);
            DefinedNameClone {
                name: dn.name.clone(),
                local_sheet_id: new_idx,
                formula: dn.formula.clone(),
                element_bytes,
            }
        })
        .collect();

    // ---- New <sheet> entry for workbook.xml -------------------------------
    let placeholder_rid = format!("__SHEET_RID_PLACEHOLDER_{new_sheet_n}__");
    let workbook_sheets_append = render_sheet_entry(&inputs.dst_title, new_sheet_n, &placeholder_rid);
    let workbook_rels_to_add = vec![(
        placeholder_rid,
        rt::WORKSHEET.to_string(),
        format!("worksheets/sheet{new_sheet_n}.xml"),
    )];

    // Add the worksheet content-type override.
    content_type_overrides_to_add
        .insert(0, (format!("/{new_sheet_path}"), CT_WORKSHEET.into()));

    // Avoid `unused`: subgraph is informational; we acted on inputs.source_rels
    // directly. Keep a debug-friendly check that the sheet itself was visited.
    debug_assert_eq!(
        subgraph.reachable_parts.first().map(|s| s.as_str()),
        Some(inputs.src_sheet_path.as_str())
    );

    Ok(SheetCopyMutations {
        new_sheet_path,
        new_sheet_xml,
        new_ancillary_parts,
        content_type_overrides_to_add,
        workbook_sheets_append,
        workbook_rels_to_add,
        defined_names_to_add,
        new_table_names,
        rid_remap,
    })
}

// ---------------------------------------------------------------------------
// Helpers — workbook.xml metadata parse
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Default)]
struct SheetMetadata {
    /// Sheet titles in document order (the position is the
    /// `localSheetId`).
    sheet_titles: Vec<String>,
    defined_names: Vec<ParsedDefinedName>,
}

#[derive(Debug, Clone)]
struct ParsedDefinedName {
    name: String,
    /// `Some(idx)` if scoped to a sheet position; `None` if
    /// workbook-scoped.
    local_sheet_id: Option<u32>,
    /// Formula text (verbatim, unescaped).
    formula: String,
}

fn parse_workbook_metadata(workbook_xml: &[u8]) -> Result<SheetMetadata, SheetCopyError> {
    let mut reader = XmlReader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();
    let mut meta = SheetMetadata::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                match e.local_name().as_ref() {
                    b"sheet" => {
                        let name = attr_value(&e, b"name").unwrap_or_default();
                        meta.sheet_titles.push(name);
                    }
                    b"definedName" => {
                        let name = attr_value(&e, b"name").unwrap_or_default();
                        let lsid = attr_value(&e, b"localSheetId")
                            .and_then(|v| v.parse::<u32>().ok());
                        meta.defined_names.push(ParsedDefinedName {
                            name,
                            local_sheet_id: lsid,
                            formula: String::new(),
                        });
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(SheetCopyError::Xml(format!("workbook.xml: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    // Second pass to populate formulas — quick-xml's nested-event API makes
    // a single pass tricky for `<definedName>foo</definedName>` text nodes.
    populate_defined_name_formulas(workbook_xml, &mut meta.defined_names)?;

    Ok(meta)
}

fn populate_defined_name_formulas(
    workbook_xml: &[u8],
    defined_names: &mut [ParsedDefinedName],
) -> Result<(), SheetCopyError> {
    let mut reader = XmlReader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();
    let mut in_dn: Option<usize> = None;
    let mut idx = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"definedName" => {
                in_dn = Some(idx);
            }
            Ok(Event::Text(t)) if in_dn.is_some() => {
                let text = t
                    .unescape()
                    .map(|s| s.into_owned())
                    .unwrap_or_else(|_| String::from_utf8_lossy(&t).into_owned());
                if let Some(i) = in_dn {
                    if let Some(dn) = defined_names.get_mut(i) {
                        dn.formula.push_str(&text);
                    }
                }
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"definedName" => {
                in_dn = None;
                idx += 1;
            }
            Ok(Event::Empty(e)) if e.local_name().as_ref() == b"definedName" => {
                // Self-closing definedName has no formula text.
                idx += 1;
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(SheetCopyError::Xml(format!("dn formula: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(())
}

fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for a in e.attributes().with_checks(false).flatten() {
        if a.key.as_ref() == key {
            return Some(
                a.unescape_value()
                    .map(|v| v.into_owned())
                    .unwrap_or_else(|_| String::from_utf8_lossy(a.value.as_ref()).into_owned()),
            );
        }
    }
    None
}

// ---------------------------------------------------------------------------
// Helpers — table clone + name dedup
// ---------------------------------------------------------------------------

/// Per RFC-035 §5.5 / OQ-b default:
/// `f"{base}_{N}"` starting at N=2.
fn dedup_table_name(base: &str, taken: &HashSet<String>) -> String {
    if !taken.contains(base) {
        // Even if the base isn't taken, RFC-035's contract is: every
        // CLONED table gets a non-source name. We always append `_2`+ to
        // make the divergence loud.
        let mut suffix = 2u32;
        let mut candidate = format!("{base}_{suffix}");
        while taken.contains(&candidate) {
            suffix += 1;
            candidate = format!("{base}_{suffix}");
        }
        return candidate;
    }
    let mut suffix = 2u32;
    let mut candidate = format!("{base}_{suffix}");
    while taken.contains(&candidate) {
        suffix += 1;
        candidate = format!("{base}_{suffix}");
    }
    candidate
}

/// Re-emit a `tableN.xml` part with: (1) a deduped `name` /
/// `displayName`; (2) a fresh `id` derived from `new_n` (workbook-unique
/// suffix is sufficient since callers track every emitted id via the
/// allocator). Other attributes flow through verbatim.
fn clone_table_part(
    src_bytes: &[u8],
    taken_names: &mut HashSet<String>,
    new_n: u32,
) -> Result<(Vec<u8>, String), SheetCopyError> {
    let mut reader = XmlReader::from_reader(src_bytes);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(src_bytes.len()));
    let mut buf: Vec<u8> = Vec::new();
    let mut renamed: Option<String> = None;
    let mut original_name: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Decl(d)) => {
                writer
                    .write_event(Event::Decl(d))
                    .map_err(|e| SheetCopyError::Xml(format!("table decl: {e}")))?;
            }
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if e.local_name().as_ref() == b"table" =>
            {
                let mut new_e = BytesStart::new(
                    std::str::from_utf8(e.name().as_ref())
                        .map_err(|er| SheetCopyError::Xml(format!("table tag utf8: {er}")))?
                        .to_owned(),
                );
                let mut want_name = true;
                let mut want_display = true;
                let mut want_id = true;
                for a in e.attributes().with_checks(false).flatten() {
                    let key = a.key.as_ref().to_vec();
                    let value = a
                        .unescape_value()
                        .map(|v| v.into_owned())
                        .unwrap_or_else(|_| String::from_utf8_lossy(a.value.as_ref()).into_owned());
                    match key.as_slice() {
                        b"name" => {
                            original_name = Some(value.clone());
                            let new_name = dedup_table_name(&value, taken_names);
                            renamed = Some(new_name.clone());
                            new_e.push_attribute(Attribute {
                                key: QName(b"name"),
                                value: new_name.into_bytes().into(),
                            });
                            want_name = false;
                        }
                        b"displayName" => {
                            // Mirror the deduped name onto displayName too
                            // (per ECMA-376, they share the uniqueness
                            // constraint).
                            let target = renamed.clone().unwrap_or_else(|| {
                                // We may hit displayName before name in
                                // attribute order; defer until after.
                                value.clone()
                            });
                            new_e.push_attribute(Attribute {
                                key: QName(b"displayName"),
                                value: target.into_bytes().into(),
                            });
                            want_display = false;
                        }
                        b"id" => {
                            new_e.push_attribute(Attribute {
                                key: QName(b"id"),
                                value: new_n.to_string().into_bytes().into(),
                            });
                            want_id = false;
                        }
                        _ => {
                            new_e.push_attribute(Attribute {
                                key: QName(&key),
                                value: value.into_bytes().into(),
                            });
                        }
                    }
                }
                if want_id {
                    new_e.push_attribute(Attribute {
                        key: QName(b"id"),
                        value: new_n.to_string().into_bytes().into(),
                    });
                }
                if want_name {
                    let nm = renamed.clone().unwrap_or_else(|| format!("Table{new_n}"));
                    new_e.push_attribute(Attribute {
                        key: QName(b"name"),
                        value: nm.into_bytes().into(),
                    });
                }
                if want_display {
                    let nm = renamed.clone().unwrap_or_else(|| format!("Table{new_n}"));
                    new_e.push_attribute(Attribute {
                        key: QName(b"displayName"),
                        value: nm.into_bytes().into(),
                    });
                }
                let ev = match reader.read_event_into(&mut Vec::new()) {
                    // We pre-fetched the next event into another buffer to
                    // detect Empty vs Start; revert to writing as the same
                    // kind we saw in the outer match.
                    _ => Event::Empty(new_e.clone()),
                };
                let _ = ev; // unused — we always rewrite as a Start here
                writer
                    .write_event(if buf.is_empty() {
                        Event::Empty(new_e.clone())
                    } else {
                        Event::Start(new_e)
                    })
                    .map_err(|e| SheetCopyError::Xml(format!("table root: {e}")))?;
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                writer
                    .write_event(other)
                    .map_err(|e| SheetCopyError::Xml(format!("table flow: {e}")))?;
            }
            Err(e) => return Err(SheetCopyError::Xml(format!("table parse: {e}"))),
        }
        buf.clear();
    }

    let bytes = writer.into_inner();
    let new_name = renamed.unwrap_or_else(|| {
        // Source had no name attribute — should never happen in real OOXML
        // but handle gracefully.
        let synth = format!("Table{new_n}");
        taken_names.insert(synth.clone());
        synth
    });
    let _ = original_name;
    Ok((bytes, new_name))
}

// ---------------------------------------------------------------------------
// Helpers — sheet-XML rId remap pass
// ---------------------------------------------------------------------------

/// Apply an `(old_rid → new_rid)` remap to every `r:id="…"` attribute
/// in the source sheet XML. SINGLE pass (RFC-035 §8 risk #4 mitigation)
/// — we walk every element and rewrite any attribute whose local name
/// is `id` AND whose namespace prefix is `r` (or whose qualified name
/// is `r:id`). Other attributes (and element text) flow through
/// verbatim.
fn rewrite_rids(
    src_xml: &[u8],
    remap: &HashMap<String, String>,
) -> Result<Vec<u8>, SheetCopyError> {
    if remap.is_empty() {
        return Ok(src_xml.to_vec());
    }
    let mut reader = XmlReader::from_reader(src_xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(src_xml.len()));
    let mut buf: Vec<u8> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let new_e = remap_element_attributes(&e, remap)?;
                writer
                    .write_event(Event::Start(new_e))
                    .map_err(|er| SheetCopyError::Xml(format!("write Start: {er}")))?;
            }
            Ok(Event::Empty(e)) => {
                let new_e = remap_element_attributes(&e, remap)?;
                writer
                    .write_event(Event::Empty(new_e))
                    .map_err(|er| SheetCopyError::Xml(format!("write Empty: {er}")))?;
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                writer
                    .write_event(other)
                    .map_err(|er| SheetCopyError::Xml(format!("write flow: {er}")))?;
            }
            Err(e) => return Err(SheetCopyError::Xml(format!("rid remap parse: {e}"))),
        }
        buf.clear();
    }
    Ok(writer.into_inner())
}

fn remap_element_attributes<'a>(
    e: &BytesStart<'a>,
    remap: &HashMap<String, String>,
) -> Result<BytesStart<'static>, SheetCopyError> {
    let mut new_e = BytesStart::new(
        std::str::from_utf8(e.name().as_ref())
            .map_err(|er| SheetCopyError::Xml(format!("element name utf8: {er}")))?
            .to_owned(),
    );
    for a in e.attributes().with_checks(false).flatten() {
        let key = a.key.as_ref().to_vec();
        let value = a
            .unescape_value()
            .map(|v| v.into_owned())
            .unwrap_or_else(|_| String::from_utf8_lossy(a.value.as_ref()).into_owned());
        let is_r_id = key == b"r:id" || (key.ends_with(b":id") && key.starts_with(b"r"));
        let new_value = if is_r_id {
            remap.get(&value).cloned().unwrap_or(value)
        } else {
            value
        };
        new_e.push_attribute(Attribute {
            key: QName(&key),
            value: new_value.into_bytes().into(),
        });
    }
    Ok(new_e)
}

// ---------------------------------------------------------------------------
// Helpers — defined-name / sheet-entry rendering
// ---------------------------------------------------------------------------

fn render_defined_name(name: &str, local_sheet_id: u32, formula: &str) -> Vec<u8> {
    let mut out = String::with_capacity(64 + name.len() + formula.len());
    out.push_str("<definedName name=\"");
    push_xml_attr_escape(&mut out, name);
    out.push_str(&format!("\" localSheetId=\"{local_sheet_id}\">"));
    push_xml_text_escape(&mut out, formula);
    out.push_str("</definedName>");
    out.into_bytes()
}

fn render_sheet_entry(name: &str, sheet_n: u32, rid: &str) -> Vec<u8> {
    let mut out = String::with_capacity(64 + name.len());
    out.push_str("<sheet name=\"");
    push_xml_attr_escape(&mut out, name);
    out.push_str(&format!("\" sheetId=\"{sheet_n}\" r:id=\"{rid}\"/>"));
    out.into_bytes()
}

fn push_xml_attr_escape(out: &mut String, s: &str) {
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
}

fn push_xml_text_escape(out: &mut String, s: &str) {
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(ch),
        }
    }
}

// ---------------------------------------------------------------------------
// Helpers — relative path resolution (mirrors subgraph_walk's helper)
// ---------------------------------------------------------------------------

fn parent_dir(part_path: &str) -> String {
    match part_path.rfind('/') {
        Some(idx) => part_path[..idx].to_string(),
        None => String::new(),
    }
}

fn resolve_relative(base_dir: String, target: &str) -> String {
    if let Some(stripped) = target.strip_prefix('/') {
        return stripped.to_string();
    }
    let mut segments: Vec<&str> = if base_dir.is_empty() {
        Vec::new()
    } else {
        base_dir.split('/').collect()
    };
    for part in target.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            other => segments.push(other),
        }
    }
    segments.join("/")
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    fn minimal_sheet_xml() -> Vec<u8> {
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/></worksheet>"#.to_vec()
    }

    fn sheet_xml_with_table_part(rid: &str) -> Vec<u8> {
        format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><tableParts count="1"><tablePart r:id="{rid}"/></tableParts></worksheet>"#).into_bytes()
    }

    fn workbook_xml(titles: &[&str], defined_names: &[(&str, Option<u32>, &str)]) -> Vec<u8> {
        let mut s = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>"#,
        );
        for (i, t) in titles.iter().enumerate() {
            s.push_str(&format!(
                r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
                t,
                i + 1,
                i + 1
            ));
        }
        s.push_str("</sheets>");
        if !defined_names.is_empty() {
            s.push_str("<definedNames>");
            for (n, lsid, f) in defined_names {
                match lsid {
                    Some(id) => s.push_str(&format!(
                        r#"<definedName name="{n}" localSheetId="{id}">{f}</definedName>"#
                    )),
                    None => s.push_str(&format!(r#"<definedName name="{n}">{f}</definedName>"#)),
                }
            }
            s.push_str("</definedNames>");
        }
        s.push_str("</workbook>");
        s.into_bytes()
    }

    fn rels_with(entries: &[(&str, &str, TargetMode)]) -> RelsGraph {
        let mut g = RelsGraph::new();
        for (rt_str, target, mode) in entries {
            g.add(rt_str, target, *mode);
        }
        g
    }

    fn table_part_bytes(name: &str, id: u32) -> Vec<u8> {
        format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="{id}" name="{name}" displayName="{name}" ref="A1:B2"><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>"#).into_bytes()
    }

    fn one_sheet_workbook(extra_titles: &[&str], defined_names: &[(&str, Option<u32>, &str)]) -> Vec<u8> {
        let mut titles = vec!["Template"];
        titles.extend_from_slice(extra_titles);
        workbook_xml(&titles, defined_names)
    }

    #[test]
    fn basic_clone_no_rels() {
        let mut alloc = PartIdAllocator::from_zip_parts(["xl/worksheets/sheet1.xml"].iter().copied());
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Template Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan ok");

        assert_eq!(mutations.new_sheet_path, "xl/worksheets/sheet2.xml");
        // No ancillary parts (only the dest rels file, but it's empty here so
        // no rels file is emitted).
        assert!(mutations.new_ancillary_parts.is_empty());
        assert!(mutations.defined_names_to_add.is_empty());
        assert!(mutations.new_table_names.is_empty());
        // CT override for the worksheet itself.
        assert_eq!(
            mutations.content_type_overrides_to_add[0].0,
            "/xl/worksheets/sheet2.xml"
        );
        assert_eq!(mutations.content_type_overrides_to_add[0].1, CT_WORKSHEET);
        // Workbook rels has the new worksheet entry.
        assert_eq!(mutations.workbook_rels_to_add.len(), 1);
        assert_eq!(mutations.workbook_rels_to_add[0].1, rt::WORKSHEET);
        // The placeholder rId appears verbatim in workbook_sheets_append.
        let entry = std::str::from_utf8(&mutations.workbook_sheets_append).unwrap();
        assert!(entry.contains(r#"name="Template Copy""#), "{entry}");
        assert!(entry.contains("__SHEET_RID_PLACEHOLDER_"), "{entry}");
    }

    #[test]
    fn clone_with_one_table_dedups_name() {
        let mut alloc = PartIdAllocator::from_zip_parts(
            [
                "xl/worksheets/sheet1.xml",
                "xl/tables/table1.xml",
            ]
            .iter()
            .copied(),
        );
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), sheet_xml_with_table_part("rId1"));
        zip_parts.insert("xl/tables/table1.xml".into(), table_part_bytes("Sales", 1));
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let mut existing_table_names: HashSet<String> = HashSet::new();
        existing_table_names.insert("Sales".into());
        let source_rels =
            rels_with(&[(rt::TABLE, "../tables/table1.xml", TargetMode::Internal)]);

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Template Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan ok");

        // One new table part allocated (table2.xml since seed = max+1 = 2).
        let table_part = mutations
            .new_ancillary_parts
            .iter()
            .find(|(p, _)| p == "xl/tables/table2.xml")
            .expect("table2.xml emitted");
        let body = std::str::from_utf8(&table_part.1).unwrap();
        assert!(body.contains(r#"name="Sales_2""#), "{body}");
        assert!(body.contains(r#"displayName="Sales_2""#), "{body}");
        assert_eq!(mutations.new_table_names, vec!["Sales_2".to_string()]);

        // The cloned sheet XML's r:id should now reference the new rId.
        let new_sheet = std::str::from_utf8(&mutations.new_sheet_xml).unwrap();
        let new_rid = mutations.rid_remap.get("rId1").expect("rId1 remapped");
        assert!(new_sheet.contains(&format!(r#"r:id="{new_rid}""#)), "{new_sheet}");
    }

    #[test]
    fn clone_with_n_tables_dedups_each() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert(
            "xl/worksheets/sheet1.xml".into(),
            br#"<?xml version="1.0"?><worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts></worksheet>"#.to_vec(),
        );
        zip_parts.insert("xl/tables/table1.xml".into(), table_part_bytes("Sales", 1));
        zip_parts.insert("xl/tables/table2.xml".into(), table_part_bytes("Costs", 2));
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let mut existing_table_names: HashSet<String> = HashSet::new();
        existing_table_names.insert("Sales".into());
        existing_table_names.insert("Costs".into());
        let source_rels = rels_with(&[
            (rt::TABLE, "../tables/table1.xml", TargetMode::Internal),
            (rt::TABLE, "../tables/table2.xml", TargetMode::Internal),
        ]);

        // Pre-seed the allocator past existing files.
        for p in zip_parts.keys() {
            alloc.observe(p);
        }

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan ok");

        assert_eq!(mutations.new_table_names, vec!["Sales_2".to_string(), "Costs_2".to_string()]);

        // Both tables emitted with fresh suffixes.
        let table_paths: Vec<&str> = mutations
            .new_ancillary_parts
            .iter()
            .map(|(p, _)| p.as_str())
            .filter(|p| p.starts_with("xl/tables/"))
            .collect();
        assert!(table_paths.contains(&"xl/tables/table3.xml"), "{table_paths:?}");
        assert!(table_paths.contains(&"xl/tables/table4.xml"), "{table_paths:?}");
    }

    #[test]
    fn clone_with_external_hyperlink_alias_url() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = rels_with(&[
            (rt::HYPERLINK, "https://example.com/x", TargetMode::External),
            (rt::HYPERLINK, "mailto:a@b.com", TargetMode::External),
        ]);

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "T2".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan ok");

        // Both hyperlinks remap to fresh rIds in the new dest rels graph.
        assert_eq!(mutations.rid_remap.len(), 2);
        // The destination rels file is emitted (non-empty).
        let rels_part = mutations
            .new_ancillary_parts
            .iter()
            .find(|(p, _)| p.ends_with(".rels"))
            .expect("rels file emitted");
        let rels_body = std::str::from_utf8(&rels_part.1).unwrap();
        assert!(rels_body.contains("https://example.com/x"), "{rels_body}");
        assert!(rels_body.contains("mailto:a@b.com"), "{rels_body}");
    }

    #[test]
    fn clone_with_comments_and_vml() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert(
            "xl/worksheets/sheet1.xml".into(),
            br#"<?xml version="1.0"?><worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><legacyDrawing r:id="rId2"/></worksheet>"#.to_vec(),
        );
        zip_parts.insert("xl/comments1.xml".into(), b"<comments/>".to_vec());
        zip_parts.insert("xl/drawings/vmlDrawing1.vml".into(), b"<xml/>".to_vec());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = rels_with(&[
            (rt::COMMENTS, "../comments1.xml", TargetMode::Internal),
            (rt::VML_DRAWING, "../drawings/vmlDrawing1.vml", TargetMode::Internal),
        ]);
        for p in zip_parts.keys() {
            alloc.observe(p);
        }

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "T2".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan ok");

        let paths: Vec<&str> = mutations
            .new_ancillary_parts
            .iter()
            .map(|(p, _)| p.as_str())
            .collect();
        assert!(paths.contains(&"xl/comments2.xml"), "{paths:?}");
        assert!(paths.contains(&"xl/drawings/vmlDrawing2.vml"), "{paths:?}");
    }

    #[test]
    fn clone_with_drawing_aliases_image() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        zip_parts.insert("xl/drawings/drawing1.xml".into(), b"<wsDr/>".to_vec());
        // Drawing's own rels file → image1.png.
        let drawing_rels_xml = rels_with(&[(rt::IMAGE, "../media/image1.png", TargetMode::Internal)]).serialize();
        zip_parts.insert("xl/drawings/_rels/drawing1.xml.rels".into(), drawing_rels_xml);
        zip_parts.insert("xl/media/image1.png".into(), b"\x89PNG".to_vec());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = rels_with(&[(
            rt::DRAWING,
            "../drawings/drawing1.xml",
            TargetMode::Internal,
        )]);
        for p in zip_parts.keys() {
            alloc.observe(p);
        }

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "T2".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan ok");

        // New drawing emitted at drawing2.xml.
        let drawing_part = mutations
            .new_ancillary_parts
            .iter()
            .find(|(p, _)| p == "xl/drawings/drawing2.xml")
            .expect("drawing emitted");
        assert_eq!(drawing_part.1.as_slice(), b"<wsDr/>");

        // Cloned drawing rels file emitted, with image still pointing at
        // the ORIGINAL image1.png (alias, not deep clone).
        let drawing_rels_part = mutations
            .new_ancillary_parts
            .iter()
            .find(|(p, _)| p == "xl/drawings/_rels/drawing2.xml.rels")
            .expect("drawing rels emitted");
        let body = std::str::from_utf8(&drawing_rels_part.1).unwrap();
        assert!(body.contains("../media/image1.png"), "{body}");

        // The image part is NOT in new_ancillary_parts (aliasing).
        assert!(!mutations
            .new_ancillary_parts
            .iter()
            .any(|(p, _)| p == "xl/media/image1.png"));
    }

    #[test]
    fn clone_with_sheet_scoped_defined_name() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        // src_idx = 0 (Template). new_idx = 2 (after Sheet2 inserts).
        let workbook_bytes = workbook_xml(
            &["Template", "Other"],
            &[
                ("_xlnm.Print_Area", Some(0), "Template!$A$1:$E$10"),
                ("WorkbookScopeName", None, "Other!$A$1"),
                ("OtherSheetName", Some(1), "Other!$B$2"),
            ],
        );
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Template Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan ok");

        // Only the Template-scoped name was cloned.
        assert_eq!(mutations.defined_names_to_add.len(), 1);
        let dn = &mutations.defined_names_to_add[0];
        assert_eq!(dn.name, "_xlnm.Print_Area");
        // new_idx = sheet_titles.len() = 2.
        assert_eq!(dn.local_sheet_id, 2);
        assert_eq!(dn.formula, "Template!$A$1:$E$10");
        // Element bytes carry the new localSheetId.
        let body = std::str::from_utf8(&dn.element_bytes).unwrap();
        assert!(body.contains(r#"localSheetId="2""#), "{body}");
        assert!(body.contains("_xlnm.Print_Area"), "{body}");
    }

    #[test]
    fn clone_without_sheet_scoped_names_yields_empty_list() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = workbook_xml(
            &["Template", "Other"],
            &[("WorkbookScopeName", None, "A1")],
        );
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Template Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan ok");

        assert!(mutations.defined_names_to_add.is_empty());
    }

    #[test]
    fn validation_duplicate_dst_title_errors() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = workbook_xml(&["Template", "Existing"], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let err = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Existing".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect_err("duplicate dst should error");
        match err {
            SheetCopyError::DuplicateDestinationTitle(s) => assert_eq!(s, "Existing"),
            other => panic!("wrong error: {other:?}"),
        }
    }

    #[test]
    fn validation_missing_source_title_errors() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = workbook_xml(&["Template"], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let err = plan_sheet_copy(SheetCopyInputs {
            src_title: "Nonexistent".into(),
            dst_title: "Foo".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect_err("missing src should error");
        assert!(matches!(err, SheetCopyError::MissingSourceTitle(_)));
    }

    #[test]
    fn validation_missing_source_bytes_errors() {
        let mut alloc = PartIdAllocator::new();
        let zip_parts: HashMap<String, Vec<u8>> = HashMap::new(); // empty
        let workbook_bytes = workbook_xml(&["Template"], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let err = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect_err("missing src bytes should error");
        assert!(matches!(err, SheetCopyError::MissingSourceSheetBytes(_)));
    }

    #[test]
    fn rid_remap_applied_to_sheet_xml_in_one_pass() {
        // Source sheet has TWO different r:id references; both must be
        // rewritten in a single pass to point at fresh dest rIds.
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        let sheet_src = br#"<?xml version="1.0"?><worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><tableParts count="1"><tablePart r:id="rId1"/></tableParts><legacyDrawing r:id="rId2"/></worksheet>"#.to_vec();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), sheet_src);
        zip_parts.insert("xl/tables/table1.xml".into(), table_part_bytes("Sales", 1));
        zip_parts.insert("xl/drawings/vmlDrawing1.vml".into(), b"<xml/>".to_vec());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = rels_with(&[
            (rt::TABLE, "../tables/table1.xml", TargetMode::Internal),
            (rt::VML_DRAWING, "../drawings/vmlDrawing1.vml", TargetMode::Internal),
        ]);
        for p in zip_parts.keys() {
            alloc.observe(p);
        }

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan ok");

        // Both source rIds must be mapped.
        let new_for_table = mutations.rid_remap.get("rId1").expect("rId1 mapped");
        let new_for_vml = mutations.rid_remap.get("rId2").expect("rId2 mapped");
        assert_ne!(new_for_table, new_for_vml);

        let new_xml = std::str::from_utf8(&mutations.new_sheet_xml).unwrap();
        assert!(new_xml.contains(&format!(r#"r:id="{new_for_table}""#)), "{new_xml}");
        assert!(new_xml.contains(&format!(r#"r:id="{new_for_vml}""#)), "{new_xml}");
        // Old rIds should be gone.
        assert!(!new_xml.contains(r#"r:id="rId1""#) || new_for_table == "rId1");
    }

    #[test]
    fn dedup_table_name_uses_underscore_with_n_starting_at_two() {
        let mut taken = HashSet::new();
        assert_eq!(dedup_table_name("Sales", &taken), "Sales_2");
        taken.insert("Sales".into());
        assert_eq!(dedup_table_name("Sales", &taken), "Sales_2");
        taken.insert("Sales_2".into());
        assert_eq!(dedup_table_name("Sales", &taken), "Sales_3");
        taken.insert("Sales_3".into());
        assert_eq!(dedup_table_name("Sales", &taken), "Sales_4");
        // A base ending in digits still gets _N — no ambiguity.
        let taken2 = HashSet::new();
        assert_eq!(dedup_table_name("Sales2024", &taken2), "Sales2024_2");
    }

    #[test]
    fn placeholder_rid_reference_consistent_between_sheet_entry_and_workbook_rels() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan ok");

        // The placeholder in workbook_sheets_append matches the placeholder
        // in workbook_rels_to_add[0].0 — caller does one rename.
        let placeholder = &mutations.workbook_rels_to_add[0].0;
        let entry = std::str::from_utf8(&mutations.workbook_sheets_append).unwrap();
        assert!(entry.contains(placeholder), "entry={entry} placeholder={placeholder}");
    }

    #[test]
    fn allocator_advances_across_sequential_calls() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = one_sheet_workbook(&["Other"], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let m1 = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "C1".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan 1 ok");
        // Second call: the workbook_xml argument is the same (caller hasn't
        // applied m1 yet), but the allocator has already advanced.
        let m2 = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "C2".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan 2 ok");

        assert_ne!(m1.new_sheet_path, m2.new_sheet_path);
        let n1 = m1
            .new_sheet_path
            .strip_prefix("xl/worksheets/sheet")
            .and_then(|s| s.strip_suffix(".xml"))
            .and_then(|s| s.parse::<u32>().ok())
            .unwrap();
        let n2 = m2
            .new_sheet_path
            .strip_prefix("xl/worksheets/sheet")
            .and_then(|s| s.strip_suffix(".xml"))
            .and_then(|s| s.parse::<u32>().ok())
            .unwrap();
        assert!(n2 > n1, "{n1} < {n2}");
    }

    #[test]
    fn workbook_with_no_defined_names_block_does_not_error() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = workbook_xml(&["Template"], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan ok");
        assert!(mutations.defined_names_to_add.is_empty());
    }

    #[test]
    fn unknown_rel_type_is_aliased_not_dropped() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = rels_with(&[
            ("http://example.com/rels/weird", "../weird/path.bin", TargetMode::Internal),
        ]);

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
        deep_copy_images: false,
        })
        .expect("plan ok");
        // The rId is in the remap (aliased into dest rels graph).
        assert_eq!(mutations.rid_remap.len(), 1);
    }
}
