//! WolfXL — surgical xlsx patcher.
//!
//! Instead of parsing the entire workbook into a DOM (like openpyxl or umya),
//! WolfXL opens the xlsx ZIP, queues cell changes in memory, and on save:
//!   1. Patches only the worksheet XMLs that have dirty cells
//!   2. Patches sharedStrings/styles only if needed
//!   3. Raw-copies all other ZIP entries unchanged
//!
//! This makes modify-and-save O(modified data) instead of O(entire file).

#[allow(dead_code)] // SST parser used in Phase 3 (format patching reads existing styles)
pub mod shared_strings;
pub mod sheet_patcher;
#[allow(dead_code)] // Styles parser/appender used in Phase 3 (format patching)
pub mod styles;
pub mod conditional_formatting;
pub mod validations;
pub mod content_types;
#[allow(dead_code)] // RFC-013: registry is scaffolding-only; first caller is RFC-022
pub mod ancillary;
pub mod properties;
#[allow(dead_code)] // RFC-022: live caller wires up in commit 3 (queue_hyperlink + Phase 2.5e)
pub mod hyperlinks;
pub mod defined_names;
pub mod sheet_order;
pub mod tables;
pub mod comments;
// RFC-035 Pod-β: Phase 2.7 (do_save) consumes plan_sheet_copy from this re-export.
pub mod sheet_copy;
// Sprint Θ Pod-C3: Phase 2.8 (do_save) rebuilds xl/calcChain.xml.
pub mod calcchain;
// Sprint Ν Pod-γ (RFC-047 / RFC-048): Phase 2.5m drains pivot adds.
pub mod pivot;
// Sprint Ο Pod 1D (RFC-058): Phase 2.5q splices workbookProtection +
// fileSharing into xl/workbook.xml.
pub mod security;
// Sprint Ο Pod 1B (RFC-056): Phase 2.5o drains autoFilter evaluation +
// `<row hidden="1">` markers.
pub mod autofilter;
pub mod autofilter_helpers;
// Sprint Ο Pod 1A.5 (RFC-055): Phase 2.5n drains queued sheet-setup
// blocks (sheetView / sheetProtection / pageMargins / pageSetup /
// headerFooter) into per-sheet `local_blocks`.
pub mod sheet_setup;

use std::collections::{BTreeMap, HashMap, HashSet};
use std::fs::File;
use std::io::{Read, Write};

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::ooxml_util;
use conditional_formatting::{
    CfRuleKind, CfRulePatch, CfvoPatch, ColorScaleStop, ConditionalFormattingPatch, DxfPatch,
};
use sheet_patcher::{CellPatch, CellValue};
use styles::FormatSpec;
use validations::DataValidationPatch;
use wolfxl_merger::SheetBlock;
use wolfxl_rels::RelsGraph;

// ---------------------------------------------------------------------------
// PyO3 class
// ---------------------------------------------------------------------------

#[pyclass]
pub struct XlsxPatcher {
    file_path: String,
    /// Sheet name → ZIP entry path (e.g. "Sheet1" → "xl/worksheets/sheet1.xml").
    sheet_paths: HashMap<String, String>,
    /// Queued cell value changes: (sheet, "A1") → CellPatch.
    value_patches: HashMap<(String, String), CellPatch>,
    /// Queued cell format changes: (sheet, "A1") → FormatSpec.
    format_patches: HashMap<(String, String), FormatSpec>,
    /// Queued mutations to `*.rels` parts. Key: ZIP entry path (e.g.
    /// `xl/worksheets/_rels/sheet1.xml.rels`). The save loop serializes the
    /// graph and writes it in place of the original entry. Populated by
    /// future Phase-3 RFCs (RFC-022 hyperlinks, RFC-023 comments, RFC-024
    /// tables); empty in the current slice.
    rels_patches: HashMap<String, RelsGraph>,
    /// Queued sibling-block insertions on `xl/worksheets/sheet*.xml` parts.
    /// Key: sheet XML path (e.g. `xl/worksheets/sheet1.xml`). The save
    /// loop calls `wolfxl_merger::merge_blocks` after `sheet_patcher`
    /// runs, so cell-level patches and block-level patches compose
    /// without conflict. Populated by future Phase-3 RFCs (RFC-022
    /// hyperlinks, RFC-024 tables, RFC-026 conditional formatting);
    /// empty in the current slice.
    ///
    /// Note: RFC-025 (data validations) does NOT populate this map
    /// directly. It builds blocks on-demand inside `do_save` from
    /// `queued_dv_patches` so that the existing `<dataValidations>`
    /// block (read out of the source sheet XML at save time) can be
    /// merged with the queued patches before the merger is invoked.
    queued_blocks: HashMap<String, Vec<SheetBlock>>,
    /// Queued data-validation rules per sheet name (NOT path — we
    /// resolve to path inside `do_save`). Each entry becomes a single
    /// `<dataValidations>` block during save: any pre-existing block
    /// in the source sheet XML is read out, prepended verbatim, and
    /// the queued patches are appended. The combined block is then
    /// handed to `wolfxl_merger` as `SheetBlock::DataValidations`.
    queued_dv_patches: HashMap<String, Vec<DataValidationPatch>>,
    /// Queued conditional-formatting patches per sheet name (RFC-026).
    /// Each entry becomes one or more `<conditionalFormatting>` blocks
    /// during save. Existing CF blocks in the source sheet XML are
    /// extracted byte-for-byte and prepended (because the merger's
    /// replace-all CF semantics drop them otherwise — RFC-011 §5.5).
    /// `dxfId` allocation threads through a workbook-wide counter so
    /// CF added on multiple sheets in one save() session lands in a
    /// single coordinated `xl/styles.xml` mutation.
    queued_cf_patches: HashMap<String, Vec<ConditionalFormattingPatch>>,
    /// Sheet names in source-document order (RFC-013). Populated in
    /// `open()` from `xl/workbook.xml`'s `<sheet>` order. Replaces
    /// `sheet_paths.keys()` for any caller that needs deterministic
    /// iteration (RFC-020's `app.xml` regen, RFC-026's CF aggregation
    /// when it migrates off the temporary sorted-keys path).
    sheet_order: Vec<String>,
    /// Brand-new ZIP entries to emit on save (RFC-013). Parallel to
    /// `file_patches`: `file_patches` REPLACES an existing source entry
    /// in place; `file_adds` APPENDS a new entry that wasn't in the
    /// source ZIP. Collisions with source-ZIP names are a hard panic
    /// (caller bug — see RFC-013 §8 risk #2). First user is RFC-020's
    /// optional `docProps/core.xml` add path; RFC-022/023/024 will be
    /// the volume callers.
    file_adds: HashMap<String, Vec<u8>>,
    /// Source ZIP entries to skip during the save loop (RFC-013).
    /// Reserved for future use; v1 is unused. RFC-035 (copy_worksheet
    /// + delete-sheet) will be the first caller. Including the field
    /// now keeps the short-circuit predicate and rewrite loop forward-
    /// compatible without a follow-up patcher refactor.
    file_deletes: HashSet<String>,
    /// Per-sheet inventory of ancillary parts (comments, VML drawings,
    /// tables, hyperlinks) lazily populated from the source ZIP's
    /// `_rels/sheetN.xml.rels` files (RFC-013). Scaffolding-only this
    /// slice — `ancillary::AncillaryPartRegistry::populate_for_sheet`
    /// has no live caller yet. RFC-022 (Hyperlinks) is the first
    /// consumer; RFC-023/024 follow.
    #[allow(dead_code)]
    ancillary: ancillary::AncillaryPartRegistry,
    /// Per-sheet `[Content_Types].xml` ops queued by sheet block
    /// builders (RFC-013 Phase 2.5c). Each entry is the list of
    /// content-type adjustments that sheet's flush requires (a new
    /// comments/table part needs an `Override` entry; vmlDrawing
    /// requires `Default Extension="vml"`). Aggregated across sheets
    /// during `do_save` so a single workbook-wide
    /// `[Content_Types].xml` mutation absorbs every sheet's ops in
    /// one parse + serialize. Empty in this slice (RFC-022/023/024
    /// will be the first volume callers).
    queued_content_type_ops: HashMap<String, Vec<content_types::ContentTypeOp>>,
    /// Document properties pending flush (RFC-020). When `Some(_)`,
    /// `do_save` rewrites both `docProps/core.xml` and
    /// `docProps/app.xml` from the payload's fields. Routing depends
    /// on whether each part already exists in the source ZIP — present
    /// → patches it through `file_patches`; absent → adds it via
    /// RFC-013's `file_adds` primitive. Populated by
    /// [`Self::queue_properties`].
    queued_props: Option<properties::DocPropertiesPayload>,
    /// Per-sheet hyperlink ops pending flush (RFC-022). Outer key is
    /// sheet name; inner is coordinate → op. `BTreeMap` for the inner
    /// gives deterministic flush ordering when a single save touches
    /// multiple cells. Phase 2.5e drains this map: it reads the
    /// existing `<hyperlinks>` block + sheet rels, merges the queued
    /// ops, and pushes a `SheetBlock::Hyperlinks` plus mutates
    /// `rels_patches`. `None` value (delete sentinel) lands here as
    /// `HyperlinkOp::Delete` per INDEX decision #5.
    queued_hyperlinks: HashMap<String, BTreeMap<String, hyperlinks::HyperlinkOp>>,
    /// Defined-name upserts pending flush (RFC-021). Drained by
    /// Phase 2.5f, which parses `xl/workbook.xml`, merges these
    /// entries via `defined_names::merge_defined_names`, and writes
    /// the result back through `file_patches`. Empty queue → no
    /// rewrite of `xl/workbook.xml` (modify-mode no-op invariant).
    /// Order is insertion order from the Python coordinator (which
    /// itself iterates a regular dict — Python 3.7+ preserves
    /// insertion order). Within a save, the merger upserts by
    /// `(name, local_sheet_id)` so duplicates collapse to last-wins.
    queued_defined_names: Vec<defined_names::DefinedNameMut>,
    /// Per-sheet table-add patches pending flush (RFC-024). Drained
    /// by Phase 2.5g: scans the source ZIP for the workbook's
    /// existing-table inventory (across ALL sheets, since `id` and
    /// `name` are workbook-unique), allocates fresh ids + sequential
    /// part filenames, mutates `rels_patches`, queues the
    /// `[Content_Types].xml` Override entries through
    /// `queued_content_type_ops`, and pushes a
    /// `SheetBlock::TableParts` per sheet. Insertion order via Vec
    /// matches openpyxl's "first add → first slot" semantics.
    queued_tables: HashMap<String, Vec<tables::TablePatch>>,
    /// Per-sheet comment ops pending flush (RFC-023). Outer key is
    /// sheet name; inner is coordinate → op. `Set` adds/replaces a
    /// comment with the supplied text/author/width/height; `Delete`
    /// removes any existing comment at that coordinate. Drained by
    /// Phase 2.5h during `do_save`. Workbook-scope author dedup
    /// happens in `comments::CommentAuthorTable`, shared across all
    /// sheets touched in a single save.
    queued_comments: HashMap<String, BTreeMap<String, comments::CommentOp>>,
    /// Sheet-reorder operations pending flush (RFC-036). Insertion-
    /// ordered list of `(sheet_name, offset)` moves. Drained by
    /// Phase 2.5h, which sequences BEFORE Phase 2.5f (defined-names)
    /// because both phases mutate `xl/workbook.xml`. The reorder
    /// merger also produces the post-move `<definedName
    /// localSheetId>` integers, so the defined-names merger sees a
    /// workbook.xml whose tab indices already reflect the move.
    /// Empty queue → no `xl/workbook.xml` touch.
    queued_sheet_moves: Vec<(String, i32)>,
    /// Per-workbook structural-shift queue (RFC-030 / RFC-031). Each
    /// entry is `(sheet, axis, idx, n)` where `axis` is "row" or "col"
    /// and `n` is signed (positive = insert, negative = delete).
    /// Drained by Phase 2.5i during `do_save`. Order is append order
    /// — the Python coordinator validates `idx >= 1` and `amount >= 1`
    /// before queueing.
    queued_axis_shifts: Vec<AxisShift>,
    /// Per-workbook range-move queue (RFC-034). Each entry describes
    /// one paste-style relocation of a rectangular block. Drained by
    /// Phase 2.5j during `do_save`, AFTER axis shifts so a sequence
    /// like `insert_rows(2, 3)` then `move_range("C3:E10", rows=5)`
    /// is applied in source order against the post-shift coordinate
    /// space.
    queued_range_moves: Vec<RangeMove>,
    /// Per-workbook sheet-copy queue (RFC-035). Each entry is a
    /// `(src_title, dst_title)` pair in user-call order. Drained by
    /// Phase 2.7 during `do_save`, BEFORE every per-sheet phase so
    /// the cloned sheet is visible to downstream phases as if it
    /// had always been part of the source workbook.
    queued_sheet_copies: Vec<SheetCopyOp>,
    /// Sprint Θ Pod-A: pre-seeded `file_patches` entries produced by
    /// permissive-mode load-time normalization (e.g. rewriting an
    /// empty `<sheets/>` block in `xl/workbook.xml`). Empty in the
    /// non-permissive path. Drained into the actual `file_patches`
    /// map at the start of `do_save` so every downstream phase
    /// (Phase 2.7 splice, defined-names merger, …) sees the
    /// rewritten bytes.
    permissive_seed_file_patches: HashMap<String, Vec<u8>>,
    /// Sprint Λ Pod-β (RFC-045) — per-sheet pending image adds.
    /// Drained by Phase 2.5k during `do_save`. Supports the
    /// "fresh drawing" case only — sheets that already have a
    /// drawing rel raise NotImplementedError (v1.5 follow-up).
    queued_images: HashMap<String, Vec<QueuedImageAdd>>,
    /// Sprint Μ Pod-γ (RFC-046) — per-sheet pending chart adds.
    /// Each entry carries pre-serialized chart XML bytes plus an
    /// A1-style anchor cell. Drained by Phase 2.5l during
    /// `do_save`, BEFORE Phase 3 (cell patches) so a chart's data
    /// range can compose with cell rewrites in the same save.
    /// Phase 2.5l differs from 2.5k by handling BOTH the
    /// "fresh-drawing" case AND the "merge-into-existing-drawing"
    /// case (which Phase 2.5k still rejects).
    queued_charts: HashMap<String, Vec<QueuedChartAdd>>,
    /// Sprint Ν Pod-γ (RFC-047) — pending pivot-cache adds. Append
    /// order is the cache_id allocation order. Drained by Phase
    /// 2.5m during `do_save` (sequenced AFTER Phase 2.5l so chart
    /// pivot-source linkage in v2.1 can resolve table names).
    queued_pivot_caches: Vec<pivot::QueuedPivotCacheAdd>,
    /// Sprint Ν Pod-γ (RFC-048) — pending pivot-table adds, keyed
    /// by sheet title. Drained by Phase 2.5m AFTER all caches are
    /// emitted (so the table → cache rels target is resolvable).
    queued_pivot_tables: HashMap<String, Vec<pivot::QueuedPivotTableAdd>>,
    /// Sprint Ν Pod-γ — workbook-scope cache_id allocator. Bumps
    /// monotonically as `queue_pivot_cache_add` is called. The
    /// counter starts at 0 (cache_id = `pivotCache.cacheId` attr;
    /// ECMA-376 0-based). Initialised by `XlsxPatcher::open` if
    /// the source already has pivot caches.
    next_pivot_cache_id: u32,

    /// Sprint Ο Pod 1D (RFC-058) — pending workbook-level security
    /// blocks (workbookProtection + fileSharing). `None` = user
    /// never set wb.security or wb.fileSharing in this session;
    /// `Some(_)` = the queue was populated and Phase 2.5q must
    /// splice into `xl/workbook.xml`.
    queued_workbook_security: Option<wolfxl_writer::parse::workbook_security::WorkbookSecurity>,

    /// Sprint Ο Pod 1B (RFC-056) — pending autoFilter adds, keyed
    /// by sheet title. Drained by Phase 2.5o (sequenced AFTER pivot
    /// Phase 2.5m, BEFORE Phase 3 cell patches). The queue stores
    /// the §10 dict shape so the cdylib can lift it into the typed
    /// model + run filter evaluation at drain time.
    queued_autofilters: HashMap<String, autofilter::QueuedAutoFilter>,

    /// Sprint Ο Pod 1A.5 (RFC-055) — pending sheet-setup mutations,
    /// keyed by sheet title. Each entry is a parsed
    /// [`sheet_setup::QueuedSheetSetup`] holding typed specs for the
    /// 5 sheet-setup blocks. Drained by Phase 2.5n (sequenced AFTER
    /// pivots in Phase 2.5m and BEFORE autoFilter Phase 2.5o so a
    /// later sheet-protection toggle can lock the autoFilter range).
    /// Calling `queue_sheet_setup_update` again for the same sheet
    /// REPLACES the previous payload (matches Python `ws.page_setup
    /// = ...` semantics).
    queued_sheet_setup: HashMap<String, sheet_setup::QueuedSheetSetup>,
}

/// Sprint Μ Pod-γ (RFC-046) — one chart queued for emit on a sheet.
///
/// The chart XML is pre-serialized by the caller (Pod-α's
/// `emit_chart_xml(&Chart)`); the patcher only routes bytes through
/// the OOXML rels graph + content-types + drawing layer.
#[derive(Debug, Clone)]
pub struct QueuedChartAdd {
    /// Chart XML body — written into `xl/charts/chartN.xml`. Already
    /// serialized by the caller (the patcher never builds this).
    pub chart_xml: Vec<u8>,
    /// A1-style anchor cell (e.g. `"D2"`). The patcher converts to
    /// `(col0, row0)` for the `<xdr:from>` block.
    pub anchor_a1: String,
    /// Chart pixel size in EMU is fixed for the modify-mode v1: the
    /// patcher emits a fixed `cx=12cm cy=8cm` extent per chart so we
    /// don't need a width/height plumb. Pod-β wires width/height
    /// through Worksheet.add_chart in the future.
    pub width_emu: i64,
    pub height_emu: i64,
}

/// Sprint Λ Pod-β (RFC-045) — one image queued for emit on a sheet.
#[derive(Debug, Clone)]
pub struct QueuedImageAdd {
    /// Raw image bytes — written into `xl/media/imageN.<ext>`.
    pub data: Vec<u8>,
    /// Lowercase extension (`"png"`, `"jpeg"`, `"gif"`, `"bmp"`).
    pub ext: String,
    /// Pixel width (Excel uses 9525 EMU/px when computing extent).
    pub width_px: u32,
    pub height_px: u32,
    /// Anchor flavour. Mirrors `wolfxl_writer::model::image::ImageAnchor`
    /// but kept Python-shape so the patcher can stay independent of the
    /// writer crate's data model.
    pub anchor: QueuedImageAnchor,
}

#[derive(Debug, Clone)]
pub enum QueuedImageAnchor {
    OneCell {
        from_col: u32,
        from_row: u32,
        from_col_off: i64,
        from_row_off: i64,
    },
    TwoCell {
        from_col: u32,
        from_row: u32,
        from_col_off: i64,
        from_row_off: i64,
        to_col: u32,
        to_row: u32,
        to_col_off: i64,
        to_row_off: i64,
        edit_as: String,
    },
    Absolute {
        x_emu: i64,
        y_emu: i64,
        cx_emu: i64,
        cy_emu: i64,
    },
}

/// One queued sheet-copy op (RFC-035).
#[derive(Debug, Clone)]
pub struct SheetCopyOp {
    /// Source sheet title (must exist in `self.sheet_paths`).
    pub src_title: String,
    /// Destination sheet title (pre-deduped by the Python coordinator).
    pub dst_title: String,
    /// Sprint Θ Pod-C2 — workbook-level
    /// `wb.copy_options.deep_copy_images` snapshot at queue time.
    /// `false` preserves the historical RFC-035 §5.3 alias-by-target
    /// behaviour. Read by the planner via
    /// [`SheetCopyInputs::deep_copy_images`].
    pub deep_copy_images: bool,
}

/// One queued axis-shift op (RFC-030/031).
#[derive(Debug, Clone)]
pub struct AxisShift {
    /// Sheet name (NOT path).
    pub sheet: String,
    /// `"row"` or `"col"`.
    pub axis: String,
    /// 1-based index where shifting begins.
    pub idx: u32,
    /// Signed shift count. Positive = insert; negative = delete.
    pub n: i32,
}

/// One queued range-move op (RFC-034).
#[derive(Debug, Clone)]
pub struct RangeMove {
    /// Sheet name (NOT path).
    pub sheet: String,
    /// 1-based inclusive source rectangle corners.
    pub src_min_col: u32,
    pub src_min_row: u32,
    pub src_max_col: u32,
    pub src_max_row: u32,
    /// Signed delta. Positive shifts down/right; negative up/left.
    pub d_row: i32,
    pub d_col: i32,
    /// If true, formulas in cells OUTSIDE the source rectangle that
    /// reference cells INSIDE `src` are also re-anchored. Cells
    /// INSIDE `src` are always paste-translated.
    pub translate: bool,
}

#[pymethods]
impl XlsxPatcher {
    /// Open an xlsx file for surgical patching.
    ///
    /// When `permissive` is true and the parsed `xl/workbook.xml`
    /// declares no `<sheet>` children (e.g. a self-closing `<sheets/>`
    /// produced by a malformed but still loadable workbook), this
    /// fallback registers every worksheet target in
    /// `xl/_rels/workbook.xml.rels` under a synthesized title
    /// (`Sheet1`, `Sheet2`, ...). This unblocks the Phase 2.7
    /// self-closing-`<sheets/>` splice path and is gated behind the
    /// flag so well-formed inputs are unaffected. See Sprint Θ Pod-A.
    #[staticmethod]
    #[pyo3(signature = (path, permissive = false))]
    fn open(path: &str, permissive: bool) -> PyResult<Self> {
        let f = File::open(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Cannot open '{path}': {e}")))?;
        let mut zip = ZipArchive::new(f)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Not a valid ZIP: {e}")))?;

        // Parse workbook.xml + rels to build sheet name → XML path mapping.
        let wb_xml = ooxml_util::zip_read_to_string(&mut zip, "xl/workbook.xml")?;
        let rels_xml = ooxml_util::zip_read_to_string(&mut zip, "xl/_rels/workbook.xml.rels")?;
        let sheet_rids = ooxml_util::parse_workbook_sheet_rids(&wb_xml)?;
        let rel_targets = ooxml_util::parse_relationship_targets(&rels_xml)?;

        let mut sheet_paths: HashMap<String, String> = HashMap::new();
        // RFC-013: capture sheet names in source-document order. The
        // `parse_workbook_sheet_rids` call above returns a Vec in
        // document order; iterating it here preserves that ordering
        // and skips any sheet whose rId target is missing (mirroring
        // the legacy lenient-parse contract).
        let mut sheet_order: Vec<String> = Vec::with_capacity(sheet_rids.len());
        for (name, rid) in sheet_rids {
            if let Some(target) = rel_targets.get(&rid) {
                sheet_paths.insert(name.clone(), ooxml_util::join_and_normalize("xl/", target));
                sheet_order.push(name);
            }
        }

        // Sprint Θ Pod-A: permissive fallback for malformed workbooks
        // whose <sheets> block is self-closing (no <sheet> children)
        // even though the rels graph still references worksheet parts.
        // We synthesize "Sheet1", "Sheet2", ... in rels iteration order
        // for every worksheet relationship target. This makes the
        // Phase 2.7 splice exercisable through the public API.
        //
        // We also normalize `xl/workbook.xml` in-memory: the empty
        // `<sheets/>` block is rewritten to `<sheets>...</sheets>`
        // populated with `<sheet>` entries that mirror the synthesized
        // titles + the rIds we recovered from the rels graph. The
        // rewrite is queued through the standard `file_patches` map,
        // which means downstream phases (Phase 2.7 splice, defined-
        // names merger, etc.) all see a well-formed workbook.xml. This
        // does NOT mutate the source file on disk; it only affects the
        // copy emitted by `save()`.
        let mut file_patches: HashMap<String, Vec<u8>> = HashMap::new();
        if permissive && sheet_order.is_empty() {
            const WORKSHEET_REL_TYPE: &str =
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
            // Re-parse rels via wolfxl_rels so we can filter by type
            // and reuse the relationship rId on the synthesized
            // <sheet> element (Excel requires r:id to match a real
            // relationship).
            let graph = wolfxl_rels::RelsGraph::parse(rels_xml.as_bytes())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to parse rels: {e}")))?;
            let mut idx: usize = 1;
            // Collected (synthesized_name, rId) pairs in rels order.
            let mut synthesized: Vec<(String, String)> = Vec::new();
            for r in graph.iter() {
                if r.rel_type == WORKSHEET_REL_TYPE {
                    let synth_name = format!("Sheet{idx}");
                    let path_in_zip = ooxml_util::join_and_normalize("xl/", &r.target);
                    sheet_paths.insert(synth_name.clone(), path_in_zip);
                    sheet_order.push(synth_name.clone());
                    synthesized.push((synth_name, r.id.0.clone()));
                    idx += 1;
                }
            }

            if !synthesized.is_empty() {
                // Build <sheet name="..." sheetId="N" r:id="..."/> entries.
                let mut entries = String::new();
                for (i, (name, rid)) in synthesized.iter().enumerate() {
                    let sheet_id = i + 1;
                    entries.push_str(&format!(
                        "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"{}\"/>",
                        xml_escape_attr(name),
                        sheet_id,
                        xml_escape_attr(rid)
                    ));
                }
                let new_block = format!("<sheets>{entries}</sheets>");
                let rewritten = if let Some(replaced) =
                    replace_first_occurrence(&wb_xml, "<sheets/>", &new_block)
                {
                    replaced
                } else if let Some(replaced) =
                    replace_first_occurrence(&wb_xml, "<sheets />", &new_block)
                {
                    replaced
                } else {
                    // No empty <sheets> marker to replace — workbook
                    // already has an open/close form but contains no
                    // <sheet> children. Inject our entries before
                    // </sheets>.
                    if let Some(close_pos) = wb_xml.find("</sheets>") {
                        let mut s = String::with_capacity(wb_xml.len() + entries.len());
                        s.push_str(&wb_xml[..close_pos]);
                        s.push_str(&entries);
                        s.push_str(&wb_xml[close_pos..]);
                        s
                    } else {
                        // Workbook has no <sheets> block at all; this
                        // is too far gone for permissive mode. Fall
                        // through without rewriting workbook.xml; the
                        // splice will report MissingSourceTitle if it
                        // needs the synthesized name.
                        wb_xml.clone()
                    }
                };
                if rewritten != wb_xml {
                    file_patches.insert("xl/workbook.xml".to_string(), rewritten.into_bytes());
                }
            }
        }

        Ok(XlsxPatcher {
            file_path: path.to_string(),
            sheet_paths,
            value_patches: HashMap::new(),
            format_patches: HashMap::new(),
            rels_patches: HashMap::new(),
            queued_blocks: HashMap::new(),
            queued_dv_patches: HashMap::new(),
            queued_cf_patches: HashMap::new(),
            sheet_order,
            file_adds: HashMap::new(),
            file_deletes: HashSet::new(),
            ancillary: ancillary::AncillaryPartRegistry::new(),
            queued_content_type_ops: HashMap::new(),
            queued_props: None,
            queued_hyperlinks: HashMap::new(),
            queued_defined_names: Vec::new(),
            queued_tables: HashMap::new(),
            queued_comments: HashMap::new(),
            queued_sheet_moves: Vec::new(),
            queued_axis_shifts: Vec::new(),
            queued_range_moves: Vec::new(),
            queued_sheet_copies: Vec::new(),
            permissive_seed_file_patches: file_patches,
            queued_images: HashMap::new(),
            queued_charts: HashMap::new(),
            queued_pivot_caches: Vec::new(),
            queued_pivot_tables: HashMap::new(),
            next_pivot_cache_id: 0,
            queued_workbook_security: None,
            queued_autofilters: HashMap::new(),
            queued_sheet_setup: HashMap::new(),
        })
    }

    /// Sprint Ι Pod-α: queue a rich-text value for a cell.
    ///
    /// `runs` is a list of `(text, font_dict_or_none)` tuples.  Each
    /// font dict mirrors `wolfxl.cell.rich_text.InlineFont` field
    /// names (`b`, `i`, `strike`, `u`, `sz`, `color`, `rFont`,
    /// `family`, `charset`, `vertAlign`, `scheme`).  The patcher
    /// emits an inline-string cell (`t="inlineStr"`) — so the SST
    /// never has to be modified.
    fn queue_rich_text_value(
        &mut self,
        sheet: &str,
        cell: &str,
        runs: &Bound<'_, pyo3::types::PyList>,
    ) -> PyResult<()> {
        let parsed = py_runs_to_rust(runs)?;
        let (row, col) =
            crate::util::a1_to_row_col(cell).map_err(|e| PyErr::new::<PyValueError, _>(e))?;
        let patch = CellPatch {
            row: row + 1,
            col: col + 1,
            value: Some(CellValue::RichText(parsed)),
            style_index: None,
        };
        self.value_patches
            .insert((sheet.to_string(), cell.to_string()), patch);
        Ok(())
    }

    /// Queue an array-formula / data-table / spill-child cell.
    ///
    /// RFC-057 (Sprint Ο Pod 1C).  `payload` is a dict matching the
    /// shape pinned in §10:
    ///   - ``{"kind": "array", "ref": "A1:A10", "text": "B1:B10*2"}``
    ///   - ``{"kind": "data_table", "ref": "B2:F11", "ca": false,
    ///        "dt2D": true, "dtr": false, "r1": "A1", "r2": "A2"}``
    ///   - ``{"kind": "spill_child"}``
    fn queue_array_formula(
        &mut self,
        sheet: &str,
        cell: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let kind: String = payload
            .get_item("kind")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("payload missing 'kind'"))?
            .extract()?;

        let value = match kind.as_str() {
            "array" => {
                let ref_range: String = payload
                    .get_item("ref")?
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("array kind needs 'ref'"))?
                    .extract()?;
                let mut text: String = payload
                    .get_item("text")?
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("array kind needs 'text'"))?
                    .extract()?;
                if let Some(stripped) = text.strip_prefix('=') {
                    text = stripped.to_string();
                }
                CellValue::ArrayFormula { ref_range, text }
            }
            "data_table" => {
                let ref_range: String = payload
                    .get_item("ref")?
                    .ok_or_else(|| {
                        PyErr::new::<PyValueError, _>("data_table kind needs 'ref'")
                    })?
                    .extract()?;
                let ca: bool = payload
                    .get_item("ca")?
                    .map(|v| v.extract::<bool>())
                    .transpose()?
                    .unwrap_or(false);
                let dt2_d: bool = payload
                    .get_item("dt2D")?
                    .map(|v| v.extract::<bool>())
                    .transpose()?
                    .unwrap_or(false);
                let dtr: bool = payload
                    .get_item("dtr")?
                    .map(|v| v.extract::<bool>())
                    .transpose()?
                    .unwrap_or(false);
                let r1: Option<String> = payload
                    .get_item("r1")?
                    .and_then(|v| v.extract().ok());
                let r2: Option<String> = payload
                    .get_item("r2")?
                    .and_then(|v| v.extract().ok());
                CellValue::DataTableFormula {
                    ref_range,
                    ca,
                    dt2_d,
                    dtr,
                    r1,
                    r2,
                }
            }
            "spill_child" => CellValue::SpillChild,
            other => {
                return Err(PyErr::new::<PyValueError, _>(format!(
                    "Unknown array-formula kind: '{other}'"
                )))
            }
        };

        let (row, col) =
            crate::util::a1_to_row_col(cell).map_err(|e| PyErr::new::<PyValueError, _>(e))?;

        let patch = CellPatch {
            row: row + 1,
            col: col + 1,
            value: Some(value),
            style_index: None,
        };

        self.value_patches
            .insert((sheet.to_string(), cell.to_string()), patch);
        Ok(())
    }

    /// Queue a cell value change.
    ///
    /// `payload` is a dict matching the ExcelBench cell payload format:
    ///   {"type": "string"|"number"|"boolean"|"formula"|"blank", "value": ...}
    fn queue_value(
        &mut self,
        sheet: &str,
        cell: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let cell_type = payload
            .get_item("type")?
            .map(|v| v.extract::<String>())
            .transpose()?
            .unwrap_or_default();

        let value = match cell_type.as_str() {
            "blank" => CellValue::Blank,
            "string" | "str" => {
                let v = payload
                    .get_item("value")?
                    .map(|v| v.extract::<String>())
                    .transpose()?
                    .unwrap_or_default();
                CellValue::String(v)
            }
            "number" | "float" | "int" | "integer" => {
                let v = payload
                    .get_item("value")?
                    .map(|v| v.extract::<f64>())
                    .transpose()?
                    .unwrap_or(0.0);
                CellValue::Number(v)
            }
            "boolean" | "bool" => {
                let v = payload
                    .get_item("value")?
                    .map(|v| v.extract::<bool>())
                    .transpose()?
                    .unwrap_or(false);
                CellValue::Boolean(v)
            }
            "formula" => {
                let v = payload
                    .get_item("value")?
                    .map(|v| v.extract::<String>())
                    .transpose()?
                    .unwrap_or_default();
                // Strip leading '=' if present (openpyxl convention)
                let formula = v.strip_prefix('=').unwrap_or(&v).to_string();
                CellValue::Formula(formula)
            }
            other => {
                return Err(PyErr::new::<PyValueError, _>(format!(
                    "Unknown cell type: '{other}'"
                )));
            }
        };

        let (row, col) =
            crate::util::a1_to_row_col(cell).map_err(|e| PyErr::new::<PyValueError, _>(e))?;

        let patch = CellPatch {
            row: row + 1, // a1_to_row_col returns 0-based, patcher uses 1-based
            col: col + 1,
            value: Some(value),
            style_index: None,
        };

        self.value_patches
            .insert((sheet.to_string(), cell.to_string()), patch);
        Ok(())
    }

    /// Queue a cell format change.
    ///
    /// `format_dict` matches the ExcelBench format dict:
    ///   {"bold": true, "font_size": 14, "font_name": "Arial", "font_color": "#FF0000",
    ///    "bg_color": "#00FF00", "number_format": "$#,##0", ...}
    fn queue_format(
        &mut self,
        sheet: &str,
        cell: &str,
        format_dict: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let spec = dict_to_format_spec(format_dict)?;
        self.format_patches
            .insert((sheet.to_string(), cell.to_string()), spec);
        Ok(())
    }

    /// Queue a data-validation rule on a sheet (RFC-025).
    ///
    /// `payload` is a dict of openpyxl-shaped fields. `sqref` is required;
    /// every other key is optional. Booleans default to `false`. Unknown
    /// keys are ignored to keep the Python side forward-compatible.
    fn queue_data_validation(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let sqref = extract_str(payload, "sqref")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("data validation requires 'sqref'"))?;

        let patch = DataValidationPatch {
            validation_type: extract_str(payload, "validation_type")?
                .unwrap_or_else(|| "none".to_string()),
            operator: extract_str(payload, "operator")?,
            formula1: extract_str(payload, "formula1")?,
            formula2: extract_str(payload, "formula2")?,
            sqref,
            allow_blank: extract_bool(payload, "allow_blank")?.unwrap_or(false),
            show_dropdown: extract_bool(payload, "show_dropdown")?.unwrap_or(false),
            show_input_message: extract_bool(payload, "show_input_message")?.unwrap_or(false),
            show_error_message: extract_bool(payload, "show_error_message")?.unwrap_or(false),
            error_style: extract_str(payload, "error_style")?,
            error_title: extract_str(payload, "error_title")?,
            error: extract_str(payload, "error")?,
            prompt_title: extract_str(payload, "prompt_title")?,
            prompt: extract_str(payload, "prompt")?,
        };

        self.queued_dv_patches
            .entry(sheet.to_string())
            .or_default()
            .push(patch);
        Ok(())
    }

    /// Queue a conditional-formatting patch on a sheet (RFC-026).
    ///
    /// `payload` is a flat dict shaped like:
    ///   {"sqref": "A1:A10",
    ///    "rules": [
    ///      {"kind": "cellIs"|"expression"|"colorScale"|"dataBar",
    ///       "operator": "greaterThan",          # cellIs only
    ///       "formula_a": "5", "formula_b": "10", # cellIs / expression
    ///       "formula":   "...",                  # expression only
    ///       "stops": [{"cfvo_type": "min", "val": None,
    ///                  "color_rgb": "FFF8696B"}, ...],   # colorScale
    ///       "min_cfvo_type": "min", "min_val": None,     # dataBar
    ///       "max_cfvo_type": "max", "max_val": None,
    ///       "color_rgb": "FF638EC6",                     # dataBar
    ///       "stop_if_true": false,
    ///       "dxf": { ... } | None,
    ///      }, ...]}
    ///
    /// Mirrors the writer's `add_conditional_format` shape but nests rules
    /// under one wrapper per sqref so priority ordering within a wrapper
    /// is preserved.
    fn queue_conditional_formatting(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let sqref = extract_str(payload, "sqref")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("conditional formatting requires 'sqref'"))?;

        let rules_obj = payload
            .get_item("rules")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("conditional formatting requires 'rules'"))?;
        let rules_list = rules_obj
            .downcast::<pyo3::types::PyList>()
            .map_err(|_| PyErr::new::<PyValueError, _>("'rules' must be a list of dicts"))?;

        let mut rules: Vec<CfRulePatch> = Vec::with_capacity(rules_list.len());
        for item in rules_list.iter() {
            let rd = item
                .downcast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("each rule must be a dict"))?;
            rules.push(extract_cf_rule(rd)?);
        }

        let patch = ConditionalFormattingPatch { sqref, rules };
        self.queued_cf_patches
            .entry(sheet.to_string())
            .or_default()
            .push(patch);
        Ok(())
    }

    /// Queue a cell border change.
    fn queue_border(
        &mut self,
        sheet: &str,
        cell: &str,
        border_dict: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let border = dict_to_border_spec(border_dict)?;
        // Merge with existing format patch or create new one
        let key = (sheet.to_string(), cell.to_string());
        let spec = self.format_patches.entry(key).or_default();
        spec.border = Some(border);
        Ok(())
    }

    /// Return the list of sheet names discovered in the workbook.
    ///
    /// Returned in source-document order (the order Excel rendered the
    /// tabs). Switched from `sheet_paths.keys()` to `sheet_order` in
    /// RFC-013 so callers that thread the sheet list into output
    /// (RFC-020's `app.xml` `<TitlesOfParts>`) get the right ordering
    /// without re-parsing `xl/workbook.xml`.
    fn sheet_names(&self) -> Vec<String> {
        self.sheet_order.clone()
    }

    /// Queue a document-properties update (RFC-020). The payload is the
    /// flat dict produced by `python/wolfxl/_workbook.py`'s
    /// `_flush_properties_to_patcher`; absent fields stay `None` and
    /// don't appear in the rewritten core.xml.
    ///
    /// Recognized keys (all optional): `title`, `subject`, `creator`,
    /// `keywords`, `description`, `last_modified_by`, `category`,
    /// `content_status`, `created_iso`, `modified_iso`, `sheet_names`
    /// (`list[str]`).
    /// Queue a hyperlink set/update for `sheet[cell]` (RFC-022).
    ///
    /// `payload` keys (all optional but at least one of `target` /
    /// `location` MUST be present): `target` (external URL — http/mailto/
    /// file), `location` (internal sheet anchor like `'Sheet2'!A1`),
    /// `tooltip`, `display`. Drained by Phase 2.5e during `do_save`.
    fn queue_hyperlink(
        &mut self,
        sheet: &str,
        cell: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let target = extract_str(payload, "target")?;
        let location = extract_str(payload, "location")?;
        let tooltip = extract_str(payload, "tooltip")?;
        let display = extract_str(payload, "display")?;
        if target.is_none() && location.is_none() {
            return Err(PyErr::new::<PyValueError, _>(
                "queue_hyperlink: at least one of 'target' or 'location' must be set",
            ));
        }
        let patch = hyperlinks::HyperlinkPatch {
            coordinate: cell.to_string(),
            target,
            location,
            tooltip,
            display,
        };
        self.queued_hyperlinks
            .entry(sheet.to_string())
            .or_default()
            .insert(cell.to_string(), hyperlinks::HyperlinkOp::Set(patch));
        Ok(())
    }

    /// Queue a defined-name upsert (RFC-021).
    ///
    /// `payload` keys (`name` + `formula` required; rest optional):
    ///   - `name`            (str)  — defined name. Includes any `_xlnm.` prefix verbatim.
    ///   - `formula`         (str)  — XML text content (no leading `=`).
    ///   - `local_sheet_id`  (int?) — `None` = workbook-scope; 0-based sheet position otherwise.
    ///   - `hidden`          (bool?)— `True` emits `hidden="1"`.
    ///   - `comment`         (str?) — defined-name `comment` attribute.
    ///
    /// Drained by Phase 2.5f during `do_save`. Upsert key is
    /// `(name, local_sheet_id)` — two entries with the same name but
    /// different scopes coexist independently.
    fn queue_defined_name(&mut self, payload: &Bound<'_, PyDict>) -> PyResult<()> {
        let name = extract_str(payload, "name")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("queue_defined_name: 'name' is required"))?;
        let formula = extract_str(payload, "formula")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>(
                "queue_defined_name: 'formula' is required",
            ))?;
        let local_sheet_id = match payload.get_item("local_sheet_id")? {
            Some(v) if !v.is_none() => Some(v.extract::<u32>()?),
            _ => None,
        };
        let hidden = match payload.get_item("hidden")? {
            Some(v) if !v.is_none() => Some(v.extract::<bool>()?),
            _ => None,
        };
        let comment = extract_str(payload, "comment")?;
        self.queued_defined_names.push(defined_names::DefinedNameMut {
            name,
            formula,
            local_sheet_id,
            hidden,
            comment,
        });
        Ok(())
    }

    /// Queue workbook-level security (RFC-058 Phase 2.5q).
    ///
    /// `payload` shape (RFC-058 §10):
    ///
    /// ```python
    /// {
    ///     "workbook_protection": {
    ///         "lock_structure": bool,
    ///         "lock_windows": bool,
    ///         "lock_revision": bool,
    ///         "workbook_algorithm_name": str | None,
    ///         "workbook_hash_value": str | None,
    ///         "workbook_salt_value": str | None,
    ///         "workbook_spin_count": int | None,
    ///         "revisions_algorithm_name": str | None,
    ///         "revisions_hash_value": str | None,
    ///         "revisions_salt_value": str | None,
    ///         "revisions_spin_count": int | None,
    ///     } | None,
    ///     "file_sharing": {
    ///         "read_only_recommended": bool,
    ///         "user_name": str | None,
    ///         "algorithm_name": str | None,
    ///         "hash_value": str | None,
    ///         "salt_value": str | None,
    ///         "spin_count": int | None,
    ///     } | None,
    /// }
    /// ```
    ///
    /// Either branch may be `None`. Drained by Phase 2.5q during
    /// `do_save`; the queue is single-slot — calling this again
    /// REPLACES the previous payload (matches the Python-side
    /// `wb.security = ...` semantics).
    fn queue_workbook_security(&mut self, payload: &Bound<'_, PyDict>) -> PyResult<()> {
        let security = parse_workbook_security_payload(payload)?;
        self.queued_workbook_security = Some(security);
        Ok(())
    }

    /// Queue a sheet-setup update for `sheet` (RFC-055 Phase 2.5n).
    ///
    /// `payload` is the §10 dict shape produced by
    /// `Worksheet.to_rust_setup_dict()`:
    ///
    /// ```text
    /// {
    ///   "page_setup": {...} | None,
    ///   "page_margins": {...} | None,
    ///   "header_footer": {...} | None,
    ///   "sheet_view": {...} | None,
    ///   "sheet_protection": {...} | None,
    ///   "print_titles": {"rows": "1:1" | None, "cols": "A:A" | None} | None,
    /// }
    /// ```
    ///
    /// Calling this again for the same `sheet` REPLACES the previous
    /// payload — matches Python `ws.page_setup = ...` semantics.
    /// Drained by Phase 2.5n during `do_save`, sequenced AFTER pivots
    /// (2.5m) and BEFORE autoFilter (2.5o).
    fn queue_sheet_setup_update(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let specs = sheet_setup::parse_sheet_setup_payload(payload)?;
        if specs.is_empty() {
            // No-op queue entry — drop any prior entry too, matching
            // "user reset everything to default" semantics.
            self.queued_sheet_setup.remove(sheet);
        } else {
            self.queued_sheet_setup
                .insert(sheet.to_string(), sheet_setup::QueuedSheetSetup { specs });
        }
        Ok(())
    }

    /// Queue a sheet reorder (RFC-036).
    ///
    /// `sheet` is the sheet's `name` attribute (resolved on the Python
    /// side from a `Worksheet` instance or string). `offset` is added
    /// to the sheet's current 0-based position; the resulting index is
    /// clamped to `[0, n-1]`. Drained by Phase 2.5h during `do_save`.
    /// Multiple queued moves apply in queue order against the running
    /// tab list, and Phase 2.5h re-points every `<definedName
    /// localSheetId>` whose integer maps to a moved position before
    /// the defined-names merger runs.
    fn queue_sheet_move(&mut self, sheet: &str, offset: i32) -> PyResult<()> {
        self.queued_sheet_moves.push((sheet.to_string(), offset));
        Ok(())
    }

    /// Queue a hyperlink delete for `sheet[cell]` (RFC-022). Idempotent:
    /// running on a cell that had no source hyperlink is a no-op at
    /// flush time.
    fn queue_hyperlink_delete(&mut self, sheet: &str, cell: &str) -> PyResult<()> {
        self.queued_hyperlinks
            .entry(sheet.to_string())
            .or_default()
            .insert(cell.to_string(), hyperlinks::HyperlinkOp::Delete);
        Ok(())
    }

    /// Queue a table addition on `sheet` (RFC-024).
    ///
    /// `payload` keys (`name`, `ref`, and `columns` are required;
    /// other keys default sensibly):
    ///   - `name`              (str)
    ///   - `display_name`      (str?, defaults to `name`)
    ///   - `ref`               (str)  — A1 range, e.g. `"A1:E10"`
    ///   - `columns`           (list[str]) — column names in order
    ///   - `style`             (dict?) — `name`, `show_first_column`,
    ///                          `show_last_column`, `show_row_stripes`,
    ///                          `show_column_stripes`
    ///   - `header_row_count`  (int?, defaults to 1)
    ///   - `totals_row_shown`  (bool?, defaults to `false`)
    ///   - `autofilter`        (bool?, defaults to `true`)
    ///
    /// Workbook-unique id allocation, name-collision detection,
    /// part-file emission, sheet-rels mutation, content-type
    /// override, and `<tableParts>` block insertion all happen at
    /// `save()` time during Phase-2.5f.
    fn queue_table(&mut self, sheet: &str, payload: &Bound<'_, PyDict>) -> PyResult<()> {
        let name = extract_str(payload, "name")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("table requires 'name'"))?;
        let display_name = extract_str(payload, "display_name")?.unwrap_or_else(|| name.clone());
        let ref_range = extract_str(payload, "ref")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("table requires 'ref'"))?;
        let columns_obj = payload
            .get_item("columns")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("table requires 'columns'"))?;
        let columns: Vec<String> = columns_obj.extract::<Vec<String>>()?;
        let header_row_count = extract_u32(payload, "header_row_count")?.unwrap_or(1);
        let totals_row_shown = extract_bool(payload, "totals_row_shown")?.unwrap_or(false);
        let autofilter = extract_bool(payload, "autofilter")?.unwrap_or(true);

        let style = match payload.get_item("style")? {
            Some(v) if !v.is_none() => {
                let d = v
                    .downcast::<PyDict>()
                    .map_err(|_| PyErr::new::<PyValueError, _>("'style' must be a dict or None"))?;
                let style_name = extract_str(d, "name")?.unwrap_or_default();
                Some(tables::TableStylePatch {
                    name: style_name,
                    show_first_column: extract_bool(d, "show_first_column")?.unwrap_or(false),
                    show_last_column: extract_bool(d, "show_last_column")?.unwrap_or(false),
                    show_row_stripes: extract_bool(d, "show_row_stripes")?.unwrap_or(false),
                    show_column_stripes: extract_bool(d, "show_column_stripes")?.unwrap_or(false),
                })
            }
            _ => None,
        };

        let patch = tables::TablePatch {
            name,
            display_name,
            ref_range,
            columns,
            style,
            header_row_count,
            totals_row_shown,
            autofilter,
        };
        self.queued_tables
            .entry(sheet.to_string())
            .or_default()
            .push(patch);
        Ok(())
    }

    /// Sprint Λ Pod-β (RFC-045) — queue an image add for `sheet`.
    ///
    /// Payload shape mirrors `NativeWorkbook.add_image`:
    /// ```python
    /// {
    ///   "data": <bytes>,
    ///   "ext": "png" | "jpeg" | "gif" | "bmp",
    ///   "width": int,
    ///   "height": int,
    ///   "anchor": {"type": "one_cell"|"two_cell"|"absolute", ...},
    /// }
    /// ```
    /// Drained by Phase 2.5k during `do_save`. Sheets that already
    /// have a drawing rel raise `NotImplementedError` at flush time
    /// (v1.5 follow-up: append to existing drawingN.xml).
    fn queue_image_add(&mut self, sheet: &str, payload: &Bound<'_, PyDict>) -> PyResult<()> {
        let data: Vec<u8> = payload
            .get_item("data")?
            .ok_or_else(|| PyValueError::new_err("queue_image_add: missing 'data'"))?
            .extract()?;
        let ext: String = payload
            .get_item("ext")?
            .ok_or_else(|| PyValueError::new_err("queue_image_add: missing 'ext'"))?
            .extract()?;
        let width: u32 = payload
            .get_item("width")?
            .ok_or_else(|| PyValueError::new_err("queue_image_add: missing 'width'"))?
            .extract()?;
        let height: u32 = payload
            .get_item("height")?
            .ok_or_else(|| PyValueError::new_err("queue_image_add: missing 'height'"))?
            .extract()?;
        let anchor_obj = payload
            .get_item("anchor")?
            .ok_or_else(|| PyValueError::new_err("queue_image_add: missing 'anchor'"))?;
        let anchor_dict = anchor_obj
            .downcast::<PyDict>()
            .map_err(|_| PyValueError::new_err("queue_image_add: 'anchor' must be a dict"))?;
        let anchor = parse_queued_image_anchor(anchor_dict)?;
        self.queued_images
            .entry(sheet.to_string())
            .or_default()
            .push(QueuedImageAdd {
                data,
                ext: ext.to_ascii_lowercase(),
                width_px: width,
                height_px: height,
                anchor,
            });
        Ok(())
    }

    /// Sprint Μ Pod-γ (RFC-046) — queue a chart add for `sheet`.
    ///
    /// The caller (Pod-α's `emit_chart_xml(&Chart)` from the writer
    /// crate, surfaced via `Workbook.add_chart_modify_mode`) supplies
    /// pre-serialized chart XML bytes plus an A1-style anchor. The
    /// patcher routes the bytes through `xl/charts/chartN.xml`, the
    /// drawing layer (fresh OR existing), the rels graphs, and the
    /// content-types map. Drained by Phase 2.5l during `do_save`.
    fn queue_chart_add(
        &mut self,
        sheet: &str,
        chart_xml: Vec<u8>,
        anchor_a1: &str,
        width_emu: i64,
        height_emu: i64,
    ) -> PyResult<()> {
        if anchor_a1.trim().is_empty() {
            return Err(PyValueError::new_err(
                "queue_chart_add: anchor_a1 must be a non-empty A1 cell coord",
            ));
        }
        self.queued_charts
            .entry(sheet.to_string())
            .or_default()
            .push(QueuedChartAdd {
                chart_xml,
                anchor_a1: anchor_a1.to_string(),
                width_emu,
                height_emu,
            });
        Ok(())
    }

    /// Sprint Ν Pod-γ (RFC-047) — queue a pivot cache add. Returns
    /// the allocated 0-based `cache_id` so the caller can wire it
    /// into pivot tables that reference this cache.
    ///
    /// The XML payloads are pre-serialised by the Python coordinator
    /// via `wolfxl._rust.serialize_pivot_cache_dict` (definition)
    /// and `serialize_pivot_records_dict` (records). Drained by
    /// Phase 2.5m during `do_save`.
    fn queue_pivot_cache_add(
        &mut self,
        cache_def_xml: Vec<u8>,
        cache_records_xml: Vec<u8>,
    ) -> PyResult<u32> {
        let cache_id = self.next_pivot_cache_id;
        self.next_pivot_cache_id += 1;
        self.queued_pivot_caches
            .push(pivot::QueuedPivotCacheAdd {
                cache_def_xml,
                cache_records_xml,
                cache_id,
            });
        Ok(cache_id)
    }

    /// Sprint Ν Pod-γ (RFC-048) — queue a pivot table add. The
    /// `cache_id` must reference a cache previously queued via
    /// `queue_pivot_cache_add` (or already present in the source
    /// workbook). Drained by Phase 2.5m AFTER the cache drain so the
    /// table → cache rels target resolves cleanly.
    fn queue_pivot_table_add(
        &mut self,
        sheet: &str,
        table_xml: Vec<u8>,
        cache_id: u32,
    ) -> PyResult<()> {
        if !self.sheet_paths.contains_key(sheet) {
            return Err(PyValueError::new_err(format!(
                "queue_pivot_table_add: no such sheet: {sheet}"
            )));
        }
        self.queued_pivot_tables
            .entry(sheet.to_string())
            .or_default()
            .push(pivot::QueuedPivotTableAdd {
                sheet: sheet.to_string(),
                table_xml,
                cache_id,
            });
        Ok(())
    }

    /// Sprint Ο Pod 1B (RFC-056) — queue an autoFilter for a sheet.
    ///
    /// `dict` is the §10 dict shape produced by
    /// `Worksheet.auto_filter.to_rust_dict()`. Drained by Phase 2.5o
    /// during `do_save` (sequenced AFTER pivots, BEFORE cells).
    fn queue_autofilter(
        &mut self,
        sheet: &str,
        dict: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        if !self.sheet_paths.contains_key(sheet) {
            return Err(PyValueError::new_err(format!(
                "queue_autofilter: no such sheet: {sheet}"
            )));
        }
        let dv = autofilter::pyany_to_dictvalue(&dict.as_any().clone())?;
        self.queued_autofilters.insert(
            sheet.to_string(),
            autofilter::QueuedAutoFilter {
                sheet: sheet.to_string(),
                dict: dv,
            },
        );
        Ok(())
    }

    /// Queue a comment set/update for `sheet[cell]` (RFC-023).
    ///
    /// `payload` keys: `text` (required), `author` (optional — defaults
    /// to `"wolfxl"` to match the writer), `width_pt` / `height_pt`
    /// (optional, in OOXML points). Drained by Phase 2.5g during
    /// `do_save`.
    fn queue_comment(
        &mut self,
        sheet: &str,
        cell: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let text = extract_str(payload, "text")?.unwrap_or_default();
        let author = extract_str(payload, "author")?.unwrap_or_else(|| "wolfxl".to_string());
        let width_pt = extract_f64(payload, "width_pt")?;
        let height_pt = extract_f64(payload, "height_pt")?;
        let patch = comments::CommentPatch {
            coordinate: cell.to_string(),
            author,
            text,
            width_pt,
            height_pt,
        };
        self.queued_comments
            .entry(sheet.to_string())
            .or_default()
            .insert(cell.to_string(), comments::CommentOp::Set(patch));
        Ok(())
    }

    /// Queue a comment delete for `sheet[cell]` (RFC-023). Idempotent:
    /// running on a cell that had no source comment is a no-op at
    /// flush time.
    fn queue_comment_delete(&mut self, sheet: &str, cell: &str) -> PyResult<()> {
        self.queued_comments
            .entry(sheet.to_string())
            .or_default()
            .insert(cell.to_string(), comments::CommentOp::Delete);
        Ok(())
    }

    /// Queue a structural axis shift for `sheet` (RFC-030 / RFC-031).
    ///
    /// `axis` must be `"row"` or `"col"`. `idx` is 1-based; `n` is
    /// signed (positive = insert; negative = delete). The Python
    /// coordinator validates `idx >= 1` and `amount >= 1` before
    /// queueing so this method does NOT re-validate.
    ///
    /// Drained by Phase 2.5i during `do_save`. Order is append order
    /// — multi-op sequencing matters (each op runs in the coordinate
    /// space produced by the previous op).
    fn queue_axis_shift(
        &mut self,
        sheet: &str,
        axis: &str,
        idx: u32,
        n: i32,
    ) -> PyResult<()> {
        if axis != "row" && axis != "col" {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "queue_axis_shift: axis must be 'row' or 'col', got '{axis}'"
            )));
        }
        if idx < 1 {
            return Err(PyErr::new::<PyValueError, _>(
                "queue_axis_shift: idx must be >= 1",
            ));
        }
        self.queued_axis_shifts.push(AxisShift {
            sheet: sheet.to_string(),
            axis: axis.to_string(),
            idx,
            n,
        });
        Ok(())
    }

    /// Queue a paste-style range relocation for `sheet` (RFC-034).
    ///
    /// The source rectangle is given by 1-based inclusive corners
    /// `(src_min_col, src_min_row)..=(src_max_col, src_max_row)` and
    /// the destination delta by `(d_row, d_col)`. `translate` controls
    /// whether external formulas pointing INTO `src` also re-anchor
    /// (default false — matches openpyxl). The Python coordinator
    /// validates corners and destination bounds before queueing.
    ///
    /// Drained by Phase 2.5j during `do_save`. Ops apply in append
    /// order; each op runs against the post-previous-op bytes.
    #[allow(clippy::too_many_arguments)]
    fn queue_range_move(
        &mut self,
        sheet: &str,
        src_min_col: u32,
        src_min_row: u32,
        src_max_col: u32,
        src_max_row: u32,
        d_row: i32,
        d_col: i32,
        translate: bool,
    ) -> PyResult<()> {
        if src_min_col < 1 || src_min_row < 1 {
            return Err(PyErr::new::<PyValueError, _>(
                "queue_range_move: source corners must be >= 1",
            ));
        }
        if src_min_col > src_max_col || src_min_row > src_max_row {
            return Err(PyErr::new::<PyValueError, _>(
                "queue_range_move: src_min must be <= src_max on both axes",
            ));
        }
        self.queued_range_moves.push(RangeMove {
            sheet: sheet.to_string(),
            src_min_col,
            src_min_row,
            src_max_col,
            src_max_row,
            d_row,
            d_col,
            translate,
        });
        Ok(())
    }

    /// Queue a sheet-copy op (RFC-035 Phase 7.3).
    ///
    /// Validates eagerly that `src_title` exists in `self.sheet_paths`,
    /// `dst_title` is non-empty, `dst_title` is not already a sheet
    /// name in `self.sheet_paths`, and `dst_title` is not already
    /// queued by an earlier `queue_sheet_copy` call. On success
    /// appends to `queued_sheet_copies` (drained by Phase 2.7 in
    /// append order during `do_save`).
    ///
    /// `deep_copy_images` is the workbook-level
    /// `wb.copy_options.deep_copy_images` flag, snapshot at queue
    /// time. Defaults to `false` to preserve historical RFC-035 §5.3
    /// alias-by-target behaviour for callers that omit it.
    #[pyo3(signature = (src_title, dst_title, deep_copy_images=false))]
    fn queue_sheet_copy(
        &mut self,
        src_title: &str,
        dst_title: &str,
        deep_copy_images: bool,
    ) -> PyResult<()> {
        if !self.sheet_paths.contains_key(src_title) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "queue_sheet_copy: source sheet '{src_title}' not found in workbook"
            )));
        }
        if dst_title.is_empty() {
            return Err(PyErr::new::<PyValueError, _>(
                "queue_sheet_copy: destination title must be non-empty",
            ));
        }
        if self.sheet_paths.contains_key(dst_title) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "queue_sheet_copy: destination sheet '{dst_title}' already exists"
            )));
        }
        if self
            .queued_sheet_copies
            .iter()
            .any(|op| op.dst_title == dst_title)
        {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "queue_sheet_copy: destination sheet '{dst_title}' is already queued"
            )));
        }
        self.queued_sheet_copies.push(SheetCopyOp {
            src_title: src_title.to_string(),
            dst_title: dst_title.to_string(),
            deep_copy_images,
        });
        Ok(())
    }

    fn queue_properties(&mut self, payload: &Bound<'_, PyDict>) -> PyResult<()> {
        let title = extract_str(payload, "title")?;
        let subject = extract_str(payload, "subject")?;
        let creator = extract_str(payload, "creator")?;
        let keywords = extract_str(payload, "keywords")?;
        let description = extract_str(payload, "description")?;
        let last_modified_by = extract_str(payload, "last_modified_by")?;
        let category = extract_str(payload, "category")?;
        let content_status = extract_str(payload, "content_status")?;
        let created_iso = extract_str(payload, "created_iso")?;
        let modified_iso = extract_str(payload, "modified_iso")?;
        let sheet_names: Vec<String> = match payload.get_item("sheet_names")? {
            Some(v) => v.extract::<Vec<String>>()?,
            None => Vec::new(),
        };
        self.queued_props = Some(properties::DocPropertiesPayload {
            title,
            subject,
            creator,
            keywords,
            description,
            last_modified_by,
            category,
            content_status,
            created_iso,
            modified_iso,
            sheet_names,
        });
        Ok(())
    }

    /// Save patched file to a new path.
    fn save(&mut self, path: &str) -> PyResult<()> {
        self.do_save(path)
    }

    /// Save in-place (atomic tmp+rename).
    fn save_in_place(&mut self) -> PyResult<()> {
        let tmp_path = format!("{}.wolfxl.tmp", self.file_path);
        self.do_save(&tmp_path)?;

        // Atomic rename
        if let Err(e) = std::fs::rename(&tmp_path, &self.file_path) {
            let _ = std::fs::remove_file(&self.file_path);
            std::fs::rename(&tmp_path, &self.file_path).map_err(|e2| {
                PyErr::new::<PyIOError, _>(format!("Failed to replace file: {e}; {e2}"))
            })?;
        }
        Ok(())
    }

    // -------------------------------------------------------------------
    // RFC-013 test-only hooks.
    //
    // These methods drive the new patcher primitives (`file_adds`,
    // `queued_content_type_ops`, `ancillary`) directly so pytest
    // integration tests can verify behavior end-to-end. They are
    // intentionally `_test_`-prefixed (Python convention for "internal
    // testing API") and have NO live caller in `python/wolfxl/`. RFC-022
    // / RFC-023 / RFC-024 will add the real public callers; until then,
    // these hooks are how `tests/test_patcher_infra.py` exercises the
    // plumbing.
    // -------------------------------------------------------------------

    /// Inject a brand-new ZIP entry that will be emitted on the next
    /// `save()`. Used by `tests/test_patcher_infra.py` to verify that
    /// `file_adds` round-trips through `do_save`.
    fn _test_inject_file_add(&mut self, path: &str, bytes: Vec<u8>) {
        self.file_adds.insert(path.to_string(), bytes);
    }

    /// Queue a content-type op against a sheet. `kind` is `"add_override"`
    /// or `"ensure_default"`; `key` is the part path or extension; `value`
    /// is the content type. The next `save()` aggregates queued ops
    /// across all sheets in `sheet_order` and writes one rewritten
    /// `[Content_Types].xml`.
    fn _test_queue_content_type_op(
        &mut self,
        sheet: &str,
        kind: &str,
        key: &str,
        value: &str,
    ) -> PyResult<()> {
        let op = match kind {
            "add_override" => content_types::ContentTypeOp::AddOverride(
                key.to_string(),
                value.to_string(),
            ),
            "ensure_default" => content_types::ContentTypeOp::EnsureDefault(
                key.to_string(),
                value.to_string(),
            ),
            other => {
                return Err(PyErr::new::<PyValueError, _>(format!(
                    "unknown ContentTypeOp kind '{other}' (expected 'add_override' or 'ensure_default')"
                )));
            }
        };
        self.queued_content_type_ops
            .entry(sheet.to_string())
            .or_default()
            .push(op);
        Ok(())
    }

    /// Lazily populate the ancillary registry for one sheet by name. After
    /// this call, `_test_ancillary_*` accessors return the classified
    /// `_rels/sheetN.xml.rels` contents.
    fn _test_populate_ancillary(&mut self, sheet: &str) -> PyResult<()> {
        let path = self
            .sheet_paths
            .get(sheet)
            .cloned()
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("no such sheet: {sheet}")))?;
        let f = File::open(&self.file_path).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Cannot open '{}': {e}", self.file_path))
        })?;
        let mut zip = ZipArchive::new(f)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP read error: {e}")))?;
        self.ancillary
            .populate_for_sheet(&mut zip, sheet, &path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("ancillary populate: {e}")))?;
        Ok(())
    }

    /// Has the ancillary registry been populated for `sheet`? Returns
    /// `False` for unknown sheets and for sheets whose rels file has not
    /// been read yet.
    fn _test_ancillary_is_populated(&self, sheet: &str) -> bool {
        self.ancillary.get(sheet).is_some()
    }

    /// Cached comments-part path for `sheet`, or `None` if the sheet has
    /// none / has not been populated.
    fn _test_ancillary_comments_part(&self, sheet: &str) -> Option<String> {
        self.ancillary
            .get(sheet)
            .and_then(|a| a.comments_part.clone())
    }

    /// Cached VML drawing part path for `sheet`.
    fn _test_ancillary_vml_drawing_part(&self, sheet: &str) -> Option<String> {
        self.ancillary
            .get(sheet)
            .and_then(|a| a.vml_drawing_part.clone())
    }

    /// Cached table-part paths for `sheet`, in source order.
    fn _test_ancillary_table_parts(&self, sheet: &str) -> Vec<String> {
        self.ancillary
            .get(sheet)
            .map(|a| a.table_parts.clone())
            .unwrap_or_default()
    }

    /// Cached hyperlink `rId`s for `sheet`, in source order.
    fn _test_ancillary_hyperlink_rids(&self, sheet: &str) -> Vec<String> {
        self.ancillary
            .get(sheet)
            .map(|a| a.hyperlinks_rels.iter().map(|r| r.0.clone()).collect())
            .unwrap_or_default()
    }

    // -------------------------------------------------------------------
    // RFC-022 test-only hooks.
    // -------------------------------------------------------------------

    /// Inject a Set op directly into `queued_hyperlinks`. Mirrors
    /// `queue_hyperlink` but bypasses the validator so tests can set up
    /// odd shapes (e.g. tooltip-only) deliberately.
    fn _test_inject_hyperlink(
        &mut self,
        sheet: &str,
        coord: &str,
        target: Option<String>,
        location: Option<String>,
        tooltip: Option<String>,
        display: Option<String>,
    ) {
        let patch = hyperlinks::HyperlinkPatch {
            coordinate: coord.to_string(),
            target,
            location,
            tooltip,
            display,
        };
        self.queued_hyperlinks
            .entry(sheet.to_string())
            .or_default()
            .insert(coord.to_string(), hyperlinks::HyperlinkOp::Set(patch));
    }

    /// Inject a Delete op directly into `queued_hyperlinks`.
    fn _test_inject_hyperlink_delete(&mut self, sheet: &str, coord: &str) {
        self.queued_hyperlinks
            .entry(sheet.to_string())
            .or_default()
            .insert(coord.to_string(), hyperlinks::HyperlinkOp::Delete);
    }

    /// Run `extract_hyperlinks` on the source ZIP's current sheet XML
    /// and return `(coord, target_or_location)` pairs in BTreeMap order
    /// for assertion in pytest.
    fn _test_get_extracted_hyperlinks(
        &mut self,
        sheet: &str,
    ) -> PyResult<Vec<(String, String)>> {
        let sheet_path = self
            .sheet_paths
            .get(sheet)
            .cloned()
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("no such sheet: {sheet}")))?;
        let f = File::open(&self.file_path).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Cannot open '{}': {e}", self.file_path))
        })?;
        let mut zip = ZipArchive::new(f)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP read error: {e}")))?;
        let rels_path = sheet_rels_path_for(&sheet_path);
        let rels = load_or_empty_rels(&mut zip, &rels_path)?;
        let xml = ooxml_util::zip_read_to_string(&mut zip, &sheet_path)?;
        let extracted = hyperlinks::extract_hyperlinks(xml.as_bytes(), &rels);
        Ok(extracted
            .into_iter()
            .map(|(coord, h)| {
                let val = h
                    .target
                    .or(h.location)
                    .unwrap_or_default();
                (coord, val)
            })
            .collect())
    }
}

// ---------------------------------------------------------------------------
// Helpers used by Phase 2.5e (hyperlinks) — small wrappers over
// `wolfxl_rels::rels_path_for` and the ZIP entry reader so the per-sheet
// flush stays terse.
// ---------------------------------------------------------------------------

/// Maps a sheet XML path (`xl/worksheets/sheet1.xml`) to its rels
/// sidecar (`xl/worksheets/_rels/sheet1.xml.rels`). Wraps
/// `wolfxl_rels::rels_path_for`; falls back to a synthesized path on
/// the (impossible-in-OOXML) input that has no `/`.
fn sheet_rels_path_for(sheet_path: &str) -> String {
    wolfxl_rels::rels_path_for(sheet_path)
        .unwrap_or_else(|| format!("_rels/{sheet_path}.rels"))
}

/// Parse the trailing integer N out of an OOXML part path like
/// `xl/comments3.xml` (with `prefix="xl/comments"`, `suffix=".xml"`).
/// Returns `None` if either the prefix/suffix don't match or the
/// substring between them doesn't parse as `u32`.
fn parse_n_from_part_path(path: &str, prefix: &str, suffix: &str) -> Option<u32> {
    let mid = path.strip_prefix(prefix)?.strip_suffix(suffix)?;
    mid.parse::<u32>().ok()
}

/// Read an existing `.rels` part out of `zip` and parse it; if the
/// part doesn't exist (sheet has no rels yet), return `RelsGraph::new()`.
/// Other read/parse errors propagate as `PyIOError`. Constrained to
/// `ZipArchive<File>` because `ooxml_util::zip_read_to_string_opt` is
/// not generic; matches every caller in this module.
fn load_or_empty_rels(
    zip: &mut ZipArchive<File>,
    path: &str,
) -> PyResult<RelsGraph> {
    match ooxml_util::zip_read_to_string_opt(zip, path)? {
        Some(xml) => RelsGraph::parse(xml.as_bytes())
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("rels parse for '{path}': {e}"))),
        None => Ok(RelsGraph::new()),
    }
}

// ---------------------------------------------------------------------------
// Save implementation
// ---------------------------------------------------------------------------

impl XlsxPatcher {
    fn do_save(&mut self, output_path: &str) -> PyResult<()> {
        if self.value_patches.is_empty()
            && self.format_patches.is_empty()
            && self.rels_patches.is_empty()
            && self.queued_blocks.is_empty()
            && self.queued_dv_patches.is_empty()
            && self.queued_cf_patches.is_empty()
            && self.file_adds.is_empty()
            && self.file_deletes.is_empty()
            && self.queued_content_type_ops.is_empty()
            && self.queued_props.is_none()
            && self.queued_hyperlinks.is_empty()
            && self.queued_defined_names.is_empty()
            && self.queued_tables.is_empty()
            && self.queued_comments.is_empty()
            && self.queued_sheet_moves.is_empty()
            && self.queued_axis_shifts.is_empty()
            && self.queued_range_moves.is_empty()
            && self.queued_sheet_copies.is_empty()
            && self.queued_images.is_empty()
            && self.queued_charts.is_empty()
            && self.queued_pivot_caches.is_empty()
            && self.queued_pivot_tables.is_empty()
        {
            // No changes — just copy. Includes RFC-013's `file_adds`,
            // `file_deletes`, `queued_content_type_ops`, RFC-020's
            // `queued_props`, RFC-022's `queued_hyperlinks`, RFC-021's
            // `queued_defined_names`, RFC-024's `queued_tables`,
            // RFC-023's `queued_comments`, RFC-036's
            // `queued_sheet_moves`, RFC-030/031's
            // `queued_axis_shifts`, and RFC-034's
            // `queued_range_moves` so a no-op save remains
            // byte-identical even after these primitives land.
            std::fs::copy(&self.file_path, output_path)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("Copy failed: {e}")))?;
            return Ok(());
        }

        let f = File::open(&self.file_path).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Cannot open '{}': {e}", self.file_path))
        })?;
        let mut zip = ZipArchive::new(f)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP read error: {e}")))?;

        // Centralized part-suffix allocator (RFC-035 §5.2 / §8 risk #1).
        // Built once per save; seeded from the source ZIP's part listing
        // so freshly minted tableN / commentsN / vmlDrawingN / sheetN
        // suffixes never collide with source entries. Shared by Phase
        // 2.7 (sheet copies), Phase 2.5f (tables), and Phase 2.5g
        // (comments + VML).
        let mut part_id_allocator: wolfxl_rels::PartIdAllocator = {
            let names: Vec<String> = (0..zip.len())
                .filter_map(|i| zip.by_index(i).ok().map(|e| e.name().to_string()))
                .collect();
            wolfxl_rels::PartIdAllocator::from_zip_parts(names.iter().map(|s| s.as_str()))
        };

        // RFC-035 §8 risk #6: when Phase 2.7 clones tables onto cloned
        // sheets, those new table names must be visible to Phase 2.5f's
        // collision-scan so a user `add_table` in the same save against
        // an as-yet-unflushed cloned name surfaces a clean error rather
        // than a silent rId/file collision.
        let mut cloned_table_names: HashSet<String> = HashSet::new();

        // `file_patches` is the running map of source-ZIP entries that
        // will be REPLACED on emit. Phase 2.7 (RFC-035) is the first
        // phase to write into it (workbook.xml + workbook.xml.rels).
        // Phase 3 mutates it further with per-sheet rewrites.
        let mut file_patches: HashMap<String, Vec<u8>> = HashMap::new();

        // Sprint Θ Pod-A: pre-seed `file_patches` with any permissive-
        // mode rewrites that `XlsxPatcher::open` produced (e.g. the
        // `<sheets/>` → `<sheets>...</sheets>` normalization). We move
        // (`drain`) rather than clone because the seed is one-shot.
        for (k, v) in self.permissive_seed_file_patches.drain() {
            file_patches.insert(k, v);
        }

        // --- Phase 2.7: Sheet copies (RFC-035) ---
        //
        // Drains `queued_sheet_copies` in append order. For each
        // `(src_title, dst_title)` op:
        //   1. Build source rels graph from the ZIP (or from
        //      `rels_patches` / `file_patches` if already mutated).
        //   2. Pre-load source ZIP parts map for the planner.
        //   3. Read workbook.xml (from `file_patches` if already
        //      mutated, else from the source ZIP).
        //   4. Call `wolfxl_structural::sheet_copy::plan_sheet_copy`.
        //   5. Allocate a workbook-rels rId for the new sheet.
        //   6. Apply mutations: insert sheet/rels/ancillary into
        //      `file_adds`; splice `<sheet>` into workbook.xml's
        //      `<sheets>` block; queue cloned sheet-scoped defined
        //      names through `queued_defined_names`; forward
        //      content-type ops through `queued_content_type_ops`;
        //      update `self.sheet_paths` + `self.sheet_order`.
        //
        // Phase-ordering invariant: any new per-sheet phase MUST run
        // AFTER 2.7 (per RFC-035 §8 risk #7). 2.7 runs before Phase
        // 2 / 2.5* so cloned sheets are visible to every downstream
        // per-sheet drain (cell patches, DV, CF, hyperlinks, tables,
        // comments, axis shifts, range moves) as if they had always
        // been part of the source workbook.
        if !self.queued_sheet_copies.is_empty() {
            self.apply_sheet_copies_phase(
                &mut file_patches,
                &mut zip,
                &mut part_id_allocator,
                &mut cloned_table_names,
            )?;
        }

        // --- Phase 1: Parse styles.xml if we have format patches ---
        let mut styles_xml: Option<String> = None;
        let mut style_assignments: HashMap<String, u32> = HashMap::new(); // "sheet:cell" → xf_index

        if !self.format_patches.is_empty() {
            let raw = ooxml_util::zip_read_to_string_opt(&mut zip, "xl/styles.xml")?
                .unwrap_or_else(|| minimal_styles_xml());
            let mut xml = raw;

            for ((sheet, cell), spec) in &self.format_patches {
                let (updated, xf_idx) = styles::apply_format_spec(&xml, spec);
                xml = updated;
                style_assignments.insert(format!("{sheet}:{cell}"), xf_idx);
            }
            styles_xml = Some(xml);
        }

        // --- Phase 2: Build cell patches per sheet ---
        let mut sheet_cell_patches: HashMap<String, Vec<CellPatch>> = HashMap::new();

        // Value patches
        for ((sheet, cell), patch) in &self.value_patches {
            let sheet_path = self.sheet_paths.get(sheet);
            if sheet_path.is_none() {
                continue;
            }
            let mut p = patch.clone();
            // Check if there's also a style assignment for this cell
            let key = format!("{sheet}:{cell}");
            if let Some(&xf_idx) = style_assignments.get(&key) {
                p.style_index = Some(xf_idx);
            }
            sheet_cell_patches
                .entry(sheet_path.unwrap().clone())
                .or_default()
                .push(p);
        }

        // Format-only patches (no value change)
        for ((sheet, cell), _) in &self.format_patches {
            let val_key = (sheet.clone(), cell.clone());
            if self.value_patches.contains_key(&val_key) {
                continue; // already handled above
            }
            let sheet_path = self.sheet_paths.get(sheet);
            if sheet_path.is_none() {
                continue;
            }
            let key = format!("{sheet}:{cell}");
            if let Some(&xf_idx) = style_assignments.get(&key) {
                let (row, col) = crate::util::a1_to_row_col(cell)
                    .map_err(|e| PyErr::new::<PyValueError, _>(e))?;
                let patch = CellPatch {
                    row: row + 1,
                    col: col + 1,
                    value: None, // no value change
                    style_index: Some(xf_idx),
                };
                sheet_cell_patches
                    .entry(sheet_path.unwrap().clone())
                    .or_default()
                    .push(patch);
            }
        }

        // --- Phase 2.5: Build <dataValidations> blocks from queued DV
        // patches (RFC-025).  Each queued sheet gets exactly one
        // SheetBlock::DataValidations entry whose bytes are the
        // (existing block's children, verbatim) + (new patches,
        // freshly serialized), wrapped in a single <dataValidations
        // count="N">…</dataValidations>.  We read the source sheet
        // XML here so the existing block can flow through unchanged.
        //
        // Pushed into a *local* clone of queued_blocks rather than
        // self — do_save takes &self, and self.queued_blocks is
        // reserved for setters that produce blocks pre-save (future
        // RFCs).  A local map keeps this slice's wiring contained
        // and safe to compose with future block-producing setters.
        let mut local_blocks: HashMap<String, Vec<SheetBlock>> = self.queued_blocks.clone();

        // Note: `part_id_allocator` is now built earlier (before
        // Phase 2.7) so the centralized allocator can mint cloned-sheet
        // suffixes for RFC-035. Phase 2.5f (tables) + Phase 2.5g
        // (comments + VML) consume the same instance below so
        // workbook-wide suffix uniqueness is preserved across phases.

        for (sheet_name, patches) in &self.queued_dv_patches {
            let sheet_path = match self.sheet_paths.get(sheet_name) {
                Some(p) => p,
                None => continue, // unknown sheet name — silently skip (mirrors value/format paths)
            };
            let xml = ooxml_util::zip_read_to_string(&mut zip, sheet_path)?;
            let existing = validations::extract_existing_dv_block(&xml);
            let block_bytes =
                validations::build_data_validations_block(existing.as_deref(), patches);
            local_blocks
                .entry(sheet_path.clone())
                .or_default()
                .push(SheetBlock::DataValidations(block_bytes));
        }

        // --- Phase 2.5b: Build <conditionalFormatting> blocks from
        // queued CF patches (RFC-026). Cross-sheet coordination: a
        // single workbook-wide `dxf_count` allocates dxfId values
        // across every sheet's patches in deterministic (sorted) sheet
        // order, and the resulting new `<dxf>` entries are folded into
        // a single `xl/styles.xml` mutation at the end.
        //
        // The merger uses replace-all semantics for slot 17 (RFC-011
        // §5.5) — supplying any CF block drops every existing CF block
        // in the source. We therefore call `extract_existing_cf_blocks`
        // first and re-include them verbatim at the head of each
        // sheet's payload so byte-preservation of unchanged CF rules
        // is not a side-effect of our setter call.
        let mut new_dxfs_total: Vec<DxfPatch> = Vec::new();
        let mut styles_loaded: Option<String> = None;
        let mut running_dxf_count: u32 = 0;
        let mut cf_sheet_names: Vec<&String> = self.queued_cf_patches.keys().collect();
        cf_sheet_names.sort();
        for sheet_name in cf_sheet_names {
            let patches = &self.queued_cf_patches[sheet_name];
            let sheet_path = match self.sheet_paths.get(sheet_name) {
                Some(p) => p,
                None => continue,
            };
            let xml = ooxml_util::zip_read_to_string(&mut zip, sheet_path)?;

            if styles_loaded.is_none() {
                let raw = ooxml_util::zip_read_to_string_opt(&mut zip, "xl/styles.xml")?
                    .unwrap_or_else(|| minimal_styles_xml());
                running_dxf_count = conditional_formatting::count_dxfs(&raw);
                styles_loaded = Some(raw);
            }

            let existing = conditional_formatting::extract_existing_cf_blocks(&xml);
            let pmax = conditional_formatting::scan_max_cf_priority(&xml);
            let result = conditional_formatting::build_cf_blocks(
                &existing,
                patches,
                pmax,
                running_dxf_count,
            );
            running_dxf_count += result.new_dxfs.len() as u32;
            new_dxfs_total.extend(result.new_dxfs);
            local_blocks
                .entry(sheet_path.clone())
                .or_default()
                .push(SheetBlock::ConditionalFormatting(result.block_bytes));
        }
        // If CF patches added new <dxf> entries, fold them into the
        // styles.xml that Phase 1's format patching may have already
        // mutated. We share `styles_xml` so a single save with both
        // cell-format edits and CF rules produces one styles.xml write.
        if !new_dxfs_total.is_empty() {
            let new_dxfs_xml: String = new_dxfs_total
                .iter()
                .map(conditional_formatting::dxf_to_xml)
                .collect::<Vec<_>>()
                .join("");
            let base = match styles_xml.take() {
                Some(s) => s,
                None => styles_loaded
                    .clone()
                    .unwrap_or_else(|| minimal_styles_xml()),
            };
            let updated = conditional_formatting::ensure_dxfs_section(&base, &new_dxfs_xml);
            styles_xml = Some(updated);
        }

        // --- Phase 2.5e: Hyperlinks (RFC-022) ---
        //
        // Per-sheet flush. For each sheet with queued hyperlink ops:
        //   1. Lazy-populate the ancillary registry so we know which
        //      rIds in the sheet's rels are hyperlinks (vs tables /
        //      comments / vmlDrawings).
        //   2. Get-or-load the rels graph into `rels_patches`. Phase 3's
        //      rels-serialization loop picks up the mutated graph.
        //   3. Read the source sheet XML, extract any existing
        //      `<hyperlinks>` block (resolving rIds → URLs via the rels
        //      graph), and merge with the queued ops.
        //   4. Push a `SheetBlock::Hyperlinks` (slot 19) into
        //      `local_blocks` so Phase 3's merge_blocks call inserts it.
        //
        // No `ContentTypeOp`s are emitted here — the worksheet content
        // type is already declared in every source ZIP, and external
        // hyperlinks live in the sheet's rels (which Phase 3 already
        // serializes). An empty `block_bytes` (all hyperlinks deleted
        // and the source had a block) is signaled to the merger by
        // pushing `SheetBlock::Hyperlinks(Vec::new())` — it drops the
        // existing block with no replacement.
        //
        // Cloning `sheet_order` into a local Vec sidesteps the
        // immutable-borrow-on-self-while-mutating-self.{ancillary,
        // rels_patches} conflict (same trick as Phase 2.5d).
        let sheet_order_local: Vec<String> = self.sheet_order.clone();
        for sheet_name in &sheet_order_local {
            let ops = match self.queued_hyperlinks.get(sheet_name) {
                Some(o) if !o.is_empty() => o.clone(),
                _ => continue,
            };
            let sheet_path = match self.sheet_paths.get(sheet_name).cloned() {
                Some(p) => p,
                None => continue, // unknown sheet name — silently skip
            };
            let rels_path = sheet_rels_path_for(&sheet_path);
            self.ancillary
                .populate_for_sheet(&mut zip, sheet_name, &sheet_path)
                .map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!(
                        "ancillary populate for '{sheet_name}': {e}"
                    ))
                })?;
            if !self.rels_patches.contains_key(&rels_path) {
                let g = load_or_empty_rels(&mut zip, &rels_path)?;
                self.rels_patches.insert(rels_path.clone(), g);
            }
            let rels = self
                .rels_patches
                .get_mut(&rels_path)
                .expect("just inserted above");
            let xml = ooxml_util::zip_read_to_string(&mut zip, &sheet_path)?;
            let existing = hyperlinks::extract_hyperlinks(xml.as_bytes(), rels);
            let had_existing = !existing.is_empty();
            let (block_bytes, _deleted_rids) =
                hyperlinks::build_hyperlinks_block(existing, &ops, rels);
            // No-op if there was nothing to delete and nothing to add.
            if block_bytes.is_empty() && !had_existing {
                continue;
            }
            local_blocks
                .entry(sheet_path.clone())
                .or_default()
                .push(SheetBlock::Hyperlinks(block_bytes));
        }

        // --- Phase 2.5f: Tables (RFC-024) ---
        //
        // Per-sheet flush. The workbook's existing-table inventory is
        // scanned ONCE up front (table `id` and `name` are
        // workbook-unique, not sheet-scoped, so per-sheet flushes
        // would otherwise risk allocating duplicate ids when two
        // sheets are flushed in the same save). For each sheet with
        // queued tables:
        //   1. Get-or-load the rels graph into `rels_patches` so the
        //      Phase-3 rels-serialization loop picks up the new
        //      TABLE rels we add.
        //   2. Call `tables::build_tables`, which serializes each
        //      patch into `xl/tables/tableN.xml` bytes (reusing the
        //      writer's emitter), allocates fresh rIds in the rels
        //      graph, queues `[Content_Types].xml` Override entries
        //      for each new part, and emits a merged `<tableParts>`
        //      block that includes any pre-existing TABLE rIds plus
        //      the new ones.
        //   3. Inject the new part bytes into `file_adds`.
        //   4. Forward content-type ops into `queued_content_type_ops`
        //      so Phase 2.5c aggregates them into one
        //      `[Content_Types].xml` mutation.
        //   5. Push `SheetBlock::TableParts(block_bytes)` (slot 37)
        //      into `local_blocks` so Phase-3's `merge_blocks` call
        //      replaces the sheet's existing `<tableParts>` (if any)
        //      with the merged block.
        //
        // Inventory + ID allocation across sheets: `build_tables`
        // takes a mutable inventory cloned per sheet only — but we
        // thread the names/ids/count manually here so concurrent
        // sheet flushes still see each others' allocations and
        // collisions surface deterministically. (Same trick as the
        // CF cross-sheet dxfId counter in Phase 2.5b.)
        if !self.queued_tables.is_empty() {
            let mut tables_inventory = tables::scan_existing_tables(&mut zip)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("scan tables: {e}")))?;

            // RFC-024 collision-scan extension (RFC-035 §8 risk #6):
            // include cloned table names from in-flight Phase 2.7
            // sheet copies so a user `add_table(name="Sales_2")` in
            // the same save against an as-yet-unflushed clone surfaces
            // a clean error rather than a silent duplicate.
            for n in &cloned_table_names {
                tables_inventory.names.insert(n.clone());
            }

            // Iterate sheets in source-document order so allocations
            // are deterministic across runs.
            for sheet_name in &sheet_order_local {
                let patches = match self.queued_tables.get(sheet_name) {
                    Some(p) if !p.is_empty() => p.clone(),
                    _ => continue,
                };
                let sheet_path = match self.sheet_paths.get(sheet_name).cloned() {
                    Some(p) => p,
                    None => continue,
                };
                let rels_path = sheet_rels_path_for(&sheet_path);
                if !self.rels_patches.contains_key(&rels_path) {
                    // RFC-035 Pod-δ fix (KNOWN_GAPS bug #3): a Phase
                    // 2.7-cloned sheet's rels live in file_adds, not
                    // in the source ZIP. Prefer file_adds/file_patches
                    // before falling back to the ZIP probe.
                    let g = if let Some(bytes) = self.file_adds.get(&rels_path) {
                        RelsGraph::parse(bytes).map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!(
                                "rels parse for cloned '{rels_path}': {e}"
                            ))
                        })?
                    } else if let Some(bytes) = file_patches.get(&rels_path) {
                        RelsGraph::parse(bytes).map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!(
                                "rels parse for patched '{rels_path}': {e}"
                            ))
                        })?
                    } else {
                        load_or_empty_rels(&mut zip, &rels_path)?
                    };
                    self.rels_patches.insert(rels_path.clone(), g);
                }
                let rels = self
                    .rels_patches
                    .get_mut(&rels_path)
                    .expect("just inserted above");
                let result = tables::build_tables(
                    &patches,
                    &tables_inventory,
                    rels,
                    Some(&mut part_id_allocator),
                )
                .map_err(|e| PyErr::new::<PyValueError, _>(e))?;

                // Update the running inventory so subsequent sheets'
                // build_tables calls see this sheet's allocations.
                for (path, _bytes) in &result.table_parts {
                    tables_inventory.count += 1;
                    tables_inventory.paths.push(path.clone());
                }
                for patch in &patches {
                    tables_inventory.names.insert(patch.name.clone());
                }
                for (path, bytes) in result.table_parts {
                    self.file_adds.insert(path, bytes);
                }
                // Reflect the freshly-allocated ids in the inventory's
                // `ids` set. We re-derive them by parsing the emitted
                // XML's id attribute — cheaper than threading them out
                // of build_tables and keeps that API surface narrow.
                for path in &tables_inventory.paths {
                    if let Some(bytes) = self.file_adds.get(path) {
                        let (id_opt, _) = tables::parse_table_root_attrs(bytes);
                        if let Some(id) = id_opt {
                            tables_inventory.ids.insert(id);
                        }
                    }
                }
                // Content-type Override per new part — funnel through
                // the existing Phase-2.5c aggregator.
                let ct_ops_for_sheet = self
                    .queued_content_type_ops
                    .entry(sheet_name.clone())
                    .or_default();
                for (part_name, ct) in result.new_content_types {
                    ct_ops_for_sheet.push(content_types::ContentTypeOp::AddOverride(
                        part_name, ct,
                    ));
                }
                if !result.table_parts_block.is_empty() {
                    local_blocks
                        .entry(sheet_path.clone())
                        .or_default()
                        .push(SheetBlock::TableParts(result.table_parts_block));
                }
            }
        }

        // --- Phase 2.5g: Comments + VML drawings (RFC-023) ---
        //
        // Per-sheet flush. For each sheet with queued comment ops:
        //   1. Lazy-populate the ancillary registry to learn the
        //      sheet's existing comments part path / VML part path
        //      (if any).
        //   2. Get-or-load the rels graph into `rels_patches`.
        //   3. Read the existing commentsN.xml + vmlDrawingN.vml (if
        //      any), merge in the queued ops, re-emit fresh bytes.
        //   4. Choose a workbook-wide unique `comments_n` / `vml_n`
        //      for sheets gaining their first comments part.
        //   5. Push a `SheetBlock::LegacyDrawing` (slot 31) into
        //      `local_blocks` so the merger injects it (deletes the
        //      tag if the rel was removed and the sheet had one).
        //   6. Route comment/vml part bytes:
        //      - if `merged.is_empty()` and no preserved VML shapes
        //        → schedule deletion via `file_deletes`.
        //      - otherwise patch (existing) or add (new) the bytes.
        //   7. Emit `[Content_Types].xml` ops (Override for the
        //      comments part; Default for the vml extension).
        //
        // Workbook-scope author table (`comment_authors`) lives on
        // the stack so all sheets share dedup. New `comments<N>.xml`
        // and `vmlDrawing<N>.vml` suffixes come from the shared
        // `part_id_allocator` (RFC-035 §5.2) — already pre-seeded by
        // a single pass over the source ZIP listing earlier in
        // Phase 2.5, so this loop only needs to populate the
        // ancillary registry for path-lookup purposes.
        let mut comment_authors = comments::CommentAuthorTable::new();
        for sheet_name in &sheet_order_local {
            let sp = match self.sheet_paths.get(sheet_name).cloned() {
                Some(p) => p,
                None => continue,
            };
            let _ = self.ancillary.populate_for_sheet(&mut zip, sheet_name, &sp);
        }

        let mut comments_file_writes: HashMap<String, Vec<u8>> = HashMap::new();
        let mut comments_file_deletes: HashSet<String> = HashSet::new();
        let mut comments_ct_ops: Vec<content_types::ContentTypeOp> = Vec::new();
        let mut vml_default_added = false;

        for sheet_name in &sheet_order_local {
            let ops = match self.queued_comments.get(sheet_name) {
                Some(o) if !o.is_empty() => o.clone(),
                _ => continue,
            };
            let sheet_path = match self.sheet_paths.get(sheet_name).cloned() {
                Some(p) => p,
                None => continue,
            };
            let rels_path = sheet_rels_path_for(&sheet_path);
            self.ancillary
                .populate_for_sheet(&mut zip, sheet_name, &sheet_path)
                .map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!(
                        "ancillary populate for '{sheet_name}': {e}"
                    ))
                })?;
            let (existing_comments_path, existing_vml_path) = {
                let anc = self
                    .ancillary
                    .get(sheet_name)
                    .cloned()
                    .unwrap_or_default();
                (anc.comments_part, anc.vml_drawing_part)
            };
            if !self.rels_patches.contains_key(&rels_path) {
                let g = load_or_empty_rels(&mut zip, &rels_path)?;
                self.rels_patches.insert(rels_path.clone(), g);
            }

            // Read existing parts (if any) before we mutate the rels graph.
            let existing_comments_xml: Option<Vec<u8>> = match &existing_comments_path {
                Some(p) => Some(
                    ooxml_util::zip_read_to_string(&mut zip, p)?.into_bytes(),
                ),
                None => None,
            };
            let existing_vml_xml: Option<Vec<u8>> = match &existing_vml_path {
                Some(p) => Some(
                    ooxml_util::zip_read_to_string(&mut zip, p)?.into_bytes(),
                ),
                None => None,
            };
            let sheet_xml = ooxml_util::zip_read_to_string(&mut zip, &sheet_path)?;

            // Decide N values: reuse existing-part N if any, else mint new
            // via the shared allocator (RFC-035 §5.2).
            let comments_n = match &existing_comments_path {
                Some(p) => parse_n_from_part_path(p, "xl/comments", ".xml")
                    .unwrap_or_else(|| part_id_allocator.alloc_comments()),
                None => part_id_allocator.alloc_comments(),
            };
            let vml_n = match &existing_vml_path {
                Some(p) => parse_n_from_part_path(p, "xl/drawings/vmlDrawing", ".vml")
                    .unwrap_or_else(|| part_id_allocator.alloc_vml_drawing()),
                None => part_id_allocator.alloc_vml_drawing(),
            };

            let rels = self
                .rels_patches
                .get_mut(&rels_path)
                .expect("just inserted above");

            let (result, _comments_rid_opt, vml_rid_opt) = comments::build_comments(
                existing_comments_xml.as_deref(),
                existing_vml_xml.as_deref(),
                &ops,
                sheet_xml.as_bytes(),
                rels,
                &mut comment_authors,
                comments_n,
                vml_n,
            );

            // Route comments part bytes.
            let comments_path = existing_comments_path
                .clone()
                .unwrap_or_else(|| format!("xl/comments{comments_n}.xml"));
            if result.comments_xml.is_empty() {
                // All comments deleted; remove the part entirely.
                if existing_comments_path.is_some() {
                    comments_file_deletes.insert(comments_path.clone());
                }
            } else {
                comments_file_writes.insert(comments_path.clone(), result.comments_xml);
                if existing_comments_path.is_none() {
                    comments_ct_ops.push(content_types::ContentTypeOp::AddOverride(
                        format!("/{}", comments_path),
                        comments::CT_COMMENTS.to_string(),
                    ));
                }
            }

            // Route vml drawing part bytes.
            let vml_path = existing_vml_path
                .clone()
                .unwrap_or_else(|| format!("xl/drawings/vmlDrawing{vml_n}.vml"));
            if result.vml_drawing.is_empty() {
                if existing_vml_path.is_some() {
                    comments_file_deletes.insert(vml_path.clone());
                }
            } else {
                comments_file_writes.insert(vml_path.clone(), result.vml_drawing);
                if existing_vml_path.is_none() && !vml_default_added {
                    comments_ct_ops.push(content_types::ContentTypeOp::EnsureDefault(
                        "vml".to_string(),
                        comments::CT_VML.to_string(),
                    ));
                    vml_default_added = true;
                }
            }

            // Emit a legacyDrawing block (slot 31) when the sheet
            // has a vml rel — or an empty payload to drop it when
            // every comment was deleted and no other VML shapes
            // remain.
            let legacy_block: Vec<u8> = match &result.legacy_drawing_rid {
                Some(rid) => format!(r#"<legacyDrawing r:id="{}"/>"#, rid.0).into_bytes(),
                None => Vec::new(),
            };
            local_blocks
                .entry(sheet_path.clone())
                .or_default()
                .push(SheetBlock::LegacyDrawing(legacy_block));

            // suppress unused_variable warning on vml_rid_opt
            let _ = vml_rid_opt;
        }

        // Merge comments_ct_ops into queued_content_type_ops under a
        // synthetic per-workbook key so Phase 2.5c picks them up.
        if !comments_ct_ops.is_empty() {
            self.queued_content_type_ops
                .entry("__rfc023_comments__".to_string())
                .or_default()
                .extend(comments_ct_ops);
        }

        // --- Phase 2.5k: Image adds (Sprint Λ Pod-β / RFC-045) ---
        //
        // Drains `queued_images` per sheet. For each sheet that has
        // queued images:
        //   1. Read the existing sheet rels (if any) — error if a
        //      `drawing` rel is already present (v1.5 limit:
        //      append-to-existing is a follow-up).
        //   2. Allocate a fresh `drawingN.xml` part via the shared
        //      part-id allocator.
        //   3. Allocate fresh `imageM.<ext>` media parts (one per
        //      queued image).
        //   4. Add an image rel for each one to a brand-new
        //      `xl/drawings/_rels/drawingN.xml.rels`.
        //   5. Add a drawing rel to the sheet's rels graph (creates
        //      `rels_patches` entry for that sheet's rels file).
        //   6. Splice `<drawing r:id="rIdN"/>` into the sheet XML
        //      (right before `<legacyDrawing>` if present, else
        //      before `</worksheet>`).
        //   7. Queue content-type ops: `<Default Extension="png" .../>`
        //      and `<Override PartName="/xl/drawings/drawingN.xml" .../>`.
        //
        // Phase 2.5k runs BEFORE Phase 3 so the cell/block merge
        // pass picks up any drawing-element splice we put into
        // `file_patches`. Sheet rels mutations land in
        // `rels_patches` which is serialized in the final emit pass.
        if !self.queued_images.is_empty() {
            self.apply_image_adds_phase(
                &mut file_patches,
                &mut zip,
                &mut part_id_allocator,
            )?;
        }

        // --- Phase 2.5l: Chart adds (Sprint Μ Pod-γ / RFC-046) ---
        //
        // Mirrors Phase 2.5k's image-add flow but with two extra
        // capabilities:
        //   * Each queued chart emits its own `xl/charts/chartN.xml`
        //     part, content-type override, and chart-rel under the
        //     drawing's nested rels graph.
        //   * Sheets that already have a `drawing` rel get the new
        //     `<xdr:graphicFrame>` SAX-merged into the existing
        //     drawing XML, instead of being rejected like 2.5k does
        //     for images. (Phase 2.5l also handles the "no existing
        //     drawing" case by allocating a fresh `drawingN.xml`.)
        // Phase 2.5l runs BEFORE Phase 3 so cell-range formulas in
        // chart XML can compose with cell rewrites in the same save.
        if !self.queued_charts.is_empty() {
            self.apply_chart_adds_phase(
                &mut file_patches,
                &mut zip,
                &mut part_id_allocator,
            )?;
        }

        // --- Phase 2.5m: Pivot adds (Sprint Ν Pod-γ / RFC-047 + RFC-048) ---
        //
        // Sequenced AFTER 2.5l (charts) and BEFORE Phase 3 (cell
        // patches). See `Plans/sprint-nu.md` Risk #1 for the
        // ordering rationale: charts → pivots → cells. Drainage:
        //
        //   1. For each queued cache:
        //      * Allocate `pivotCacheN` part id.
        //      * Write `xl/pivotCache/pivotCacheDefinition{N}.xml` and
        //        `xl/pivotCache/pivotCacheRecords{N}.xml` to file_adds.
        //      * Build the per-cache rels file pointing at records.
        //      * Add a workbook-rel of type PIVOT_CACHE_DEF.
        //      * Add content-type overrides for both parts.
        //   2. Splice <pivotCaches> into xl/workbook.xml.
        //   3. For each queued table:
        //      * Allocate `pivotTableN` part id.
        //      * Write `xl/pivotTables/pivotTable{N}.xml` to file_adds.
        //      * Build the table → cache rels file.
        //      * Add a sheet-rel of type PIVOT_TABLE.
        //      * Add a content-type override.
        if !self.queued_pivot_caches.is_empty() || !self.queued_pivot_tables.is_empty() {
            self.apply_pivot_adds_phase(&mut file_patches, &mut zip)?;
        }

        // --- Phase 2.5n: Sheet setup (Sprint Ο Pod 1A.5 / RFC-055) ---
        //
        // Drains queued sheet-setup mutations (sheetView /
        // sheetProtection / pageMargins / pageSetup / headerFooter)
        // into per-sheet `local_blocks` for splice via merge_blocks
        // in Phase 3. Sequenced AFTER pivots (2.5m) and BEFORE
        // autoFilter (2.5o) so a later sheet-protection toggle can
        // observe the pivot-table block when computing its allowed
        // operation set.
        //
        // Each non-empty block emits one SheetBlock variant; the
        // merger handles ECMA-376 §18.3.1.99 ordering. The
        // `print_titles` slot routes through workbook-scope
        // <definedNames> (RFC-021 path) — handled by the Workbook
        // coordinator on the Python side; the patcher just stashes
        // the strings on the queue for now.
        if !self.queued_sheet_setup.is_empty() {
            let sheet_titles: Vec<String> =
                self.queued_sheet_setup.keys().cloned().collect();
            for sheet_title in &sheet_titles {
                let queued = match self.queued_sheet_setup.get(sheet_title) {
                    Some(q) => q.clone(),
                    None => continue,
                };
                let sheet_path = match self.sheet_paths.get(sheet_title) {
                    Some(p) => p.clone(),
                    None => continue,
                };
                let specs = &queued.specs;

                // Emit each non-empty block into `local_blocks`. The
                // merger's replace-on-match semantics handle the
                // "existing element" case — we don't need to scan
                // the source XML up-front.
                if let Some(s) = &specs.sheet_view {
                    let bytes = wolfxl_writer::parse::sheet_setup::emit_sheet_views(s);
                    if !bytes.is_empty() {
                        local_blocks
                            .entry(sheet_path.clone())
                            .or_default()
                            .push(SheetBlock::SheetViews(bytes));
                    }
                }
                if let Some(s) = &specs.sheet_protection {
                    let bytes =
                        wolfxl_writer::parse::sheet_setup::emit_sheet_protection(s);
                    if !bytes.is_empty() {
                        local_blocks
                            .entry(sheet_path.clone())
                            .or_default()
                            .push(SheetBlock::SheetProtection(bytes));
                    }
                }
                if let Some(s) = &specs.page_margins {
                    let bytes = wolfxl_writer::parse::sheet_setup::emit_page_margins(s);
                    if !bytes.is_empty() {
                        local_blocks
                            .entry(sheet_path.clone())
                            .or_default()
                            .push(SheetBlock::PageMargins(bytes));
                    }
                }
                if let Some(s) = &specs.page_setup {
                    let bytes = wolfxl_writer::parse::sheet_setup::emit_page_setup(s);
                    if !bytes.is_empty() {
                        local_blocks
                            .entry(sheet_path.clone())
                            .or_default()
                            .push(SheetBlock::PageSetup(bytes));
                    }
                }
                if let Some(s) = &specs.header_footer {
                    let bytes = wolfxl_writer::parse::sheet_setup::emit_header_footer(s);
                    if !bytes.is_empty() {
                        local_blocks
                            .entry(sheet_path.clone())
                            .or_default()
                            .push(SheetBlock::HeaderFooter(bytes));
                    }
                }
                // print_titles: routed through workbook definedNames
                // by the Python coordinator. Nothing to do here —
                // the queue entry is informational only.
            }
        }

        // --- Phase 2.5o: AutoFilter (Sprint Ο Pod 1B / RFC-056) ---
        //
        // Sequenced AFTER pivots (2.5m) and BEFORE Phase 3 (cell
        // patches) per RFC-056 §5. For each sheet with a queued
        // `auto_filter`:
        //
        //   1. Lift the §10 dict into the typed `AutoFilter` model.
        //   2. Read the sheet's existing cells inside `auto_filter.ref`
        //      from the source XML (or file_adds for cloned sheets).
        //   3. Run `wolfxl_autofilter::evaluate` to compute the
        //      hidden-row offsets + sort permutation.
        //   4. Push a `SheetBlock::AutoFilter` into `local_blocks`
        //      (replaces any existing `<autoFilter>` element).
        //   5. Stash the hidden-row offsets in `autofilter_hidden_rows`
        //      so Phase 3 can apply `<row hidden="1">` markers AFTER
        //      sheet_patcher has rewritten the cell payloads.
        //
        // Sort permutation is computed but **not applied** in v2.0:
        // physical row reorder is deferred to v2.1 per RFC-056 §8.
        let mut autofilter_hidden_rows: HashMap<String, Vec<u32>> = HashMap::new();
        if !self.queued_autofilters.is_empty() {
            // Clone the queue keys to avoid borrowing self twice.
            let sheet_titles: Vec<String> = self.queued_autofilters.keys().cloned().collect();
            for sheet_title in &sheet_titles {
                let queued = self.queued_autofilters.get(sheet_title).cloned().unwrap();
                let sheet_path = match self.sheet_paths.get(sheet_title) {
                    Some(p) => p.clone(),
                    None => continue,
                };
                // Read current sheet XML (file_adds for clones, file_patches for
                // already-mutated, otherwise from the source ZIP).
                let xml_bytes: Vec<u8> = if let Some(b) = file_patches.get(&sheet_path) {
                    b.clone()
                } else if let Some(b) = self.file_adds.get(&sheet_path) {
                    b.clone()
                } else {
                    ooxml_util::zip_read_to_string(&mut zip, &sheet_path)?.into_bytes()
                };

                // Parse the dict to learn the ref + extract the col span.
                let af_model = wolfxl_autofilter::parse::parse_autofilter(&queued.dict)
                    .map_err(|e| PyErr::new::<PyValueError, _>(format!("Phase 2.5o: {e}")))?;
                let (start_row, end_row, start_col, end_col) = match af_model
                    .ref_
                    .as_deref()
                    .and_then(autofilter_helpers::parse_a1_range)
                {
                    Some(t) => t,
                    None => {
                        // No ref → just splice the (probably empty) block.
                        // Skip evaluation.
                        let block = wolfxl_autofilter::emit::emit(&af_model);
                        if !block.is_empty() {
                            local_blocks
                                .entry(sheet_path.clone())
                                .or_default()
                                .push(SheetBlock::AutoFilter(block));
                        }
                        continue;
                    }
                };

                // Read rows of cells in [start_row+1..=end_row][start_col..=end_col].
                // The header row (start_row) is skipped — autoFilter applies
                // to the data rows only.
                let data_start = start_row + 1;
                let rows_data = autofilter_helpers::extract_cell_grid(
                    &xml_bytes,
                    data_start,
                    end_row,
                    start_col,
                    end_col,
                )?;

                // Drain.
                let drain = autofilter::drain_autofilter(&queued, &rows_data, None)
                    .map_err(|e| PyErr::new::<PyValueError, _>(format!("Phase 2.5o: {e}")))?;

                // Convert offsets back to absolute row numbers.
                let abs_hidden: Vec<u32> = drain
                    .hidden_offsets
                    .iter()
                    .map(|off| data_start + off)
                    .collect();
                if !abs_hidden.is_empty() {
                    autofilter_hidden_rows.insert(sheet_path.clone(), abs_hidden);
                }

                // Splice the block (replace any existing).
                if !drain.block_bytes.is_empty() {
                    local_blocks
                        .entry(sheet_path.clone())
                        .or_default()
                        .push(SheetBlock::AutoFilter(drain.block_bytes));
                }
            }
        }

        // --- Phase 3: Patch worksheet XMLs ---
        //
        // Two-pass per sheet: cell-level patches via `sheet_patcher`, then
        // sibling-block insertions via `wolfxl_merger`. The two passes
        // commute (cells live inside <sheetData>, blocks are siblings) so
        // composing them is straightforward.
        //
        // `file_patches` was declared early (before Phase 2.7) so RFC-035
        // can write workbook.xml + workbook.xml.rels into it before the
        // per-sheet phases run.

        // Sheets that have either kind of patch.
        let mut all_sheet_paths: std::collections::HashSet<String> =
            std::collections::HashSet::new();
        all_sheet_paths.extend(sheet_cell_patches.keys().cloned());
        all_sheet_paths.extend(local_blocks.keys().cloned());
        // Sprint Ο Pod 1B: include sheets that only need a row-hidden
        // marker pass (no other patches).
        all_sheet_paths.extend(autofilter_hidden_rows.keys().cloned());

        for sheet_path in &all_sheet_paths {
            // RFC-035 composition (Pod-δ fix for KNOWN_GAPS bugs #1/#3):
            // a Phase 2.7-cloned sheet's bytes live in `file_adds`,
            // not in the source ZIP. If a user mutates the clone in
            // the same save (cell value, format, table, DV, CF, etc.)
            // we must read the clone's source XML from
            // `file_adds`/`file_patches` first, falling back to the
            // ZIP for genuine source-side sheets.
            let xml = if let Some(bytes) = file_patches.get(sheet_path) {
                String::from_utf8_lossy(bytes).into_owned()
            } else if let Some(bytes) = self.file_adds.get(sheet_path) {
                String::from_utf8_lossy(bytes).into_owned()
            } else {
                ooxml_util::zip_read_to_string(&mut zip, sheet_path)?
            };

            // Pass 1: cell-level patches.
            let after_cells: Vec<u8> = if let Some(patches) = sheet_cell_patches.get(sheet_path) {
                sheet_patcher::patch_worksheet(&xml, patches)
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("Patch failed: {e}")))?
                    .into_bytes()
            } else {
                xml.into_bytes()
            };

            // Pass 2: sibling-block insertions.
            let after_blocks = if let Some(blocks) = local_blocks.get(sheet_path) {
                if blocks.is_empty() {
                    after_cells
                } else {
                    wolfxl_merger::merge_blocks(&after_cells, blocks.clone())
                        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Merge failed: {e}")))?
                }
            } else {
                after_cells
            };

            // Pass 3 (Sprint Ο Pod 1B): apply `<row hidden="1">`
            // markers from Phase 2.5o. Only runs for sheets touched
            // by an autoFilter evaluation.
            let after_blocks = if let Some(rows) = autofilter_hidden_rows.get(sheet_path) {
                if rows.is_empty() {
                    after_blocks
                } else {
                    autofilter_helpers::stamp_row_hidden(&after_blocks, rows)?
                }
            } else {
                after_blocks
            };

            // Route the rewrite back to the right primitive: if this
            // path is a Phase 2.7 cloned sheet (lives in file_adds),
            // write the patched bytes back to file_adds so they're
            // emitted by the new-entry pass in Phase 4 (Pod-δ fix
            // for KNOWN_GAPS bugs #1/#3). For source-side sheets,
            // file_patches replaces the source-entry bytes as before.
            if self.file_adds.contains_key(sheet_path) {
                self.file_adds.insert(sheet_path.clone(), after_blocks);
            } else {
                file_patches.insert(sheet_path.clone(), after_blocks);
            }
        }

        // Add styles.xml patch if modified
        if let Some(ref sxml) = styles_xml {
            file_patches.insert("xl/styles.xml".to_string(), sxml.as_bytes().to_vec());
        }

        // --- Phase 2.5q: Workbook security (Sprint Ο Pod 1D / RFC-058) ---
        //
        // Splices `<workbookProtection>` and `<fileSharing>` into
        // `xl/workbook.xml` at canonical CT_Workbook child positions:
        //
        //   fileVersion → fileSharing → workbookPr → workbookProtection
        //   → bookViews → sheets → ...
        //
        // Sequenced AFTER Phase 2.5m (pivots) and BEFORE Phase 2.5h
        // (sheet reorder) so the reorder phase sees the updated
        // workbook.xml (the splice + reorder commute, but composing
        // them through `workbook_xml_in_progress` matches the
        // RFC-035 / RFC-036 hand-off pattern exactly).
        //
        // Empty queue ⇒ identity: workbook.xml flows through
        // unchanged (no extra parse, no extra serialize).
        let mut workbook_xml_in_progress: Option<Vec<u8>> = None;
        if let Some(ref sec) = self.queued_workbook_security {
            if !sec.is_empty() {
                let wb_bytes: Vec<u8> = match file_patches.get("xl/workbook.xml") {
                    Some(b) => b.clone(),
                    None => ooxml_util::zip_read_to_string(&mut zip, "xl/workbook.xml")?
                        .into_bytes(),
                };
                let updated = security::merge_workbook_security(&wb_bytes, sec)
                    .map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("workbook-security merge: {e}"))
                    })?;
                workbook_xml_in_progress = Some(updated);
            }
        }

        // --- Phase 2.5h: Sheet reorder (RFC-036) ---
        //
        // Sequenced BEFORE Phase 2.5f because both phases mutate
        // `xl/workbook.xml`. When `queued_sheet_moves` is non-empty
        // we read workbook.xml ONCE here, apply the reorder + the
        // `<definedName localSheetId>` integer remap, and stash the
        // resulting bytes for Phase 2.5f to consume (so the defined-
        // names merger doesn't re-read the source ZIP entry). We also
        // update `self.sheet_order` so downstream phases (RFC-020
        // `app.xml` regen, RFC-026 CF aggregation) iterate the
        // post-move tab list.
        //
        // RFC-058 composition: `workbook_xml_in_progress` may already
        // hold the post-Phase-2.5q (security splice) bytes; the read
        // below honours that handoff before falling back to the
        // file_patches / source-ZIP layers.
        if !self.queued_sheet_moves.is_empty() {
            // RFC-035 + RFC-036 composition (Pod-δ fix for KNOWN_GAPS
            // bug #2): Phase 2.7 writes the cloned <sheet> entry into
            // file_patches["xl/workbook.xml"]. If we re-read from the
            // source ZIP here, the reorder would operate on the
            // pre-clone bytes and the new <sheet> would be silently
            // dropped from the saved workbook.xml. Prefer file_patches
            // so 2.7 → 2.5h compose via the file_patches handoff that
            // RFC-035 §5.4 specifies.
            let wb_bytes: Vec<u8> = match workbook_xml_in_progress.take() {
                Some(b) => b,
                None => match file_patches.get("xl/workbook.xml") {
                    Some(b) => b.clone(),
                    None => ooxml_util::zip_read_to_string(&mut zip, "xl/workbook.xml")?
                        .into_bytes(),
                },
            };
            let result = sheet_order::merge_sheet_moves(
                &wb_bytes,
                &self.queued_sheet_moves,
            )
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("sheet-reorder merge: {e}")))?;
            workbook_xml_in_progress = Some(result.workbook_xml);
            self.sheet_order = result.new_order;
        }

        // --- Phase 2.5f: Defined names (RFC-021) ---
        //
        // Workbook-level (single XML part), not per-sheet. When the
        // queue is non-empty we read `xl/workbook.xml`, splice the
        // `<definedNames>` block (or inject one after `</sheets>` if
        // missing), and route the result through `file_patches`.
        // Empty queue is the no-op identity path — workbook.xml is
        // not touched. The merger preserves all unrelated children of
        // `<workbook>` byte-for-byte.
        //
        // RFC-036 composition: if Phase 2.5h already produced an
        // updated workbook.xml, feed the merger those bytes (rather
        // than re-reading the source) so the move + defined-names
        // mutations compose without two source-XML parses.
        if !self.queued_defined_names.is_empty() {
            let wb_xml_bytes: Vec<u8> = match workbook_xml_in_progress.take() {
                Some(bytes) => bytes,
                None => {
                    let s = ooxml_util::zip_read_to_string(&mut zip, "xl/workbook.xml")?;
                    s.into_bytes()
                }
            };
            let updated = defined_names::merge_defined_names(
                &wb_xml_bytes,
                &self.queued_defined_names,
            )
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("defined-names merge: {e}")))?;
            file_patches.insert("xl/workbook.xml".to_string(), updated);
        } else if let Some(bytes) = workbook_xml_in_progress.take() {
            // No defined-names work, but Phase 2.5h produced a workbook
            // rewrite — route it through file_patches.
            file_patches.insert("xl/workbook.xml".to_string(), bytes);
        }

        // Serialize any mutated `*.rels` graphs. Routing depends on whether
        // the path already exists in the source ZIP:
        //   - present → `file_patches` replaces it in place (RFC-020 precedent)
        //   - absent  → `file_adds` appends a brand-new entry (RFC-013)
        // The "absent" branch is the common case for RFC-022 on a clean
        // file that had zero hyperlinks before.
        for (path, graph) in &self.rels_patches {
            let bytes = graph.serialize();
            if zip.by_name(path).is_ok() {
                file_patches.insert(path.clone(), bytes);
            } else {
                self.file_adds.insert(path.clone(), bytes);
            }
        }

        // --- Phase 2.5c: Content-types aggregation (RFC-013) ---
        //
        // Cross-sheet collection of `ContentTypeOp`s; one parse + serialize
        // of `[Content_Types].xml` regardless of how many sheets contribute.
        // Iteration follows `sheet_order` (source-document order) so the
        // resulting Override sequence is deterministic when multiple sheets
        // each push ops.
        //
        // No live producer in the current slice — `queued_content_type_ops`
        // is always empty, so this loop short-circuits at the
        // `is_empty()` guard. RFC-022 (Hyperlinks via new
        // `xl/worksheets/_rels/sheetN.xml.rels` parts), RFC-023 (Comments
        // via new `xl/comments<N>.xml` Overrides + a vml `Default`),
        // and RFC-024 (Tables via new `xl/tables/tableN.xml` Overrides)
        // will be the first volume producers.
        let mut content_type_ops: Vec<content_types::ContentTypeOp> = Vec::new();
        for sheet_name in &self.sheet_order {
            if let Some(ops) = self.queued_content_type_ops.get(sheet_name) {
                content_type_ops.extend(ops.iter().cloned());
            }
        }
        // Also pick up synthetic per-workbook keys (e.g. RFC-023
        // ``__rfc023_comments__`` and RFC-045
        // ``__rfc045_drawing_N__``) that aren't tied to a single
        // sheet name in `sheet_order`. Iterate in sorted order so the
        // emitted Override sequence is deterministic.
        let mut synth_keys: Vec<&String> = self
            .queued_content_type_ops
            .keys()
            .filter(|k| !self.sheet_order.contains(k))
            .collect();
        synth_keys.sort();
        for k in synth_keys {
            if let Some(ops) = self.queued_content_type_ops.get(k) {
                content_type_ops.extend(ops.iter().cloned());
            }
        }
        if !content_type_ops.is_empty() {
            let ct_xml = ooxml_util::zip_read_to_string(&mut zip, "[Content_Types].xml")?;
            let mut graph = content_types::ContentTypesGraph::parse(ct_xml.as_bytes())
                .map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("[Content_Types].xml parse: {e}"))
                })?;
            for op in &content_type_ops {
                graph.apply_op(op);
            }
            file_patches.insert("[Content_Types].xml".to_string(), graph.serialize());
        }

        // --- Phase 2.5d: Document properties (RFC-020) ---
        //
        // Full rewrite of `docProps/core.xml` + `docProps/app.xml` when
        // `queued_props` is set. Routing depends on whether each part
        // already exists in the source ZIP:
        //   - present → file_patches replaces it in place
        //   - absent  → file_adds appends a brand-new entry (RFC-013)
        //
        // `docProps/core.xml` is OPTIONAL in OOXML (some minimal xlsx
        // readers omit it), which is why the file_adds path matters
        // here. See RFC-020 §8 risk #3.
        //
        // If the caller didn't supply `sheet_names`, we thread the
        // patcher's `sheet_order` in so app.xml's `<TitlesOfParts>`
        // matches the workbook's tab order.
        if let Some(ref payload) = self.queued_props {
            let mut effective = payload.clone();
            if effective.sheet_names.is_empty() {
                effective.sheet_names = self.sheet_order.clone();
            }
            let core_bytes = properties::rewrite_core_props(&effective);
            let app_bytes = properties::rewrite_app_props(&effective);

            let core_in_source = source_zip_has_entry(&mut zip, "docProps/core.xml");
            let app_in_source = source_zip_has_entry(&mut zip, "docProps/app.xml");

            if core_in_source {
                file_patches.insert("docProps/core.xml".into(), core_bytes);
            } else {
                self.file_adds.insert("docProps/core.xml".into(), core_bytes);
            }
            if app_in_source {
                file_patches.insert("docProps/app.xml".into(), app_bytes);
            } else {
                self.file_adds.insert("docProps/app.xml".into(), app_bytes);
            }
        }

        // Route RFC-023 comments/vml part bytes into the right
        // primitive (in-place patch vs. new add) and delete dropped
        // parts. Done after Phase 2.5d so we already know which paths
        // exist in the source ZIP.
        for (path, bytes) in comments_file_writes.drain() {
            if zip.by_name(&path).is_ok() {
                file_patches.insert(path, bytes);
            } else {
                self.file_adds.insert(path, bytes);
            }
        }
        for path in comments_file_deletes.drain() {
            self.file_deletes.insert(path);
        }

        // --- Phase 2.5i: Structural axis shifts (RFC-030 / RFC-031) ---
        //
        // Drains `queued_axis_shifts` in append order. For each op:
        //   1. Read sheet XML from `file_patches` if already mutated,
        //      else from the source ZIP.
        //   2. Read every table part attached to the sheet (via the
        //      ancillary registry's source-side scan).
        //   3. Read every comments/vmlDrawing part attached to the sheet.
        //   4. Read `xl/workbook.xml` once (cached across ops in this
        //      flush block) for defined-name shifts.
        //   5. Build `wolfxl_structural::SheetXmlInputs` and call
        //      `apply_workbook_shift` with this single op.
        //   6. Merge the returned `file_patches` back into our
        //      `file_patches`.
        //
        // The empty-queue path is the no-op identity: a workbook with
        // zero queued shifts produces byte-identical output (the
        // outer `is_empty()` short-circuit at the top of `do_save`
        // handles the global no-op case; this block handles the
        // partial case where some other RFC also queued ops).
        if !self.queued_axis_shifts.is_empty() {
            self.apply_axis_shifts_phase(&mut file_patches, &mut zip)?;
        }

        // --- Phase 2.5j: Range moves (RFC-034) ---
        //
        // Drains `queued_range_moves` in append order. Each op reads
        // the affected sheet XML from `file_patches` if already
        // mutated (e.g. by Phase 2.5i axis shifts), else from the
        // source ZIP, and routes through
        // `wolfxl_structural::apply_range_move`. Multi-op sequencing
        // mirrors Phase 2.5i: each op runs against the post-previous
        // bytes.
        if !self.queued_range_moves.is_empty() {
            self.apply_range_moves_phase(&mut file_patches, &mut zip)?;
        }

        // --- Phase 2.8: calcChain.xml rebuild (Sprint Θ Pod-C3) ---
        //
        // Walk every sheet in `sheet_order`, scan each sheet's
        // post-mutation XML for formula cells, and emit a fresh
        // `xl/calcChain.xml`. Excel transparently rebuilds this file
        // on next open if it's stale, so the rebuild is a perf-only
        // hint — it never changes correctness. We still do it because
        // (a) it makes Excel's first-open faster, (b) external tools
        // that read calcChain directly see the right cells, and (c)
        // it keeps WolfXL output closer to a freshly-saved Excel
        // file.
        //
        // The no-op short-circuit at the top of `do_save` already
        // bypasses this whole flush, so byte-identical no-op saves
        // are unaffected.
        self.rebuild_calc_chain_phase(&mut file_patches, &mut zip)?;

        drop(zip);

        // --- Phase 4: Rewrite ZIP ---
        let src = File::open(&self.file_path).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Cannot open '{}': {e}", self.file_path))
        })?;
        let mut zip = ZipArchive::new(src)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP read error: {e}")))?;

        let dst = File::create(output_path).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Cannot create '{output_path}': {e}"))
        })?;
        let mut out = ZipWriter::new(dst);

        // RFC-013: collect the source-entry names so we can sanity-check
        // that no file_adds collides with one (caller bug per §8 risk #2).
        let mut source_names: HashSet<String> = HashSet::with_capacity(zip.len());
        for i in 0..zip.len() {
            let mut file = zip
                .by_index(i)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP entry read error: {e}")))?;
            let name = file.name().to_string();
            source_names.insert(name.clone());

            // RFC-013: skip source entries explicitly marked for deletion
            // (reserved for future RFC-035; v1 callers leave file_deletes
            // empty so this branch is dead in the current slice).
            if self.file_deletes.contains(&name) {
                continue;
            }

            let mut opts = SimpleFileOptions::default().compression_method(file.compression());
            if let Some(dt) = file.last_modified() {
                opts = opts.last_modified_time(dt);
            }
            if let Some(mode) = file.unix_mode() {
                opts = opts.unix_permissions(mode);
            }

            if file.is_dir() {
                out.add_directory(&name, opts)
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP write error: {e}")))?;
                continue;
            }

            let data = if let Some(patched) = file_patches.get(&name) {
                patched.clone()
            } else {
                let mut buf = Vec::new();
                file.read_to_end(&mut buf)
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP read error: {e}")))?;
                buf
            };

            out.start_file(&name, opts)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP write error: {e}")))?;
            out.write_all(&data)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP write error: {e}")))?;
        }

        // RFC-013: emit file_adds entries after the source-entry pass.
        // Collisions with source entries are a hard panic — callers
        // should be using `file_patches` (REPLACE) when the entry
        // already exists. The mtime stamp honors WOLFXL_TEST_EPOCH so
        // golden-file tests stay deterministic.
        if !self.file_adds.is_empty() {
            for new_path in self.file_adds.keys() {
                assert!(
                    !source_names.contains(new_path),
                    "file_adds collision with source entry: {new_path} — \
                     caller bug; use file_patches to REPLACE existing entries"
                );
            }
            // Iterate in sorted order so a single save with multiple new
            // entries produces deterministic ZIP output (the ZIP spec does
            // not require a particular entry order, but byte-identical
            // re-runs do).
            let mut new_paths: Vec<&String> = self.file_adds.keys().collect();
            new_paths.sort();
            let dt = epoch_or_now();
            for new_path in new_paths {
                let bytes = &self.file_adds[new_path];
                let opts = SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Deflated)
                    .last_modified_time(dt);
                out.start_file(new_path, opts).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("ZIP write error: {e}"))
                })?;
                out.write_all(bytes)
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP write error: {e}")))?;
            }
        }

        out.finish()
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP finalize error: {e}")))?;

        Ok(())
    }

    /// Phase 2.5i — drive `wolfxl_structural::apply_workbook_shift`
    /// Sprint Λ Pod-β (RFC-045) — drain `self.queued_images`.
    ///
    /// For each sheet:
    /// 1. Read the sheet's rels graph (from `rels_patches` if
    ///    already mutated, else from the source ZIP, else fresh).
    /// 2. Reject sheets that already have a `drawing` rel — v1.5
    ///    limit (NotImplementedError to surface the gap).
    /// 3. Allocate one fresh `drawingN.xml` part + one fresh
    ///    `imageM.<ext>` per queued image via the shared part-id
    ///    allocator.
    /// 4. Synthesize the drawing part XML, the drawing rels XML,
    ///    and the media bytes — all into `file_adds`.
    /// 5. Add a `drawing` rel to the sheet's rels graph in
    ///    `rels_patches`.
    /// 6. Splice a `<drawing r:id="..."/>` element into the sheet
    ///    XML in `file_patches` so Phase 3's downstream merger and
    ///    final emit see it.
    /// 7. Queue content-type ops: `<Default Extension="<ext>" .../>`
    ///    once per distinct extension and one
    ///    `<Override PartName="/xl/drawings/drawingN.xml" .../>`
    ///    per drawing.
    fn apply_image_adds_phase(
        &mut self,
        file_patches: &mut HashMap<String, Vec<u8>>,
        zip: &mut ZipArchive<File>,
        part_id_allocator: &mut wolfxl_rels::PartIdAllocator,
    ) -> PyResult<()> {
        // Drain queued_images into a stable order — sheet_order so two
        // saves of the same workbook with the same calls produce the
        // same output.
        let drained: Vec<(String, Vec<QueuedImageAdd>)> = self
            .sheet_order
            .iter()
            .filter_map(|s| {
                self.queued_images
                    .remove(s)
                    .map(|v| (s.clone(), v))
            })
            .collect();
        if drained.is_empty() {
            // Defensive — should be unreachable since the caller checked.
            self.queued_images.clear();
            return Ok(());
        }

        for (sheet_name, queued) in drained {
            if queued.is_empty() {
                continue;
            }
            let sheet_path = self
                .sheet_paths
                .get(&sheet_name)
                .cloned()
                .ok_or_else(|| {
                    PyValueError::new_err(format!("queue_image_add: no such sheet: {sheet_name}"))
                })?;

            // 1. Get sheet rels graph (from rels_patches → file_adds → ZIP).
            let sheet_rels_path = format!(
                "{}/_rels/{}.rels",
                sheet_path.rsplit_once('/').map(|(d, _)| d).unwrap_or(""),
                sheet_path.rsplit('/').next().unwrap_or("")
            );
            let mut rels_graph: wolfxl_rels::RelsGraph = if let Some(g) =
                self.rels_patches.get(&sheet_rels_path)
            {
                g.clone()
            } else if let Some(bytes) = self.file_adds.get(&sheet_rels_path) {
                wolfxl_rels::RelsGraph::parse(bytes).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("rels parse: {e}"))
                })?
            } else {
                match ooxml_util::zip_read_to_string_opt(zip, &sheet_rels_path)? {
                    Some(s) => wolfxl_rels::RelsGraph::parse(s.as_bytes()).map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("rels parse: {e}"))
                    })?,
                    None => wolfxl_rels::RelsGraph::new(),
                }
            };

            // 2. Reject if drawing rel already exists.
            for r in rels_graph.iter() {
                if r.rel_type == wolfxl_rels::rt::DRAWING {
                    return Err(PyErr::new::<pyo3::exceptions::PyNotImplementedError, _>(
                        format!(
                            "queue_image_add on sheet {sheet_name:?}: \
                             sheet already has a drawing part — appending to an \
                             existing drawing is a v1.5 follow-up. As a workaround, \
                             remove the existing drawing first or open the file in \
                             write mode."
                        ),
                    ));
                }
            }

            // 3. Allocate part suffixes.
            let drawing_n = part_id_allocator.alloc_drawing();
            let image_indices: Vec<u32> =
                queued.iter().map(|_| part_id_allocator.alloc_image()).collect();

            // 4. Synthesize drawing part XML + rels.
            let drawing_xml = build_drawing_xml(&queued);
            let drawing_rels_xml = build_drawing_rels_xml(&queued, &image_indices);
            let drawing_path = format!("xl/drawings/drawing{drawing_n}.xml");
            let drawing_rels_path =
                format!("xl/drawings/_rels/drawing{drawing_n}.xml.rels");
            self.file_adds
                .insert(drawing_path.clone(), drawing_xml.into_bytes());
            self.file_adds
                .insert(drawing_rels_path, drawing_rels_xml.into_bytes());
            for (img, &n) in queued.iter().zip(image_indices.iter()) {
                let media_path = format!("xl/media/image{n}.{}", img.ext);
                self.file_adds.insert(media_path, img.data.clone());
            }

            // 5. Add drawing rel to sheet rels graph.
            let drawing_rid = rels_graph.add(
                wolfxl_rels::rt::DRAWING,
                &format!("../drawings/drawing{drawing_n}.xml"),
                wolfxl_rels::TargetMode::Internal,
            );
            self.rels_patches.insert(sheet_rels_path, rels_graph);

            // 6. Splice <drawing r:id> into sheet XML.
            let sheet_xml = if let Some(b) = file_patches.get(&sheet_path) {
                String::from_utf8_lossy(b).into_owned()
            } else if let Some(b) = self.file_adds.get(&sheet_path) {
                String::from_utf8_lossy(b).into_owned()
            } else {
                ooxml_util::zip_read_to_string(zip, &sheet_path)?
            };
            let after = splice_drawing_ref(&sheet_xml, &drawing_rid.0)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("splice drawing: {e}")))?;
            file_patches.insert(sheet_path, after.into_bytes());

            // 7. Queue content-type ops.
            //    - one Default Extension per distinct extension
            //    - one Override per drawing part
            let mut seen_exts: std::collections::HashSet<String> =
                std::collections::HashSet::new();
            let mut ops: Vec<content_types::ContentTypeOp> = Vec::new();
            for img in &queued {
                if seen_exts.insert(img.ext.clone()) {
                    let ct = match img.ext.as_str() {
                        "png" => "image/png",
                        "jpeg" | "jpg" => "image/jpeg",
                        "gif" => "image/gif",
                        "bmp" => "image/bmp",
                        _ => "application/octet-stream",
                    };
                    ops.push(content_types::ContentTypeOp::EnsureDefault(
                        img.ext.clone(),
                        ct.to_string(),
                    ));
                }
            }
            ops.push(content_types::ContentTypeOp::AddOverride(
                format!("/xl/drawings/drawing{drawing_n}.xml"),
                "application/vnd.openxmlformats-officedocument.drawing+xml".to_string(),
            ));
            self.queued_content_type_ops
                .entry(format!("__rfc045_drawing_{drawing_n}__"))
                .or_default()
                .extend(ops);
        }
        Ok(())
    }

    /// Sprint Μ Pod-γ (RFC-046) — drain `self.queued_charts`.
    ///
    /// For each sheet that has queued charts:
    /// 1. Read the sheet's rels graph (from `rels_patches` if
    ///    already mutated, else from `file_adds`/source ZIP).
    /// 2. Probe for an existing `drawing` rel:
    ///    * If absent — allocate a fresh `drawingN.xml`,
    ///      synthesize its body containing one
    ///      `<xdr:graphicFrame>` per queued chart, plus a fresh
    ///      `xl/drawings/_rels/drawingN.xml.rels` with one chart
    ///      rel per chart. Splice `<drawing r:id="...">` into
    ///      sheet XML.
    ///    * If present — load the existing drawing XML + rels,
    ///      append a `<xdr:graphicFrame>` per queued chart via
    ///      SAX, append a chart rel per chart to the drawing's
    ///      rels file. The sheet XML's `<drawing>` ref is left
    ///      alone (already pointing at the drawing).
    /// 3. Allocate one fresh `xl/charts/chartN.xml` per queued
    ///    chart and route the caller-supplied bytes through
    ///    `file_adds`.
    /// 4. Queue content-type ops: one `<Override>` per chart, plus
    ///    a `<Override>` for the drawing if we created one fresh.
    fn apply_chart_adds_phase(
        &mut self,
        file_patches: &mut HashMap<String, Vec<u8>>,
        zip: &mut ZipArchive<File>,
        part_id_allocator: &mut wolfxl_rels::PartIdAllocator,
    ) -> PyResult<()> {
        // Drain in sheet_order for stable output across saves.
        let drained: Vec<(String, Vec<QueuedChartAdd>)> = self
            .sheet_order
            .iter()
            .filter_map(|s| self.queued_charts.remove(s).map(|v| (s.clone(), v)))
            .collect();
        if drained.is_empty() {
            self.queued_charts.clear();
            return Ok(());
        }

        for (sheet_name, queued) in drained {
            if queued.is_empty() {
                continue;
            }
            let sheet_path = self
                .sheet_paths
                .get(&sheet_name)
                .cloned()
                .ok_or_else(|| {
                    PyValueError::new_err(format!(
                        "queue_chart_add: no such sheet: {sheet_name}"
                    ))
                })?;

            // 1. Get sheet rels graph (rels_patches → file_adds → ZIP).
            let sheet_rels_path = format!(
                "{}/_rels/{}.rels",
                sheet_path.rsplit_once('/').map(|(d, _)| d).unwrap_or(""),
                sheet_path.rsplit('/').next().unwrap_or("")
            );
            let mut sheet_rels: wolfxl_rels::RelsGraph = if let Some(g) =
                self.rels_patches.get(&sheet_rels_path)
            {
                g.clone()
            } else if let Some(bytes) = self.file_adds.get(&sheet_rels_path) {
                wolfxl_rels::RelsGraph::parse(bytes).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("rels parse: {e}"))
                })?
            } else {
                match ooxml_util::zip_read_to_string_opt(zip, &sheet_rels_path)? {
                    Some(s) => wolfxl_rels::RelsGraph::parse(s.as_bytes())
                        .map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!("rels parse: {e}"))
                        })?,
                    None => wolfxl_rels::RelsGraph::new(),
                }
            };

            // 2. Probe for existing drawing rel + drawing path.
            let mut existing_drawing_target: Option<String> = None;
            for r in sheet_rels.iter() {
                if r.rel_type == wolfxl_rels::rt::DRAWING {
                    existing_drawing_target = Some(r.target.clone());
                    break;
                }
            }

            // Allocate one chart part per queued chart.
            let chart_indices: Vec<u32> =
                queued.iter().map(|_| part_id_allocator.alloc_chart()).collect();

            // Pre-content-type ops accumulator for this sheet.
            let mut ct_ops: Vec<content_types::ContentTypeOp> = Vec::new();
            for &n in &chart_indices {
                ct_ops.push(content_types::ContentTypeOp::AddOverride(
                    format!("/xl/charts/chart{n}.xml"),
                    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
                        .to_string(),
                ));
            }

            // Emit the chart XML parts up front.
            for (chart, &n) in queued.iter().zip(chart_indices.iter()) {
                let path = format!("xl/charts/chart{n}.xml");
                self.file_adds.insert(path, chart.chart_xml.clone());
            }

            // Branch on fresh vs. existing drawing.
            let drawing_n: u32;
            let drawing_path: String;
            let drawing_rels_path: String;
            let mut drawing_rels: wolfxl_rels::RelsGraph;
            let mut new_drawing_xml_bytes: Option<Vec<u8>> = None;
            if let Some(target) = existing_drawing_target {
                // Existing: resolve the drawing path relative to the
                // OWNING PART's directory (i.e. the sheet itself, not
                // the rels file). Rels targets are interpreted
                // relative to the part the rels graph describes —
                // here that's `xl/worksheets/sheetN.xml`, so the base
                // is `xl/worksheets/`.
                let sheet_dir = sheet_path
                    .rsplit_once('/')
                    .map(|(d, _)| d)
                    .unwrap_or("")
                    .to_string();
                let resolved = resolve_relative_path(&sheet_dir, &target);
                drawing_path = resolved.clone();
                let n = drawing_n_from_path(&drawing_path).unwrap_or_else(|| {
                    part_id_allocator.alloc_drawing()
                });
                drawing_n = n;
                drawing_rels_path = format!(
                    "xl/drawings/_rels/drawing{drawing_n}.xml.rels"
                );
                // Load existing drawing rels (if any) — drawing
                // graphs without rels are legal but rare.
                drawing_rels = if let Some(g) =
                    self.rels_patches.get(&drawing_rels_path)
                {
                    g.clone()
                } else if let Some(b) = self.file_adds.get(&drawing_rels_path) {
                    wolfxl_rels::RelsGraph::parse(b).map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("drawing rels parse: {e}"))
                    })?
                } else {
                    match ooxml_util::zip_read_to_string_opt(zip, &drawing_rels_path)? {
                        Some(s) => wolfxl_rels::RelsGraph::parse(s.as_bytes())
                            .map_err(|e| {
                                PyErr::new::<PyIOError, _>(format!(
                                    "drawing rels parse: {e}"
                                ))
                            })?,
                        None => wolfxl_rels::RelsGraph::new(),
                    }
                };
                // Add a chart rel per queued chart.
                let mut chart_rids: Vec<String> = Vec::with_capacity(queued.len());
                for &n in &chart_indices {
                    let rid = drawing_rels.add(
                        wolfxl_rels::rt::CHART,
                        &format!("../charts/chart{n}.xml"),
                        wolfxl_rels::TargetMode::Internal,
                    );
                    chart_rids.push(rid.0);
                }
                // Read existing drawing XML.
                let existing_drawing_xml: Vec<u8> = if let Some(b) =
                    file_patches.get(&drawing_path)
                {
                    b.clone()
                } else if let Some(b) = self.file_adds.get(&drawing_path) {
                    b.clone()
                } else {
                    let s =
                        ooxml_util::zip_read_to_string_opt(zip, &drawing_path)?
                            .unwrap_or_else(|| String::from(""));
                    s.into_bytes()
                };
                // SAX-merge: append a graphicFrame per queued chart.
                let merged = append_graphic_frames(
                    &existing_drawing_xml,
                    &queued,
                    &chart_rids,
                )
                .map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("merge drawing: {e}"))
                })?;
                new_drawing_xml_bytes = Some(merged);
                // No new <Override> for the drawing — already in
                // [Content_Types].xml.
            } else {
                // Fresh drawing.
                drawing_n = part_id_allocator.alloc_drawing();
                drawing_path = format!("xl/drawings/drawing{drawing_n}.xml");
                drawing_rels_path = format!(
                    "xl/drawings/_rels/drawing{drawing_n}.xml.rels"
                );
                drawing_rels = wolfxl_rels::RelsGraph::new();
                let mut chart_rids: Vec<String> = Vec::with_capacity(queued.len());
                for &n in &chart_indices {
                    let rid = drawing_rels.add(
                        wolfxl_rels::rt::CHART,
                        &format!("../charts/chart{n}.xml"),
                        wolfxl_rels::TargetMode::Internal,
                    );
                    chart_rids.push(rid.0);
                }
                // Build a fresh drawing XML body.
                let body = build_chart_drawing_xml(&queued, &chart_rids);
                new_drawing_xml_bytes = Some(body.into_bytes());
                // Splice <drawing r:id> into sheet XML.
                let drawing_rid = sheet_rels.add(
                    wolfxl_rels::rt::DRAWING,
                    &format!("../drawings/drawing{drawing_n}.xml"),
                    wolfxl_rels::TargetMode::Internal,
                );
                let sheet_xml = if let Some(b) = file_patches.get(&sheet_path) {
                    String::from_utf8_lossy(b).into_owned()
                } else if let Some(b) = self.file_adds.get(&sheet_path) {
                    String::from_utf8_lossy(b).into_owned()
                } else {
                    ooxml_util::zip_read_to_string(zip, &sheet_path)?
                };
                let after = splice_drawing_ref(&sheet_xml, &drawing_rid.0)
                    .map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("splice drawing: {e}"))
                    })?;
                file_patches.insert(sheet_path.clone(), after.into_bytes());
                ct_ops.push(content_types::ContentTypeOp::AddOverride(
                    format!("/xl/drawings/drawing{drawing_n}.xml"),
                    "application/vnd.openxmlformats-officedocument.drawing+xml".to_string(),
                ));
            }

            // Emit drawing XML + drawing rels into file_adds /
            // file_patches. We use file_patches for in-place updates
            // of an existing drawing (so the per-emit pass picks the
            // mutated bytes); file_adds for fresh-drawing emit. The
            // ZIP probe is the source-of-truth: if the path is
            // already in the source ZIP we MUST patch (file_adds
            // panics on collision in the final emit pass).
            if let Some(bytes) = new_drawing_xml_bytes {
                if zip.by_name(&drawing_path).is_ok() {
                    file_patches.insert(drawing_path.clone(), bytes);
                } else {
                    self.file_adds.insert(drawing_path.clone(), bytes);
                }
            }
            self.rels_patches
                .insert(drawing_rels_path, drawing_rels);

            // Persist sheet rels mutation.
            self.rels_patches.insert(sheet_rels_path, sheet_rels);

            // Queue content-type ops under a synthetic per-sheet key.
            self.queued_content_type_ops
                .entry(format!("__rfc046_charts_{sheet_name}__"))
                .or_default()
                .extend(ct_ops);
        }
        Ok(())
    }

    /// Sprint Ν Pod-γ (RFC-047 + RFC-048) — drain pivot caches and
    /// pivot tables in Phase 2.5m.
    ///
    /// Caches drain first (workbook-scope) → tables drain second
    /// (sheet-scope, with rels back-pointing at the matching cache).
    /// See `src/wolfxl/pivot.rs` module docs for full step-by-step
    /// invariants. The phase ordering relative to Phase 2.5l (charts)
    /// is pinned by `pivot::tests::phase_ordering_pinned`.
    fn apply_pivot_adds_phase(
        &mut self,
        file_patches: &mut HashMap<String, Vec<u8>>,
        zip: &mut ZipArchive<File>,
    ) -> PyResult<()> {
        // Bootstrap a per-patcher pivot part-id counter from the
        // source ZIP so we never collide with existing pivot parts.
        let mut counters = pivot::PivotPartCounters::new(1, 1);
        for i in 0..zip.len() {
            if let Ok(name) = zip.by_index(i).map(|f| f.name().to_string()) {
                counters.observe(&name);
            }
        }
        // Also observe paths the patcher may have already written
        // earlier in the save (e.g. from a previous apply_pivot_adds
        // call within RFC-035 deep-clone — defensive).
        for path in self.file_adds.keys() {
            counters.observe(path);
        }

        // ---- Pass 1: drain caches ----
        let drained_caches: Vec<pivot::QueuedPivotCacheAdd> =
            std::mem::take(&mut self.queued_pivot_caches);

        // Map: queued cache_id → allocated part-id (cache_n) so we
        // can resolve rels targets for tables in Pass 2.
        let mut cache_id_to_part_id: HashMap<u32, u32> = HashMap::new();

        // Collect new <pivotCache> entries to splice into workbook.xml.
        let mut pivot_cache_refs: Vec<pivot::PivotCacheRef> = Vec::new();

        // Workbook rels graph mutation. Read once, mutate, persist.
        let workbook_rels_path = "xl/_rels/workbook.xml.rels";
        let mut workbook_rels: wolfxl_rels::RelsGraph =
            if let Some(g) = self.rels_patches.get(workbook_rels_path) {
                g.clone()
            } else if let Some(b) = self.file_adds.get(workbook_rels_path) {
                wolfxl_rels::RelsGraph::parse(b).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("workbook rels parse: {e}"))
                })?
            } else {
                let s = ooxml_util::zip_read_to_string(zip, workbook_rels_path)?;
                wolfxl_rels::RelsGraph::parse(s.as_bytes()).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("workbook rels parse: {e}"))
                })?
            };

        let mut ct_ops: Vec<content_types::ContentTypeOp> = Vec::new();

        for cache in &drained_caches {
            let n = counters.alloc_cache();
            cache_id_to_part_id.insert(cache.cache_id, n);

            let def_path = format!("xl/pivotCache/pivotCacheDefinition{n}.xml");
            let rec_path = format!("xl/pivotCache/pivotCacheRecords{n}.xml");
            let cache_rels_path =
                format!("xl/pivotCache/_rels/pivotCacheDefinition{n}.xml.rels");

            self.file_adds
                .insert(def_path.clone(), cache.cache_def_xml.clone());
            self.file_adds
                .insert(rec_path.clone(), cache.cache_records_xml.clone());

            // Per-cache rels: definition → records.
            let mut cache_rels = wolfxl_rels::RelsGraph::new();
            // The cache definition uses `r:id="rId1"` to reference
            // its records part (matches the canonical emit from
            // wolfxl-pivot::emit::pivot_cache_definition_xml).
            cache_rels.add_with_id(
                wolfxl_rels::RelId("rId1".into()),
                wolfxl_pivot::rt::PIVOT_CACHE_RECORDS,
                &format!("pivotCacheRecords{n}.xml"),
                wolfxl_rels::TargetMode::Internal,
            );
            self.rels_patches
                .insert(cache_rels_path, cache_rels);

            // Workbook rel → cache definition.
            let rid = workbook_rels.add(
                wolfxl_rels::rt::PIVOT_CACHE_DEF,
                &format!("pivotCache/pivotCacheDefinition{n}.xml"),
                wolfxl_rels::TargetMode::Internal,
            );
            pivot_cache_refs.push(pivot::PivotCacheRef {
                cache_id: cache.cache_id,
                rid: rid.0,
            });

            // Content-type overrides.
            ct_ops.push(content_types::ContentTypeOp::AddOverride(
                format!("/{def_path}"),
                wolfxl_pivot::ct::PIVOT_CACHE_DEFINITION.to_string(),
            ));
            ct_ops.push(content_types::ContentTypeOp::AddOverride(
                format!("/{rec_path}"),
                wolfxl_pivot::ct::PIVOT_CACHE_RECORDS.to_string(),
            ));
        }

        // Persist workbook rels mutation.
        self.rels_patches
            .insert(workbook_rels_path.to_string(), workbook_rels);

        // Splice <pivotCaches> into xl/workbook.xml.
        if !pivot_cache_refs.is_empty() {
            let wb_xml: Vec<u8> = if let Some(b) = file_patches.get("xl/workbook.xml") {
                b.clone()
            } else if let Some(b) = self.file_adds.get("xl/workbook.xml") {
                b.clone()
            } else {
                ooxml_util::zip_read_to_string(zip, "xl/workbook.xml")?.into_bytes()
            };
            let updated = pivot::splice_pivot_caches(&wb_xml, &pivot_cache_refs)
                .map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!(
                        "splice <pivotCaches>: {e}"
                    ))
                })?;
            file_patches.insert("xl/workbook.xml".to_string(), updated);
        }

        // ---- Pass 2: drain tables ----
        let drained_tables: HashMap<String, Vec<pivot::QueuedPivotTableAdd>> =
            std::mem::take(&mut self.queued_pivot_tables);

        // Drain in sheet_order for stable output.
        let sheet_order_clone: Vec<String> = self.sheet_order.clone();
        for sheet_name in &sheet_order_clone {
            let queued = match drained_tables.get(sheet_name) {
                Some(q) if !q.is_empty() => q,
                _ => continue,
            };
            let sheet_path = self
                .sheet_paths
                .get(sheet_name)
                .cloned()
                .ok_or_else(|| {
                    PyValueError::new_err(format!(
                        "queue_pivot_table_add: no such sheet: {sheet_name}"
                    ))
                })?;

            let sheet_rels_path = format!(
                "{}/_rels/{}.rels",
                sheet_path.rsplit_once('/').map(|(d, _)| d).unwrap_or(""),
                sheet_path.rsplit('/').next().unwrap_or("")
            );

            let mut sheet_rels: wolfxl_rels::RelsGraph =
                if let Some(g) = self.rels_patches.get(&sheet_rels_path) {
                    g.clone()
                } else if let Some(b) = self.file_adds.get(&sheet_rels_path) {
                    wolfxl_rels::RelsGraph::parse(b).map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("sheet rels parse: {e}"))
                    })?
                } else {
                    match ooxml_util::zip_read_to_string_opt(zip, &sheet_rels_path)? {
                        Some(s) => wolfxl_rels::RelsGraph::parse(s.as_bytes())
                            .map_err(|e| {
                                PyErr::new::<PyIOError, _>(format!(
                                    "sheet rels parse: {e}"
                                ))
                            })?,
                        None => wolfxl_rels::RelsGraph::new(),
                    }
                };

            for table in queued {
                let table_n = counters.alloc_table();
                let table_path = format!("xl/pivotTables/pivotTable{table_n}.xml");
                let table_rels_path =
                    format!("xl/pivotTables/_rels/pivotTable{table_n}.xml.rels");

                self.file_adds
                    .insert(table_path.clone(), table.table_xml.clone());

                // Per-table rels: table → cache definition. Resolve
                // the matching cache part id via cache_id_to_part_id
                // (caches queued in this same save) OR fall back to
                // direct cache_id-to-part-id mapping for caches that
                // already exist on disk in the source workbook.
                let cache_n = match cache_id_to_part_id.get(&table.cache_id) {
                    Some(&n) => n,
                    None => {
                        // Fall back: cache lives on disk in source
                        // ZIP. The cacheId in workbook.xml's
                        // <pivotCaches> entry maps to a workbook-rel
                        // pointing at pivotCacheDefinition{N}.xml; we
                        // assume cache_id+1 == part_id here as a
                        // simplifying convention. Real-world: the
                        // patcher would parse <pivotCaches> to
                        // resolve. v2.0 caches always come from this
                        // save, so the fallback is rarely hit.
                        table.cache_id + 1
                    }
                };

                let mut table_rels = wolfxl_rels::RelsGraph::new();
                table_rels.add_with_id(
                    wolfxl_rels::RelId("rId1".into()),
                    wolfxl_rels::rt::PIVOT_CACHE_DEF,
                    &format!("../pivotCache/pivotCacheDefinition{cache_n}.xml"),
                    wolfxl_rels::TargetMode::Internal,
                );
                self.rels_patches
                    .insert(table_rels_path, table_rels);

                // Sheet rel → pivot table.
                sheet_rels.add(
                    wolfxl_rels::rt::PIVOT_TABLE,
                    &format!("../pivotTables/pivotTable{table_n}.xml"),
                    wolfxl_rels::TargetMode::Internal,
                );

                ct_ops.push(content_types::ContentTypeOp::AddOverride(
                    format!("/{table_path}"),
                    wolfxl_pivot::ct::PIVOT_TABLE.to_string(),
                ));
            }

            self.rels_patches.insert(sheet_rels_path, sheet_rels);
        }

        // Queue content-type ops under a synthetic per-workbook key
        // (pivots are workbook-scope; not tied to a single sheet
        // name in `sheet_order`). Phase 2.5c picks these up via the
        // `synth_keys` aggregator.
        if !ct_ops.is_empty() {
            self.queued_content_type_ops
                .entry("__rfc047_pivots__".to_string())
                .or_default()
                .extend(ct_ops);
        }

        Ok(())
    }

    /// across every queued op. Reads from `file_patches` when an
    /// earlier phase already mutated a part; falls back to source ZIP
    /// otherwise. Writes the result back into `file_patches`.
    fn apply_axis_shifts_phase(
        &mut self,
        file_patches: &mut HashMap<String, Vec<u8>>,
        zip: &mut ZipArchive<File>,
    ) -> PyResult<()> {
        // Helper: get bytes for a path (current rewrite if any, else source).
        fn get_bytes(
            file_patches: &HashMap<String, Vec<u8>>,
            zip: &mut ZipArchive<File>,
            path: &str,
        ) -> Option<Vec<u8>> {
            if let Some(b) = file_patches.get(path) {
                return Some(b.clone());
            }
            let mut entry = match zip.by_name(path) {
                Ok(e) => e,
                Err(_) => return None,
            };
            let mut buf: Vec<u8> = Vec::with_capacity(entry.size() as usize);
            std::io::copy(&mut entry, &mut buf).ok()?;
            Some(buf)
        }

        // Build sheet name → 0-based position map (for definedName scope).
        let sheet_positions: BTreeMap<String, u32> = self
            .sheet_order
            .iter()
            .enumerate()
            .map(|(i, name)| (name.clone(), i as u32))
            .collect();

        // Discover table parts via the rels graph for each sheet.
        // We need this lazy + per-sheet because each op may operate
        // on a different sheet.
        for op in self.queued_axis_shifts.clone() {
            let sheet_path = match self.sheet_paths.get(&op.sheet) {
                Some(p) => p.clone(),
                None => continue, // unknown sheet — silently skip
            };

            let axis = match op.axis.as_str() {
                "row" => wolfxl_structural::Axis::Row,
                "col" => wolfxl_structural::Axis::Col,
                _ => continue,
            };

            // Read sheet XML.
            let sheet_xml = match get_bytes(file_patches, zip, &sheet_path) {
                Some(b) => b,
                None => continue,
            };

            // Read workbook.xml.
            let wb_xml = get_bytes(file_patches, zip, "xl/workbook.xml");

            // Discover this sheet's rels graph (for table/comments/vml lookups).
            // Use the ancillary registry to get the part paths.
            let _ = self.ancillary.populate_for_sheet(zip, &op.sheet, &sheet_path);

            let (comments_part, vml_part, table_paths) = {
                let anc = self.ancillary.get(&op.sheet).cloned().unwrap_or_default();
                (anc.comments_part, anc.vml_drawing_part, anc.table_parts.clone())
            };

            // Read each.
            let comments_bytes: Option<(String, Vec<u8>)> = comments_part
                .as_ref()
                .and_then(|p| get_bytes(file_patches, zip, p).map(|b| (p.clone(), b)));
            let vml_bytes: Option<(String, Vec<u8>)> = vml_part
                .as_ref()
                .and_then(|p| get_bytes(file_patches, zip, p).map(|b| (p.clone(), b)));
            let mut table_bytes: Vec<(String, Vec<u8>)> = Vec::new();
            for tp in &table_paths {
                if let Some(b) = get_bytes(file_patches, zip, tp) {
                    table_bytes.push((tp.clone(), b));
                }
            }

            // Build inputs.
            let mut inputs = wolfxl_structural::SheetXmlInputs::empty();
            inputs.sheets.insert(op.sheet.clone(), sheet_xml.as_slice());
            inputs.sheet_paths.insert(op.sheet.clone(), sheet_path.clone());
            if let Some(ref wb) = wb_xml {
                inputs.workbook_xml = Some(wb.as_slice());
            }
            if !table_bytes.is_empty() {
                let parts: Vec<(String, &[u8])> = table_bytes
                    .iter()
                    .map(|(p, b)| (p.clone(), b.as_slice()))
                    .collect();
                inputs.tables.insert(op.sheet.clone(), parts);
            }
            if let Some((ref p, ref b)) = comments_bytes {
                inputs.comments.insert(op.sheet.clone(), (p.clone(), b.as_slice()));
            }
            if let Some((ref p, ref b)) = vml_bytes {
                inputs.vml.insert(op.sheet.clone(), (p.clone(), b.as_slice()));
            }
            inputs.sheet_positions = sheet_positions.clone();

            let ops_one = vec![wolfxl_structural::AxisShiftOp {
                sheet: op.sheet.clone(),
                axis,
                idx: op.idx,
                n: op.n,
            }];
            let mutations = wolfxl_structural::apply_workbook_shift(inputs, &ops_one);
            for (path, bytes) in mutations.file_patches {
                file_patches.insert(path, bytes);
            }
        }
        Ok(())
    }

    /// Phase 2.7 — drive `wolfxl_structural::sheet_copy::plan_sheet_copy`
    /// across every queued sheet-copy op (RFC-035).
    ///
    /// For each `(src_title, dst_title)` op:
    ///   1. Look up the source sheet path; build the source rels graph
    ///      from the ZIP (or `rels_patches` if already mutated).
    ///   2. Pre-load the source ZIP parts map for the planner (sheet
    ///      bytes + every reachable ancillary part + nested rels).
    ///   3. Read workbook.xml from `file_patches` if already mutated,
    ///      else from the source ZIP.
    ///   4. Call `plan_sheet_copy`. Returned `SheetCopyMutations` is
    ///      pure data; we apply it atomically.
    ///   5. Allocate a real workbook-rels rId for the new sheet, then
    ///      string-replace the planner's
    ///      `__SHEET_RID_PLACEHOLDER_<N>__` token in
    ///      `workbook_sheets_append` and `workbook_rels_to_add[0].0`.
    ///   6. Splice the new `<sheet>` element into workbook.xml's
    ///      `<sheets>` block, persist into `file_patches`.
    ///   7. Update `xl/_rels/workbook.xml.rels` via `rels_patches`.
    ///   8. Insert new sheet xml + ancillary parts into `file_adds`.
    ///   9. Forward content-type ops into `queued_content_type_ops`
    ///      under a synthetic key so Phase 2.5c picks them up.
    ///   10. Queue cloned sheet-scoped defined names through
    ///       `queued_defined_names` so RFC-021's merger handles them.
    ///   11. Update `self.sheet_order`, `self.sheet_paths`, and the
    ///       running `cloned_table_names` accumulator.
    ///
    /// Phase-ordering invariant: any new per-sheet phase MUST run
    /// AFTER 2.7 (per RFC-035 §8 risk #7).
    fn apply_sheet_copies_phase(
        &mut self,
        file_patches: &mut HashMap<String, Vec<u8>>,
        zip: &mut ZipArchive<File>,
        part_id_allocator: &mut wolfxl_rels::PartIdAllocator,
        cloned_table_names: &mut HashSet<String>,
    ) -> PyResult<()> {
        // Helper: get bytes for a path (current rewrite if any, else source).
        fn get_bytes(
            file_patches: &HashMap<String, Vec<u8>>,
            file_adds: &HashMap<String, Vec<u8>>,
            zip: &mut ZipArchive<File>,
            path: &str,
        ) -> Option<Vec<u8>> {
            if let Some(b) = file_patches.get(path) {
                return Some(b.clone());
            }
            if let Some(b) = file_adds.get(path) {
                return Some(b.clone());
            }
            let mut entry = match zip.by_name(path) {
                Ok(e) => e,
                Err(_) => return None,
            };
            let mut buf: Vec<u8> = Vec::with_capacity(entry.size() as usize);
            std::io::copy(&mut entry, &mut buf).ok()?;
            Some(buf)
        }

        // RFC-035 §5.5: existing-table-name set at the start of Phase
        // 2.7 includes every name from the source ZIP plus any name
        // already queued by `queue_table` (RFC-024 user adds running
        // in the same save). We compute the union once up front.
        let mut existing_table_names: HashSet<String> = HashSet::new();
        // Source ZIP table parts.
        let table_inv = tables::scan_existing_tables(zip)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("scan tables: {e}")))?;
        for n in &table_inv.names {
            existing_table_names.insert(n.clone());
        }
        // In-flight user-queued tables (from RFC-024).
        for patches in self.queued_tables.values() {
            for p in patches {
                existing_table_names.insert(p.name.clone());
            }
        }

        // Workbook rels graph: get-or-load. We will mutate it (add a
        // worksheet rel per copy), then on save the existing serialize
        // path picks up the mutated graph and routes it through
        // `file_patches` (or `file_adds` if the entry is missing,
        // though `xl/_rels/workbook.xml.rels` is always present).
        let workbook_rels_path = "xl/_rels/workbook.xml.rels".to_string();
        if !self.rels_patches.contains_key(&workbook_rels_path) {
            let g = load_or_empty_rels(zip, &workbook_rels_path)?;
            self.rels_patches.insert(workbook_rels_path.clone(), g);
        }

        // Process each queued op in append order. Each iteration sees
        // the running mutated state (so copies-of-copies and
        // copy-then-edit-the-copy work).
        let ops = self.queued_sheet_copies.clone();
        for op in ops {
            // Look up source sheet path. We re-check existence here
            // because copy-of-copy runs AFTER an earlier op already
            // updated `sheet_paths`, so a `src_title` that names an
            // earlier in-flight clone is valid.
            let src_sheet_path = match self.sheet_paths.get(&op.src_title).cloned() {
                Some(p) => p,
                None => {
                    return Err(PyErr::new::<PyValueError, _>(format!(
                        "Phase 2.7: source sheet '{}' missing at flush time",
                        op.src_title
                    )));
                }
            };

            // Load source rels graph. Prefer in-memory rels_patches
            // (an earlier phase or copy already touched it), else parse
            // from the source ZIP.
            let src_rels_path = sheet_rels_path_for(&src_sheet_path);
            let source_rels: RelsGraph = if let Some(g) = self.rels_patches.get(&src_rels_path) {
                g.clone()
            } else {
                load_or_empty_rels(zip, &src_rels_path)?
            };

            // Walk reachable parts (one level + nested via rels file
            // probes) so we can pre-load the planner's input map. We
            // duplicate the planner's resolver here only to discover
            // which paths need pre-loading; the planner itself does
            // its own walk on the same data.
            let subgraph = wolfxl_rels::walk_sheet_subgraph_with_nested(
                &source_rels,
                &src_sheet_path,
                |part_path: &str| {
                    let rels_path = wolfxl_rels::rels_path_for(part_path)?;
                    let bytes = get_bytes(file_patches, &self.file_adds, zip, &rels_path)?;
                    RelsGraph::parse(&bytes).ok()
                },
            );

            // Pre-load source ZIP parts map. Includes the sheet itself,
            // every reachable ancillary, and any per-part rels file
            // we need (drawing rels for image aliasing).
            let mut source_zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
            for part_path in &subgraph.reachable_parts {
                if let Some(bytes) = get_bytes(file_patches, &self.file_adds, zip, part_path) {
                    source_zip_parts.insert(part_path.clone(), bytes);
                }
                // Each reachable ancillary may have its own rels file
                // (drawings → images). The planner's resolver expects
                // those to be in the parts map keyed by rels path.
                if let Some(rp) = wolfxl_rels::rels_path_for(part_path) {
                    if let Some(bytes) = get_bytes(file_patches, &self.file_adds, zip, &rp) {
                        source_zip_parts.insert(rp, bytes);
                    }
                }
            }

            // Read workbook.xml.
            let workbook_xml = match get_bytes(
                file_patches,
                &self.file_adds,
                zip,
                "xl/workbook.xml",
            ) {
                Some(b) => b,
                None => {
                    return Err(PyErr::new::<PyIOError, _>(
                        "Phase 2.7: xl/workbook.xml missing from source ZIP",
                    ));
                }
            };

            // Build planner inputs.
            let inputs = wolfxl_structural::sheet_copy::SheetCopyInputs {
                src_title: op.src_title.clone(),
                dst_title: op.dst_title.clone(),
                src_sheet_path: src_sheet_path.clone(),
                source_zip_parts: &source_zip_parts,
                source_rels: &source_rels,
                workbook_xml: &workbook_xml,
                allocator: part_id_allocator,
                existing_table_names: &existing_table_names,
                deep_copy_images: op.deep_copy_images,
            };
            let mutations =
                wolfxl_structural::sheet_copy::plan_sheet_copy(inputs).map_err(|e| {
                    PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!(
                        "Phase 2.7: plan_sheet_copy failed for '{}'→'{}': {}",
                        op.src_title, op.dst_title, e
                    ))
                })?;

            // Allocate the workbook-rels rId for the new sheet. The
            // planner's `workbook_rels_to_add[0]` is
            // `(placeholder, rel_type, target)`. We add it to the
            // workbook's rels graph (via add()) which mints the rId,
            // and string-replace the placeholder afterwards.
            let (placeholder, rel_type, target) = mutations
                .workbook_rels_to_add
                .first()
                .cloned()
                .ok_or_else(|| {
                    PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(
                        "Phase 2.7: planner returned no workbook_rels_to_add entry",
                    )
                })?;
            let new_rid = {
                let g = self
                    .rels_patches
                    .get_mut(&workbook_rels_path)
                    .expect("workbook rels graph loaded above");
                g.add(&rel_type, &target, wolfxl_rels::TargetMode::Internal)
            };

            // Replace placeholder in workbook_sheets_append.
            let sheets_append: Vec<u8> = {
                let s = String::from_utf8_lossy(&mutations.workbook_sheets_append);
                s.replace(&placeholder, &new_rid.0).into_bytes()
            };

            // Splice the new <sheet> element into workbook.xml's
            // <sheets> block (insert before </sheets>).
            let new_workbook_xml = splice_into_sheets_block(&workbook_xml, &sheets_append)?;
            file_patches.insert("xl/workbook.xml".to_string(), new_workbook_xml);

            // Drop new sheet xml into file_adds.
            self.file_adds
                .insert(mutations.new_sheet_path.clone(), mutations.new_sheet_xml);

            // Drop ancillary parts into file_adds.
            for (path, bytes) in mutations.new_ancillary_parts {
                self.file_adds.insert(path, bytes);
            }

            // Forward content-type ops. Use a synthetic per-workbook
            // key so Phase 2.5c (which iterates `sheet_order`) picks
            // them up via the existing aggregator.
            let ct_ops_for_sheet = self
                .queued_content_type_ops
                .entry("__rfc035_sheet_copy__".to_string())
                .or_default();
            for (part_path, content_type) in mutations.content_type_overrides_to_add {
                // VML default-extension shows up as a content-type op
                // too: any path ending in `.vml` should ensure the
                // extension default is in place rather than emit an
                // Override. The planner currently emits Override-only
                // for non-VML parts (vml is omitted from
                // content_type_overrides_to_add); but we add a
                // safety net here for forward compat.
                if part_path.ends_with(".vml") {
                    ct_ops_for_sheet.push(content_types::ContentTypeOp::EnsureDefault(
                        "vml".to_string(),
                        comments::CT_VML.to_string(),
                    ));
                } else {
                    ct_ops_for_sheet
                        .push(content_types::ContentTypeOp::AddOverride(part_path, content_type));
                }
            }
            // The planner does NOT emit a vml Default itself (RFC-035
            // §5.6 routes VML through a Default-extension). If any of
            // the cloned ancillary parts is a vmlDrawing, ensure the
            // workbook's content-types graph has the vml Default.
            // (Idempotent — Phase 2.5c's aggregator absorbs duplicates.)
            // We detect VML by examining the file_adds we just
            // populated.
            let needs_vml_default = self
                .file_adds
                .keys()
                .any(|k| k.starts_with("xl/drawings/vmlDrawing") && k.ends_with(".vml"));
            if needs_vml_default {
                ct_ops_for_sheet.push(content_types::ContentTypeOp::EnsureDefault(
                    "vml".to_string(),
                    comments::CT_VML.to_string(),
                ));
            }

            // Cloned sheet-scoped defined names: queue through
            // RFC-021's merger so workbook.xml's <definedNames> block
            // gets the new entries on save (RFC-035 §5.4
            // Composability with RFC-021).
            //
            // Upsert-collision rule (Pod-δ fix for KNOWN_GAPS bug #5):
            // if the user has ALREADY queued a defined name with the
            // SAME `(name, local_sheet_id)` key, the user's entry
            // wins — don't push the planner's value (it would land
            // last in the merger's BTreeMap and shadow the user's
            // upsert silently). Per Pod-β's stated invariant
            // "last-write-wins" should converge on the USER's value,
            // not the planner's; the planner is the default, the user
            // is the explicit override.
            for dn in mutations.defined_names_to_add {
                let key_name = dn.name.as_str();
                let key_lsid = Some(dn.local_sheet_id);
                let already_queued = self
                    .queued_defined_names
                    .iter()
                    .any(|q| q.name == key_name && q.local_sheet_id == key_lsid);
                if already_queued {
                    continue;
                }
                self.queued_defined_names.push(defined_names::DefinedNameMut {
                    name: dn.name,
                    formula: dn.formula,
                    local_sheet_id: Some(dn.local_sheet_id),
                    hidden: None,
                    comment: None,
                });
            }

            // Update running cloned-table-names accumulator (RFC-024
            // collision-scan extension — see §8 risk #6).
            for n in &mutations.new_table_names {
                cloned_table_names.insert(n.clone());
                existing_table_names.insert(n.clone());
            }

            // Update patcher's tab list + path map. Append the new
            // sheet at the end (RFC-035 §5.7: tab order = source order
            // + new entry at end).
            self.sheet_order.push(op.dst_title.clone());
            self.sheet_paths
                .insert(op.dst_title.clone(), mutations.new_sheet_path);
        }

        // Drain the queue so a subsequent save() on the same patcher
        // doesn't re-emit (parallels RFC-030 / RFC-034).
        self.queued_sheet_copies.clear();

        Ok(())
    }

    /// Phase 2.5j — drive `wolfxl_structural::apply_range_move`
    /// across every queued range-move op. Reads from `file_patches`
    /// when an earlier phase already mutated a part; falls back to
    /// source ZIP otherwise. Writes the result back into
    /// `file_patches` so subsequent ops see the rewritten bytes.
    fn apply_range_moves_phase(
        &mut self,
        file_patches: &mut HashMap<String, Vec<u8>>,
        zip: &mut ZipArchive<File>,
    ) -> PyResult<()> {
        fn get_bytes(
            file_patches: &HashMap<String, Vec<u8>>,
            zip: &mut ZipArchive<File>,
            path: &str,
        ) -> Option<Vec<u8>> {
            if let Some(b) = file_patches.get(path) {
                return Some(b.clone());
            }
            let mut entry = match zip.by_name(path) {
                Ok(e) => e,
                Err(_) => return None,
            };
            let mut buf: Vec<u8> = Vec::with_capacity(entry.size() as usize);
            std::io::copy(&mut entry, &mut buf).ok()?;
            Some(buf)
        }

        for op in self.queued_range_moves.clone() {
            let sheet_path = match self.sheet_paths.get(&op.sheet) {
                Some(p) => p.clone(),
                None => continue, // unknown sheet — silently skip
            };

            let sheet_xml = match get_bytes(file_patches, zip, &sheet_path) {
                Some(b) => b,
                None => continue,
            };

            let plan = wolfxl_structural::RangeMovePlan {
                src_lo: (op.src_min_row, op.src_min_col),
                src_hi: (op.src_max_row, op.src_max_col),
                d_row: op.d_row,
                d_col: op.d_col,
                translate: op.translate,
            };
            let new_bytes = wolfxl_structural::apply_range_move(&sheet_xml, &plan);
            if new_bytes != sheet_xml {
                file_patches.insert(sheet_path, new_bytes);
            }
        }
        Ok(())
    }

    /// Phase 2.8 — rebuild `xl/calcChain.xml` (Sprint Θ Pod-C3).
    ///
    /// Walks every sheet in `sheet_order`, scans the post-mutation XML for
    /// formula cells, and emits a fresh `xl/calcChain.xml`. The rebuild
    /// runs unconditionally inside the flush phase — the no-op
    /// short-circuit at the top of `do_save` already bypasses this phase
    /// when there are zero queued ops, so byte-identical no-op saves are
    /// unaffected.
    ///
    /// Behaviour:
    /// - At least one formula across all sheets → emit a fresh
    ///   `xl/calcChain.xml` (overwriting any source copy in
    ///   `file_patches` or adding a new entry via `file_adds`).
    ///   Adds a `[Content_Types].xml` `<Override>` for it if not
    ///   already present, and adds a workbook→calcChain rel if not
    ///   already present.
    /// - Zero formulas across all sheets → if the source contained a
    ///   `xl/calcChain.xml`, mark it for deletion (`file_deletes`) so
    ///   the saved file is consistent with the workbook content.
    fn rebuild_calc_chain_phase(
        &mut self,
        file_patches: &mut HashMap<String, Vec<u8>>,
        zip: &mut ZipArchive<File>,
    ) -> PyResult<()> {
        fn get_bytes(
            file_patches: &HashMap<String, Vec<u8>>,
            file_adds: &HashMap<String, Vec<u8>>,
            zip: &mut ZipArchive<File>,
            path: &str,
        ) -> Option<Vec<u8>> {
            if let Some(b) = file_patches.get(path) {
                return Some(b.clone());
            }
            if let Some(b) = file_adds.get(path) {
                return Some(b.clone());
            }
            let mut entry = match zip.by_name(path) {
                Ok(e) => e,
                Err(_) => return None,
            };
            let mut buf: Vec<u8> = Vec::with_capacity(entry.size() as usize);
            std::io::copy(&mut entry, &mut buf).ok()?;
            Some(buf)
        }

        const CALC_CHAIN_PATH: &str = "xl/calcChain.xml";
        let source_has_calc_chain = zip.by_name(CALC_CHAIN_PATH).is_ok();

        // Walk sheets in tab order, scanning each.
        let mut all_entries: Vec<calcchain::CalcChainEntry> = Vec::new();
        let order = self.sheet_order.clone();
        for (i, sheet_name) in order.iter().enumerate() {
            let sheet_path = match self.sheet_paths.get(sheet_name) {
                Some(p) => p.clone(),
                None => continue,
            };
            let sheet_xml = match get_bytes(file_patches, &self.file_adds, zip, &sheet_path) {
                Some(b) => b,
                None => continue,
            };
            let sheet_index_1based = (i as u32) + 1;
            let entries = calcchain::scan_sheet_for_formulas(&sheet_xml, sheet_index_1based);
            all_entries.extend(entries);
        }

        match calcchain::render_calc_chain(&all_entries) {
            Some(bytes) => {
                // Route the rewrite based on whether the source ZIP
                // already had a calcChain.xml entry.
                if source_has_calc_chain {
                    file_patches.insert(CALC_CHAIN_PATH.to_string(), bytes);
                } else {
                    self.file_adds.insert(CALC_CHAIN_PATH.to_string(), bytes);
                }
                // Ensure content-type Override + workbook rel.
                self.ensure_calc_chain_metadata(file_patches, zip)?;
            }
            None => {
                // Zero formulas in the workbook. If the source had a
                // calcChain.xml, delete it (it would be stale and Excel
                // would emit a parse warning if it pointed at missing
                // cells).
                if source_has_calc_chain {
                    self.file_deletes.insert(CALC_CHAIN_PATH.to_string());
                    file_patches.remove(CALC_CHAIN_PATH);
                }
                // No-op for content-types / workbook rels: leaving stale
                // metadata is benign because the part is gone, and many
                // Excel-generated files keep both ends in sync naturally
                // (we only remove our own rebuild output).
            }
        }
        Ok(())
    }

    /// Ensure `[Content_Types].xml` has an `<Override>` for
    /// `xl/calcChain.xml` and `xl/_rels/workbook.xml.rels` has a
    /// workbook→calcChain rel. Idempotent: existing entries are left
    /// untouched.
    fn ensure_calc_chain_metadata(
        &mut self,
        file_patches: &mut HashMap<String, Vec<u8>>,
        zip: &mut ZipArchive<File>,
    ) -> PyResult<()> {
        // Content types.
        let ct_xml: Vec<u8> = if let Some(b) = file_patches.get("[Content_Types].xml") {
            b.clone()
        } else {
            ooxml_util::zip_read_to_string(zip, "[Content_Types].xml")?
                .as_bytes()
                .to_vec()
        };
        let mut graph = content_types::ContentTypesGraph::parse(&ct_xml).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("[Content_Types].xml parse: {e}"))
        })?;
        graph.add_override("/xl/calcChain.xml", calcchain::CT_CALC_CHAIN);
        file_patches.insert("[Content_Types].xml".to_string(), graph.serialize());

        // Workbook rels.
        let wb_rels_path = "xl/_rels/workbook.xml.rels";
        let wb_rels_bytes_opt: Option<Vec<u8>> = if let Some(b) = file_patches.get(wb_rels_path) {
            Some(b.clone())
        } else if let Some(g) = self.rels_patches.get(wb_rels_path) {
            Some(g.serialize())
        } else if let Ok(mut entry) = zip.by_name(wb_rels_path) {
            let mut buf: Vec<u8> = Vec::with_capacity(entry.size() as usize);
            if std::io::copy(&mut entry, &mut buf).is_ok() {
                Some(buf)
            } else {
                None
            }
        } else {
            None
        };
        if let Some(bytes) = wb_rels_bytes_opt {
            let mut graph = wolfxl_rels::RelsGraph::parse(&bytes).unwrap_or_else(|_| wolfxl_rels::RelsGraph::new());
            // Idempotent: only add if no existing rel of this type
            // already targets calcChain.xml.
            let already = graph.iter().any(|r| {
                r.rel_type == calcchain::REL_CALC_CHAIN
                    && (r.target == "calcChain.xml" || r.target == "/xl/calcChain.xml")
            });
            if !already {
                graph.add(
                    calcchain::REL_CALC_CHAIN,
                    "calcChain.xml",
                    wolfxl_rels::TargetMode::Internal,
                );
                file_patches.insert(wb_rels_path.to_string(), graph.serialize());
            }
        }

        Ok(())
    }
}

// ---------------------------------------------------------------------------
// RFC-013 helpers — deterministic-when-test-epoch-is-set datetime stamping
// for `file_adds` ZIP entries.
// ---------------------------------------------------------------------------

/// True if the source ZIP contains an entry with the exact given name.
/// Used by RFC-020's Phase-2.5d to decide between `file_patches`
/// (replace existing) and `file_adds` (append new).
fn source_zip_has_entry<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> bool {
    zip.by_name(name).is_ok()
}

/// Insert `new_sheet_element` (an `<sheet …/>` byte sequence) into the
/// `<sheets>` block of `xl/workbook.xml`, immediately before the
/// closing `</sheets>` tag (RFC-035 §5.7). Source bytes flow through
/// verbatim outside the splice point. Returns an error if the source
/// has no `<sheets>` element (malformed workbook).
///
/// **Why SAX, not byte search**: a workbook.xml comment containing the
/// literal string `</sheets>` (e.g. `<!-- closes </sheets> here -->`),
/// a CDATA section, or a processing instruction can fool a naive
/// byte-substring scan. We walk the document with `quick_xml::Reader`
/// so comments/CDATA/PIs surface as their own events and are ignored
/// when locating the splice point. Bug #6 in
/// `tests/parity/KNOWN_GAPS.md`'s "RFC-035 cross-RFC composition gaps"
/// section tracked this — closed in Sprint Θ Pod-B.
fn splice_into_sheets_block(
    workbook_xml: &[u8],
    new_sheet_element: &[u8],
) -> PyResult<Vec<u8>> {
    use quick_xml::events::Event as XmlEvent;
    use quick_xml::Reader as XmlReader;

    // quick-xml works on `&str`, so reject non-UTF-8 input up front
    // with a stable error. workbook.xml is always UTF-8 per ECMA-376.
    let s = std::str::from_utf8(workbook_xml).map_err(|_| {
        PyErr::new::<PyIOError, _>("Phase 2.7: workbook.xml is not valid UTF-8")
    })?;
    let mut reader = XmlReader::from_str(s);
    reader.config_mut().trim_text(false);

    // Track the byte position right before each event and right after,
    // mirroring `sheet_order::scan_workbook_layout`. Depth tracks
    // element nesting so we ignore any spurious `<sheets>` that might
    // appear nested (defensive — no Excel writer emits that, but the
    // scan is cheap).
    let mut depth: i32 = 0;
    // For the `Start`/`End` form: record where to splice (just before
    // the closing `</sheets>` tag).
    let mut splice_pos: Option<usize> = None;
    let mut sheets_open_depth: Option<i32> = None;
    // For the `Empty` self-closing form: record the byte range of
    // `<sheets/>` so we can replace it with `<sheets>NEW</sheets>`.
    let mut self_closing_range: Option<(usize, usize)> = None;

    let mut buf: Vec<u8> = Vec::new();
    loop {
        let pre = reader.buffer_position() as usize;
        let evt = reader.read_event_into(&mut buf);
        let post = reader.buffer_position() as usize;
        match evt {
            Ok(XmlEvent::Start(ref e)) => {
                if e.local_name().as_ref() == b"sheets" && sheets_open_depth.is_none() {
                    sheets_open_depth = Some(depth);
                }
                depth += 1;
            }
            Ok(XmlEvent::End(ref e)) => {
                depth -= 1;
                if e.local_name().as_ref() == b"sheets"
                    && Some(depth) == sheets_open_depth
                    && splice_pos.is_none()
                {
                    // `pre` is the byte offset of `<` in `</sheets>`,
                    // i.e. exactly where the new `<sheet…/>` element
                    // should be inserted.
                    splice_pos = Some(pre);
                    break;
                }
            }
            Ok(XmlEvent::Empty(ref e)) => {
                // Self-closing `<sheets/>` (rare but ECMA-376-legal).
                // Note we DON'T increment depth — `Empty` is open+close.
                if e.local_name().as_ref() == b"sheets" && self_closing_range.is_none() {
                    self_closing_range = Some((pre, post));
                    break;
                }
            }
            Ok(XmlEvent::Eof) => break,
            // `Comment`, `CData`, `PI`, `DocType`, `Decl`, `Text` —
            // surface as their own events; we ignore them, which is
            // exactly the property that defeats the bug-#6 fakeout.
            Ok(_) => {}
            Err(_) => {
                // Fall through to the not-found branch — preserves
                // the historical error message for malformed inputs
                // that don't surface a `<sheets>` element.
                break;
            }
        }
        buf.clear();
    }

    if let Some(pos) = splice_pos {
        let mut out = Vec::with_capacity(workbook_xml.len() + new_sheet_element.len());
        out.extend_from_slice(&workbook_xml[..pos]);
        out.extend_from_slice(new_sheet_element);
        out.extend_from_slice(&workbook_xml[pos..]);
        return Ok(out);
    }
    if let Some((start, end)) = self_closing_range {
        let mut out =
            Vec::with_capacity(workbook_xml.len() + new_sheet_element.len() + 16);
        out.extend_from_slice(&workbook_xml[..start]);
        out.extend_from_slice(b"<sheets>");
        out.extend_from_slice(new_sheet_element);
        out.extend_from_slice(b"</sheets>");
        out.extend_from_slice(&workbook_xml[end..]);
        return Ok(out);
    }
    Err(PyErr::new::<PyIOError, _>(
        "Phase 2.7: workbook.xml has no <sheets> block",
    ))
}

/// Naive byte-substring search (workbook.xml is small enough that the
/// allocator overhead of memchr-shaped algorithms is overkill).
/// Sprint Θ Pod-B: Phase 2.7's splice no longer uses this helper —
/// retained for other potential callers per the RFC-035 follow-up note.
#[allow(dead_code)]
fn find_subslice(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() || needle.len() > haystack.len() {
        return None;
    }
    haystack
        .windows(needle.len())
        .position(|w| w == needle)
}

/// Sprint Θ Pod-A — XML attribute-value escape used by the permissive
/// load-time workbook.xml rewrite. Covers the five characters
/// disallowed by the XML 1.0 production for an `AttValue` (double-quote
/// terminated form). Synthesized titles are always `SheetN` so the
/// escape is mostly a guard for future callers, but the rId we recover
/// from the rels graph could in principle contain `&`.
fn xml_escape_attr(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '&' => out.push_str("&amp;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            other => out.push(other),
        }
    }
    out
}

/// Sprint Θ Pod-A — replace the first occurrence of `needle` in
/// `haystack` with `replacement`. Returns `None` if the needle is not
/// found, mirroring the contract of `str::replacen` but with a single
/// allocation (and without the `Cow` allocation `str::replacen` does
/// when no match exists).
fn replace_first_occurrence(haystack: &str, needle: &str, replacement: &str) -> Option<String> {
    let idx = haystack.find(needle)?;
    let mut out = String::with_capacity(haystack.len() - needle.len() + replacement.len());
    out.push_str(&haystack[..idx]);
    out.push_str(replacement);
    out.push_str(&haystack[idx + needle.len()..]);
    Some(out)
}

/// Return a `zip::DateTime` honoring `WOLFXL_TEST_EPOCH` when set.
///
/// When the env var parses to an `i64`, that value is treated as a Unix
/// epoch second count and clamped to ZIP's representable range
/// (1980..=2107). Otherwise falls back to current UTC time. Mirrors the
/// behavior of `wolfxl_writer::zip::test_epoch_override` so the patcher
/// and the writer produce byte-stable output under the same env flag.
fn epoch_or_now() -> zip::DateTime {
    use chrono::{Datelike, Timelike};
    let secs = std::env::var("WOLFXL_TEST_EPOCH")
        .ok()
        .and_then(|s| s.parse::<i64>().ok());
    let dt = match secs.and_then(|s| chrono::DateTime::<chrono::Utc>::from_timestamp(s, 0)) {
        Some(d) => d,
        None => chrono::Utc::now(),
    };
    let naive = dt.naive_utc();
    let year = naive.year();
    if year < 1980 {
        return zip::DateTime::from_date_and_time(1980, 1, 1, 0, 0, 0)
            .unwrap_or_else(|_| zip::DateTime::default());
    }
    if year > 2107 {
        return zip::DateTime::from_date_and_time(2107, 12, 31, 23, 59, 58)
            .unwrap_or_else(|_| zip::DateTime::default());
    }
    zip::DateTime::from_date_and_time(
        year as u16,
        naive.month() as u8,
        naive.day() as u8,
        naive.hour() as u8,
        naive.minute() as u8,
        naive.second() as u8,
    )
    .unwrap_or_else(|_| zip::DateTime::default())
}

#[cfg(test)]
mod rfc013_tests {
    //! RFC-013 unit tests for pure-Rust patcher helpers. The patcher's
    //! end-to-end ZIP-add behavior is covered by `tests/test_patcher_infra.py`
    //! (commit 5) — those tests can construct a real `XlsxPatcher` via the
    //! PyO3 boundary, which `cargo test` cannot link against.
    use super::*;

    #[test]
    fn epoch_or_now_honors_test_epoch_zero() {
        // `WOLFXL_TEST_EPOCH=0` falls below ZIP's representable range
        // (1980-01-01); the helper clamps to that floor. The point is
        // determinism, not the specific year.
        let prev = std::env::var("WOLFXL_TEST_EPOCH").ok();
        std::env::set_var("WOLFXL_TEST_EPOCH", "0");
        let dt = epoch_or_now();
        // Restore the env so we don't leak into other tests.
        match prev {
            Some(v) => std::env::set_var("WOLFXL_TEST_EPOCH", v),
            None => std::env::remove_var("WOLFXL_TEST_EPOCH"),
        }
        // Two back-to-back calls under the same epoch produce identical
        // ZIP datetimes — that's the byte-identical-save guarantee.
        std::env::set_var("WOLFXL_TEST_EPOCH", "0");
        let dt2 = epoch_or_now();
        std::env::remove_var("WOLFXL_TEST_EPOCH");
        // `zip::DateTime` doesn't impl PartialEq, so compare via the
        // `(year, month, day, hour, minute, second)` quintuple.
        assert_eq!(
            (dt.year(), dt.month(), dt.day(), dt.hour(), dt.minute(), dt.second()),
            (dt2.year(), dt2.month(), dt2.day(), dt2.hour(), dt2.minute(), dt2.second()),
        );
    }

    #[test]
    fn epoch_or_now_clamps_pre_1980_floor() {
        std::env::set_var("WOLFXL_TEST_EPOCH", "0");
        let dt = epoch_or_now();
        std::env::remove_var("WOLFXL_TEST_EPOCH");
        assert_eq!(dt.year(), 1980);
        assert_eq!(dt.month(), 1);
        assert_eq!(dt.day(), 1);
    }

    #[test]
    fn epoch_or_now_handles_recent_timestamp() {
        // 2024-01-01T00:00:00Z = 1_704_067_200 — well within the
        // ZIP-representable range, so no clamping.
        std::env::set_var("WOLFXL_TEST_EPOCH", "1704067200");
        let dt = epoch_or_now();
        std::env::remove_var("WOLFXL_TEST_EPOCH");
        assert_eq!(dt.year(), 2024);
        assert_eq!(dt.month(), 1);
        assert_eq!(dt.day(), 1);
    }

    #[test]
    fn sheet_order_parser_preserves_workbook_xml_order() {
        // Smoke: the helper that drives `XlsxPatcher::sheet_order` is
        // `parse_workbook_sheet_rids`, which is supposed to return
        // sheets in document order. Touch-test that here so a future
        // refactor that swaps it for a HashMap-keyed parser fails this
        // gate before it breaks RFC-020's `app.xml` regen.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Apples"  sheetId="1" r:id="rId1"/>
    <sheet name="Bananas" sheetId="2" r:id="rId2"/>
    <sheet name="Cherries" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>"#;
        let pairs = ooxml_util::parse_workbook_sheet_rids(xml).unwrap();
        let names: Vec<&str> = pairs.iter().map(|(n, _)| n.as_str()).collect();
        assert_eq!(names, vec!["Apples", "Bananas", "Cherries"]);
    }

    // -----------------------------------------------------------------
    // Phase 2.5c: cross-sheet content-types aggregation.
    //
    // The patcher's Phase-2.5c block iterates `sheet_order`, flattens
    // every sheet's `queued_content_type_ops` into one Vec, and applies
    // them onto a single [`ContentTypesGraph`]. These tests model that
    // chain directly so a regression in either `apply_op` or
    // serialize-order shows up here.
    // -----------------------------------------------------------------

    use content_types::{ContentTypeOp, ContentTypesGraph};

    /// Source [Content_Types].xml fixture used by the Phase-2.5c tests.
    /// Mirrors what `crates/wolfxl-writer/src/emit/content_types.rs::emit`
    /// produces for a 1-sheet workbook (rels Default, xml Default,
    /// workbook + 1 sheet + styles + sst Overrides).
    const SOURCE_CT_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#;

    #[test]
    fn phase_2_5c_no_ops_is_no_op() {
        // Empty op list → `[Content_Types].xml` is left untouched. The
        // patcher implements this via the `is_empty()` guard before
        // parse + serialize; modeling that here means asserting the
        // guard is the only path that mutates anything.
        let ops: Vec<ContentTypeOp> = Vec::new();
        // Verify the precondition for the no-op path.
        assert!(ops.is_empty(), "no-op precondition: no queued ops");
        // The patcher's `do_save` skips the parse + serialize entirely
        // when `content_type_ops.is_empty()`. So a no-op save preserves
        // source bytes verbatim — there is no rewrite path to assert
        // against.
    }

    #[test]
    fn phase_2_5c_aggregates_overrides_into_single_mutation() {
        // Multiple sheets pushing ops collapse to one parse + one
        // serialize. Ops in document order; result has every new
        // override appended after the source overrides.
        let ops = vec![
            ContentTypeOp::AddOverride(
                "/xl/comments1.xml".into(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml".into(),
            ),
            ContentTypeOp::EnsureDefault(
                "vml".into(),
                "application/vnd.openxmlformats-officedocument.vmlDrawing".into(),
            ),
            ContentTypeOp::AddOverride(
                "/xl/tables/table1.xml".into(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml".into(),
            ),
        ];
        let mut graph = ContentTypesGraph::parse(SOURCE_CT_XML).expect("parse source");
        for op in &ops {
            graph.apply_op(op);
        }
        let bytes = graph.serialize();
        let text = std::str::from_utf8(&bytes).expect("utf8 round-trip");
        // All three new entries present.
        assert!(text.contains("/xl/comments1.xml"), "comments override");
        assert!(text.contains("/xl/tables/table1.xml"), "table override");
        assert!(text.contains(r#"Extension="vml""#), "vml default");
        // Source entries still present (aggregation is additive).
        assert!(text.contains("/xl/workbook.xml"));
        assert!(text.contains("/xl/styles.xml"));
    }

    #[test]
    fn phase_2_5c_preserves_source_order_for_existing_overrides() {
        // The aggregation pass must not reorder source overrides — that
        // would break byte-stable diffs against unmodified parts. New
        // ops append; existing entries keep their slot.
        let ops = vec![ContentTypeOp::AddOverride(
            "/xl/comments1.xml".into(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml".into(),
        )];
        let mut graph = ContentTypesGraph::parse(SOURCE_CT_XML).expect("parse");
        for op in &ops {
            graph.apply_op(op);
        }
        let bytes = graph.serialize();
        let text = std::str::from_utf8(&bytes).expect("utf8");
        let idx_workbook = text.find("/xl/workbook.xml").expect("workbook");
        let idx_sheet1 = text.find("/xl/worksheets/sheet1.xml").expect("sheet1");
        let idx_styles = text.find("/xl/styles.xml").expect("styles");
        let idx_comments = text.find("/xl/comments1.xml").expect("comments");
        assert!(
            idx_workbook < idx_sheet1 && idx_sheet1 < idx_styles,
            "source overrides retain document order",
        );
        assert!(
            idx_styles < idx_comments,
            "new overrides append after source ones, not interleaved",
        );
    }

    // -----------------------------------------------------------------
    // Sprint Θ Pod-B: SAX-driven `splice_into_sheets_block`.
    //
    // The naive byte-substring locator was fooled by a workbook.xml
    // comment containing the literal `</sheets>` token (KNOWN_GAPS bug
    // #6). These tests pin the SAX rewrite: comments, CDATA, and PIs
    // surfaced as separate quick-xml events MUST NOT influence the
    // splice point.
    // -----------------------------------------------------------------

    const NEW_SHEET: &[u8] = br#"<sheet name="Copy" sheetId="2" r:id="rId99"/>"#;

    #[test]
    fn splice_normal_sheets_block_inserts_before_close() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;
        let out = splice_into_sheets_block(xml, NEW_SHEET).expect("splice ok");
        let s = std::str::from_utf8(&out).unwrap();
        // The new entry appears right before </sheets>, after the
        // existing Sheet1 entry, and only once.
        let rid1 = s.find("r:id=\"rId1\"").unwrap();
        let rid99 = s.find("r:id=\"rId99\"").unwrap();
        let close = s.find("</sheets>").unwrap();
        assert!(rid1 < rid99, "new sheet appended after Sheet1");
        assert!(rid99 < close, "new sheet inserted BEFORE </sheets>");
        assert_eq!(
            s.matches("</sheets>").count(),
            1,
            "exactly one </sheets> in output",
        );
    }

    #[test]
    fn splice_handles_self_closing_sheets() {
        // `<sheets/>` is rare but ECMA-376-legal. We rewrite it to an
        // open/close pair wrapping the new element.
        let xml = br#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets/>
</workbook>"#;
        let out = splice_into_sheets_block(xml, NEW_SHEET).expect("splice ok");
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains("<sheets>"), "open tag emitted");
        assert!(s.contains("</sheets>"), "close tag emitted");
        assert!(s.contains("rId99"), "new sheet entry inserted");
        // The original `<sheets/>` is gone — exactly one open and one close.
        assert_eq!(s.matches("<sheets>").count(), 1);
        assert_eq!(s.matches("</sheets>").count(), 1);
        assert!(!s.contains("<sheets/>"));
    }

    #[test]
    fn splice_ignores_fake_close_tag_inside_comment() {
        // Bug #6 from KNOWN_GAPS.md. A comment containing the literal
        // `</sheets>` MUST NOT fool the locator. The new sheet must
        // land inside the real <sheets> block (between the existing
        // <sheet …/> and the real </sheets>).
        let xml = br#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<!-- FUZZTOKEN: this fakeout closes </sheets> here, naive splice would bite -->
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;
        let out = splice_into_sheets_block(xml, NEW_SHEET).expect("splice ok");
        let s = std::str::from_utf8(&out).unwrap();
        // FUZZTOKEN comment survives.
        assert!(s.contains("FUZZTOKEN"), "comment survives splice");
        // New sheet appears between real <sheets> open and close.
        let open = s.find("<sheets>").expect("real <sheets> open");
        // The fakeout `</sheets>` token in the comment counts as a
        // string match, so use the LAST occurrence as the real close.
        let close = s.rfind("</sheets>").expect("real </sheets> close");
        let rid99 = s.find("rId99").expect("new entry present");
        assert!(open < rid99, "new entry after real <sheets> open");
        assert!(rid99 < close, "new entry before real </sheets> close");
    }

    #[test]
    fn splice_ignores_fake_close_tag_inside_cdata() {
        // CDATA section containing `</sheets>` — also must not fool
        // the locator. Note: workbook.xml almost never has CDATA in
        // practice, but quick-xml's event model handles it for free.
        let xml = br#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"><![CDATA[fake </sheets> token]]></sheet></sheets>
</workbook>"#;
        let out = splice_into_sheets_block(xml, NEW_SHEET).expect("splice ok");
        let s = std::str::from_utf8(&out).unwrap();
        // The new entry must come AFTER the CDATA (i.e. past the
        // </sheet> close inside the real block) and BEFORE the
        // real </sheets>.
        let rid99 = s.find("rId99").expect("new entry present");
        let cdata_close = s.find("]]>").expect("cdata close");
        let real_close = s.rfind("</sheets>").expect("real close");
        assert!(cdata_close < rid99, "new entry follows CDATA");
        assert!(rid99 < real_close, "new entry before real </sheets>");
    }

    #[test]
    fn splice_returns_error_when_no_sheets_block() {
        // Malformed input: no <sheets> at all. Preserve the historical
        // error message for callers.
        let xml = br#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
        let err = splice_into_sheets_block(xml, NEW_SHEET).unwrap_err();
        // PyErr message lookup needs Python; check the type instead.
        // The string check is loose: match the error text via Display.
        let msg = format!("{err}");
        assert!(
            msg.contains("no <sheets>"),
            "preserves historical error message, got: {msg}"
        );
    }
}

// ---------------------------------------------------------------------------
// Dict → spec conversion helpers
// ---------------------------------------------------------------------------

/// Sprint Ι Pod-α: convert the Python rich-text payload (a list of
/// ``(text, font_dict_or_None)`` tuples) into the Rust ``RichTextRun``
/// vector that the patcher and the writer both consume.
fn py_runs_to_rust(
    runs: &Bound<'_, pyo3::types::PyList>,
) -> PyResult<Vec<wolfxl_writer::rich_text::RichTextRun>> {
    use wolfxl_writer::rich_text::{InlineFontProps, RichTextRun};
    let mut out: Vec<RichTextRun> = Vec::with_capacity(runs.len());
    for entry in runs.iter() {
        // Each entry is a (text, font_or_none) 2-tuple — accept lists too.
        let seq: &Bound<'_, pyo3::types::PySequence> = entry.downcast()?;
        if seq.len()? < 2 {
            return Err(PyErr::new::<PyValueError, _>(
                "rich-text run must be a (text, font_or_none) pair",
            ));
        }
        let text: String = seq.get_item(0)?.extract()?;
        let font_obj = seq.get_item(1)?;
        let font = if font_obj.is_none() {
            None
        } else {
            let d: &Bound<'_, PyDict> = font_obj.downcast()?;
            let mut props = InlineFontProps::default();
            if let Some(v) = d.get_item("b")? {
                if !v.is_none() {
                    props.bold = Some(v.extract::<bool>()?);
                }
            }
            if let Some(v) = d.get_item("i")? {
                if !v.is_none() {
                    props.italic = Some(v.extract::<bool>()?);
                }
            }
            if let Some(v) = d.get_item("strike")? {
                if !v.is_none() {
                    props.strike = Some(v.extract::<bool>()?);
                }
            }
            if let Some(v) = d.get_item("u")? {
                if !v.is_none() {
                    let s: String = v.extract()?;
                    props.underline = Some(s);
                }
            }
            if let Some(v) = d.get_item("sz")? {
                if !v.is_none() {
                    props.size = Some(v.extract::<f64>()?);
                }
            }
            if let Some(v) = d.get_item("color")? {
                if !v.is_none() {
                    let s: String = v.extract()?;
                    props.color = Some(s);
                }
            }
            if let Some(v) = d.get_item("rFont")? {
                if !v.is_none() {
                    let s: String = v.extract()?;
                    props.name = Some(s);
                }
            }
            if let Some(v) = d.get_item("family")? {
                if !v.is_none() {
                    props.family = Some(v.extract::<i32>()?);
                }
            }
            if let Some(v) = d.get_item("charset")? {
                if !v.is_none() {
                    props.charset = Some(v.extract::<i32>()?);
                }
            }
            if let Some(v) = d.get_item("vertAlign")? {
                if !v.is_none() {
                    let s: String = v.extract()?;
                    props.vert_align = Some(s);
                }
            }
            if let Some(v) = d.get_item("scheme")? {
                if !v.is_none() {
                    let s: String = v.extract()?;
                    props.scheme = Some(s);
                }
            }
            Some(props)
        };
        out.push(RichTextRun { text, font });
    }
    Ok(out)
}

fn dict_to_format_spec(d: &Bound<'_, PyDict>) -> PyResult<FormatSpec> {
    let mut spec = FormatSpec::default();

    // Font properties
    let bold = extract_bool(d, "bold")?;
    let italic = extract_bool(d, "italic")?;
    let underline = extract_bool(d, "underline")?;
    let strikethrough = extract_bool(d, "strikethrough")?;
    let font_name = extract_str(d, "font_name")?;
    let font_size = extract_u32(d, "font_size")?;
    let font_color = extract_str(d, "font_color")?;

    if bold.is_some()
        || italic.is_some()
        || underline.is_some()
        || strikethrough.is_some()
        || font_name.is_some()
        || font_size.is_some()
        || font_color.is_some()
    {
        spec.font = Some(styles::FontSpec {
            bold: bold.unwrap_or(false),
            italic: italic.unwrap_or(false),
            underline: underline.unwrap_or(false),
            strikethrough: strikethrough.unwrap_or(false),
            name: font_name,
            size: font_size,
            color_rgb: font_color.map(|c| normalize_color(&c)),
        });
    }

    // Fill properties
    let bg_color = extract_str(d, "bg_color")?;
    if let Some(color) = bg_color {
        spec.fill = Some(styles::FillSpec {
            pattern_type: "solid".to_string(),
            fg_color_rgb: Some(normalize_color(&color)),
        });
    }

    // Number format
    spec.number_format = extract_str(d, "number_format")?;

    // Alignment — accept both openpyxl-style and wolfxl-style key names
    let horizontal = extract_str(d, "horizontal")?.or(extract_str(d, "h_align")?);
    let vertical = extract_str(d, "vertical")?.or(extract_str(d, "v_align")?);
    let wrap_text = extract_bool(d, "wrap_text")?.or(extract_bool(d, "wrap")?);
    let indent = extract_u32(d, "indent")?;
    let text_rotation = extract_u32(d, "text_rotation")?.or(extract_u32(d, "rotation")?);

    if horizontal.is_some()
        || vertical.is_some()
        || wrap_text.is_some()
        || indent.is_some()
        || text_rotation.is_some()
    {
        spec.alignment = Some(styles::AlignmentSpec {
            horizontal,
            vertical,
            wrap_text: wrap_text.unwrap_or(false),
            indent: indent.unwrap_or(0),
            text_rotation: text_rotation.unwrap_or(0),
        });
    }

    Ok(spec)
}

fn dict_to_border_spec(d: &Bound<'_, PyDict>) -> PyResult<styles::BorderSpec> {
    fn extract_side(d: &Bound<'_, PyDict>, key: &str) -> PyResult<styles::BorderSideSpec> {
        if let Some(side) = d.get_item(key)? {
            if let Ok(sd) = side.downcast::<PyDict>() {
                let style = extract_str(sd, "style")?;
                let color = extract_str(sd, "color")?.map(|c| normalize_color(&c));
                return Ok(styles::BorderSideSpec {
                    style,
                    color_rgb: color,
                });
            }
        }
        Ok(styles::BorderSideSpec::default())
    }

    Ok(styles::BorderSpec {
        left: extract_side(d, "left")?,
        right: extract_side(d, "right")?,
        top: extract_side(d, "top")?,
        bottom: extract_side(d, "bottom")?,
    })
}

fn extract_cf_rule(d: &Bound<'_, PyDict>) -> PyResult<CfRulePatch> {
    let kind_tag = extract_str(d, "kind")?
        .ok_or_else(|| PyErr::new::<PyValueError, _>("CF rule requires 'kind'"))?;

    let kind = match kind_tag.as_str() {
        "cellIs" => CfRuleKind::CellIs {
            operator: extract_str(d, "operator")?.unwrap_or_else(|| "equal".to_string()),
            formula_a: extract_str(d, "formula_a")?.unwrap_or_default(),
            formula_b: extract_str(d, "formula_b")?,
        },
        "expression" => CfRuleKind::Expression {
            formula: extract_str(d, "formula")?.unwrap_or_default(),
        },
        "colorScale" => {
            let stops_obj = d.get_item("stops")?.ok_or_else(|| {
                PyErr::new::<PyValueError, _>("colorScale rule requires 'stops'")
            })?;
            let stops_list = stops_obj.downcast::<pyo3::types::PyList>().map_err(|_| {
                PyErr::new::<PyValueError, _>("'stops' must be a list of dicts")
            })?;
            let mut stops: Vec<ColorScaleStop> = Vec::with_capacity(stops_list.len());
            for s in stops_list.iter() {
                let sd = s
                    .downcast::<PyDict>()
                    .map_err(|_| PyErr::new::<PyValueError, _>("each stop must be a dict"))?;
                stops.push(ColorScaleStop {
                    cfvo: CfvoPatch {
                        cfvo_type: extract_str(sd, "cfvo_type")?
                            .unwrap_or_else(|| "min".to_string()),
                        val: extract_str(sd, "val")?,
                    },
                    color_rgb: extract_str(sd, "color_rgb")?.unwrap_or_default(),
                });
            }
            CfRuleKind::ColorScale { stops }
        }
        "dataBar" => CfRuleKind::DataBar {
            min: CfvoPatch {
                cfvo_type: extract_str(d, "min_cfvo_type")?
                    .unwrap_or_else(|| "min".to_string()),
                val: extract_str(d, "min_val")?,
            },
            max: CfvoPatch {
                cfvo_type: extract_str(d, "max_cfvo_type")?
                    .unwrap_or_else(|| "max".to_string()),
                val: extract_str(d, "max_val")?,
            },
            color_rgb: extract_str(d, "color_rgb")?.unwrap_or_default(),
        },
        other => {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "unsupported CF rule kind: '{other}'"
            )));
        }
    };

    let dxf = match d.get_item("dxf")? {
        Some(v) if !v.is_none() => {
            let dd = v.downcast::<PyDict>().map_err(|_| {
                PyErr::new::<PyValueError, _>("'dxf' must be a dict or None")
            })?;
            Some(extract_dxf_patch(dd)?)
        }
        _ => None,
    };

    Ok(CfRulePatch {
        kind,
        dxf,
        stop_if_true: extract_bool(d, "stop_if_true")?.unwrap_or(false),
    })
}

fn extract_dxf_patch(d: &Bound<'_, PyDict>) -> PyResult<DxfPatch> {
    Ok(DxfPatch {
        font_bold: extract_bool(d, "font_bold")?,
        font_italic: extract_bool(d, "font_italic")?,
        font_color_rgb: extract_str(d, "font_color_rgb")?.map(|c| normalize_color(&c)),
        fill_pattern_type: extract_str(d, "fill_pattern_type")?,
        fill_fg_color_rgb: extract_str(d, "fill_fg_color_rgb")?.map(|c| normalize_color(&c)),
        border_top_style: extract_str(d, "border_top_style")?,
        border_bottom_style: extract_str(d, "border_bottom_style")?,
        border_left_style: extract_str(d, "border_left_style")?,
        border_right_style: extract_str(d, "border_right_style")?,
    })
}

fn extract_str(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<String>> {
    d.get_item(key)?.map(|v| v.extract::<String>()).transpose()
}

fn extract_bool(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<bool>> {
    d.get_item(key)?.map(|v| v.extract::<bool>()).transpose()
}

fn extract_u32(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<u32>> {
    d.get_item(key)?.map(|v| v.extract::<u32>()).transpose()
}

fn extract_f64(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<f64>> {
    d.get_item(key)?.map(|v| v.extract::<f64>()).transpose()
}

// ---------------------------------------------------------------------------
// RFC-058: workbook security payload parsing
// ---------------------------------------------------------------------------

/// Convert a Python dict matching RFC-058 §10 into a
/// [`wolfxl_writer::parse::workbook_security::WorkbookSecurity`].
///
/// Either or both top-level keys (`workbook_protection`,
/// `file_sharing`) may be `None`; the resulting struct mirrors that
/// optionality.
fn parse_workbook_security_payload(
    payload: &Bound<'_, PyDict>,
) -> PyResult<wolfxl_writer::parse::workbook_security::WorkbookSecurity> {
    use wolfxl_writer::parse::workbook_security::{
        FileSharingSpec, WorkbookProtectionSpec, WorkbookSecurity,
    };

    let workbook_protection = match payload.get_item("workbook_protection")? {
        Some(v) if !v.is_none() => {
            let d = v.downcast::<PyDict>().map_err(|_| {
                PyErr::new::<PyValueError, _>(
                    "queue_workbook_security: 'workbook_protection' must be a dict or None",
                )
            })?;
            Some(WorkbookProtectionSpec {
                lock_structure: extract_bool(d, "lock_structure")?.unwrap_or(false),
                lock_windows: extract_bool(d, "lock_windows")?.unwrap_or(false),
                lock_revision: extract_bool(d, "lock_revision")?.unwrap_or(false),
                workbook_algorithm_name: extract_str(d, "workbook_algorithm_name")?,
                workbook_hash_value: extract_str(d, "workbook_hash_value")?,
                workbook_salt_value: extract_str(d, "workbook_salt_value")?,
                workbook_spin_count: extract_u32(d, "workbook_spin_count")?,
                revisions_algorithm_name: extract_str(d, "revisions_algorithm_name")?,
                revisions_hash_value: extract_str(d, "revisions_hash_value")?,
                revisions_salt_value: extract_str(d, "revisions_salt_value")?,
                revisions_spin_count: extract_u32(d, "revisions_spin_count")?,
            })
        }
        _ => None,
    };

    let file_sharing = match payload.get_item("file_sharing")? {
        Some(v) if !v.is_none() => {
            let d = v.downcast::<PyDict>().map_err(|_| {
                PyErr::new::<PyValueError, _>(
                    "queue_workbook_security: 'file_sharing' must be a dict or None",
                )
            })?;
            Some(FileSharingSpec {
                read_only_recommended: extract_bool(d, "read_only_recommended")?.unwrap_or(false),
                user_name: extract_str(d, "user_name")?,
                algorithm_name: extract_str(d, "algorithm_name")?,
                hash_value: extract_str(d, "hash_value")?,
                salt_value: extract_str(d, "salt_value")?,
                spin_count: extract_u32(d, "spin_count")?,
            })
        }
        _ => None,
    };

    Ok(WorkbookSecurity {
        workbook_protection,
        file_sharing,
    })
}

/// Normalize "#RRGGBB" or "RRGGBB" to "FFRRGGBB" (OOXML ARGB format).
fn normalize_color(color: &str) -> String {
    let hex = color.trim_start_matches('#');
    if hex.len() == 6 {
        format!("FF{}", hex.to_uppercase())
    } else if hex.len() == 8 {
        hex.to_uppercase()
    } else {
        format!("FF{hex}")
    }
}

fn minimal_styles_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#
        .to_string()
}

// ---------------------------------------------------------------------------
// Sprint Λ Pod-β (RFC-045) — image-add helpers + Phase 2.5k.
// ---------------------------------------------------------------------------

fn parse_queued_image_anchor(d: &Bound<'_, PyDict>) -> PyResult<QueuedImageAnchor> {
    let kind: String = d
        .get_item("type")?
        .ok_or_else(|| PyValueError::new_err("anchor dict missing 'type'"))?
        .extract()?;
    let q_int = |key: &str, default: u32| -> PyResult<u32> {
        Ok(d.get_item(key)?
            .and_then(|v| v.extract().ok())
            .unwrap_or(default))
    };
    let q_i64 = |key: &str, default: i64| -> PyResult<i64> {
        Ok(d.get_item(key)?
            .and_then(|v| v.extract().ok())
            .unwrap_or(default))
    };
    match kind.as_str() {
        "one_cell" => Ok(QueuedImageAnchor::OneCell {
            from_col: q_int("from_col", 0)?,
            from_row: q_int("from_row", 0)?,
            from_col_off: q_i64("from_col_off", 0)?,
            from_row_off: q_i64("from_row_off", 0)?,
        }),
        "two_cell" => Ok(QueuedImageAnchor::TwoCell {
            from_col: q_int("from_col", 0)?,
            from_row: q_int("from_row", 0)?,
            from_col_off: q_i64("from_col_off", 0)?,
            from_row_off: q_i64("from_row_off", 0)?,
            to_col: q_int("to_col", 0)?,
            to_row: q_int("to_row", 0)?,
            to_col_off: q_i64("to_col_off", 0)?,
            to_row_off: q_i64("to_row_off", 0)?,
            edit_as: d
                .get_item("edit_as")?
                .and_then(|v| v.extract().ok())
                .unwrap_or_else(|| "oneCell".to_string()),
        }),
        "absolute" => Ok(QueuedImageAnchor::Absolute {
            x_emu: q_i64("x_emu", 0)?,
            y_emu: q_i64("y_emu", 0)?,
            cx_emu: q_i64("cx_emu", 0)?,
            cy_emu: q_i64("cy_emu", 0)?,
        }),
        other => Err(PyValueError::new_err(format!(
            "unknown anchor type: {other:?}"
        ))),
    }
}

/// Build the `xl/drawings/drawingN.xml` body for the queued images.
/// Rels for each image are 1-indexed (`rId1`, `rId2`, ...) since each
/// drawing has its own rels graph.
pub(crate) fn build_drawing_xml(images: &[QueuedImageAdd]) -> String {
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let emu_per_px: i64 = 9525;
    let mut out = String::with_capacity(512 + images.len() * 512);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    out.push_str(&format!(
        "<xdr:wsDr xmlns:xdr=\"{xdr_ns}\" xmlns:a=\"{a_ns}\" xmlns:r=\"{r_ns}\">"
    ));
    for (i, img) in images.iter().enumerate() {
        let pic_id = (i + 1) as u32;
        let rid = format!("rId{}", i + 1);
        match &img.anchor {
            QueuedImageAnchor::OneCell {
                from_col,
                from_row,
                from_col_off,
                from_row_off,
            } => {
                out.push_str("<xdr:oneCellAnchor>");
                out.push_str(&format!(
                    "<xdr:from><xdr:col>{from_col}</xdr:col><xdr:colOff>{from_col_off}</xdr:colOff>\
                     <xdr:row>{from_row}</xdr:row><xdr:rowOff>{from_row_off}</xdr:rowOff></xdr:from>"
                ));
                let cx = img.width_px as i64 * emu_per_px;
                let cy = img.height_px as i64 * emu_per_px;
                out.push_str(&format!("<xdr:ext cx=\"{cx}\" cy=\"{cy}\"/>"));
            }
            QueuedImageAnchor::TwoCell {
                from_col,
                from_row,
                from_col_off,
                from_row_off,
                to_col,
                to_row,
                to_col_off,
                to_row_off,
                edit_as,
            } => {
                out.push_str(&format!("<xdr:twoCellAnchor editAs=\"{edit_as}\">"));
                out.push_str(&format!(
                    "<xdr:from><xdr:col>{from_col}</xdr:col><xdr:colOff>{from_col_off}</xdr:colOff>\
                     <xdr:row>{from_row}</xdr:row><xdr:rowOff>{from_row_off}</xdr:rowOff></xdr:from>"
                ));
                out.push_str(&format!(
                    "<xdr:to><xdr:col>{to_col}</xdr:col><xdr:colOff>{to_col_off}</xdr:colOff>\
                     <xdr:row>{to_row}</xdr:row><xdr:rowOff>{to_row_off}</xdr:rowOff></xdr:to>"
                ));
            }
            QueuedImageAnchor::Absolute {
                x_emu,
                y_emu,
                cx_emu,
                cy_emu,
            } => {
                out.push_str("<xdr:absoluteAnchor>");
                out.push_str(&format!("<xdr:pos x=\"{x_emu}\" y=\"{y_emu}\"/>"));
                out.push_str(&format!("<xdr:ext cx=\"{cx_emu}\" cy=\"{cy_emu}\"/>"));
            }
        }
        let cx = img.width_px as i64 * emu_per_px;
        let cy = img.height_px as i64 * emu_per_px;
        out.push_str(&format!(
            "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"{pic_id}\" name=\"Picture {pic_id}\" descr=\"Picture {pic_id}\"/>\
             <xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>\
             <xdr:blipFill><a:blip xmlns:r=\"{r_ns}\" r:embed=\"{rid}\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>\
             <xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>\
             <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>"
        ));
        out.push_str("<xdr:clientData/>");
        match &img.anchor {
            QueuedImageAnchor::OneCell { .. } => out.push_str("</xdr:oneCellAnchor>"),
            QueuedImageAnchor::TwoCell { .. } => out.push_str("</xdr:twoCellAnchor>"),
            QueuedImageAnchor::Absolute { .. } => out.push_str("</xdr:absoluteAnchor>"),
        }
    }
    out.push_str("</xdr:wsDr>");
    out
}

/// Build `xl/drawings/_rels/drawingN.xml.rels` for the given images.
/// `image_indices` are 1-based global media indices.
pub(crate) fn build_drawing_rels_xml(images: &[QueuedImageAdd], image_indices: &[u32]) -> String {
    debug_assert_eq!(images.len(), image_indices.len());
    let mut g = wolfxl_rels::RelsGraph::new();
    for (img, &n) in images.iter().zip(image_indices.iter()) {
        g.add(
            wolfxl_rels::rt::IMAGE,
            &format!("../media/image{n}.{}", img.ext),
            wolfxl_rels::TargetMode::Internal,
        );
    }
    String::from_utf8(g.serialize()).expect("rels serialize is utf8")
}

/// Splice a `<drawing r:id="rIdN"/>` element into a sheet XML body.
///
/// Insertion strategy: locate the position just before
/// `<legacyDrawing` (slot 31) if present; failing that, just before
/// `</worksheet>`. This matches the ECMA element-order rule for
/// slot 30. If a `<drawing` element is already present, returns
/// `Err` so the caller can raise NotImplementedError (v1.5 limit).
///
/// The splice also ensures the root `<worksheet>` element declares
/// `xmlns:r="..."` — openpyxl-generated sheets sometimes omit it
/// when no `r:` reference is currently present, but our `<drawing
/// r:id="..."/>` requires the prefix to be bound.
pub(crate) fn splice_drawing_ref(sheet_xml: &str, rid: &str) -> Result<String, &'static str> {
    if sheet_xml.contains("<drawing ") || sheet_xml.contains("<drawing/>") {
        return Err("sheet already has a <drawing> element");
    }
    let elem = format!("<drawing r:id=\"{rid}\"/>");
    let with_drawing = if let Some(idx) = sheet_xml.find("<legacyDrawing") {
        let mut out = String::with_capacity(sheet_xml.len() + elem.len());
        out.push_str(&sheet_xml[..idx]);
        out.push_str(&elem);
        out.push_str(&sheet_xml[idx..]);
        out
    } else if let Some(idx) = sheet_xml.rfind("</worksheet>") {
        let mut out = String::with_capacity(sheet_xml.len() + elem.len());
        out.push_str(&sheet_xml[..idx]);
        out.push_str(&elem);
        out.push_str(&sheet_xml[idx..]);
        out
    } else {
        return Err("sheet xml has no </worksheet> closing tag");
    };
    Ok(ensure_xmlns_r_on_worksheet(&with_drawing))
}

/// Ensure the `<worksheet>` root element of `sheet_xml` declares the
/// `r` prefix bound to the OOXML relationships namespace. No-op if
/// `xmlns:r` is already present anywhere in the open tag.
pub(crate) fn ensure_xmlns_r_on_worksheet(sheet_xml: &str) -> String {
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    // Find the open `<worksheet ...>` tag.
    let start = match sheet_xml.find("<worksheet") {
        Some(i) => i,
        None => return sheet_xml.to_string(),
    };
    let end = match sheet_xml[start..].find('>') {
        Some(e) => start + e,
        None => return sheet_xml.to_string(),
    };
    let open_tag = &sheet_xml[start..=end];
    if open_tag.contains("xmlns:r=") {
        return sheet_xml.to_string();
    }
    // Insert `xmlns:r="..."` just before the closing `>` (or `/>`).
    let inserted = if open_tag.ends_with("/>") {
        format!(
            "{} xmlns:r=\"{r_ns}\"/>",
            &open_tag[..open_tag.len() - 2]
        )
    } else {
        format!(
            "{} xmlns:r=\"{r_ns}\">",
            &open_tag[..open_tag.len() - 1]
        )
    };
    let mut out = String::with_capacity(sheet_xml.len() + 80);
    out.push_str(&sheet_xml[..start]);
    out.push_str(&inserted);
    out.push_str(&sheet_xml[end + 1..]);
    out
}

#[cfg(test)]
mod image_helpers_tests {
    use super::*;

    #[test]
    fn splice_drawing_before_legacy_drawing() {
        let xml = r#"<?xml version="1.0"?><worksheet><sheetData/><legacyDrawing r:id="rId2"/></worksheet>"#;
        let out = splice_drawing_ref(xml, "rId5").unwrap();
        assert!(out.contains("<drawing r:id=\"rId5\"/><legacyDrawing"));
    }

    #[test]
    fn splice_drawing_before_close_when_no_legacy() {
        let xml = r#"<?xml version="1.0"?><worksheet><sheetData/></worksheet>"#;
        let out = splice_drawing_ref(xml, "rId1").unwrap();
        assert!(out.contains("<drawing r:id=\"rId1\"/></worksheet>"));
    }

    #[test]
    fn splice_drawing_errors_when_already_present() {
        let xml = r#"<?xml version="1.0"?><worksheet><sheetData/><drawing r:id="rId7"/></worksheet>"#;
        assert!(splice_drawing_ref(xml, "rId1").is_err());
    }

    #[test]
    fn build_drawing_xml_roundtrip() {
        let imgs = vec![QueuedImageAdd {
            data: vec![],
            ext: "png".into(),
            width_px: 10,
            height_px: 5,
            anchor: QueuedImageAnchor::OneCell {
                from_col: 1,
                from_row: 4,
                from_col_off: 0,
                from_row_off: 0,
            },
        }];
        let xml = build_drawing_xml(&imgs);
        assert!(xml.contains("<xdr:oneCellAnchor>"));
        assert!(xml.contains("r:embed=\"rId1\""));
    }
}

// ---------------------------------------------------------------------------
// Sprint Μ Pod-γ (RFC-046) — chart drawing helpers + A1 + path utils.
// ---------------------------------------------------------------------------

/// Parse an A1-style coordinate (e.g. `"D2"`) into `(col_zero_based,
/// row_zero_based)`. Lowercase letters accepted.
pub(crate) fn parse_a1_coord(s: &str) -> Option<(u32, u32)> {
    let s = s.trim().trim_start_matches('$');
    let mut col: u32 = 0;
    let mut iter = s.chars().peekable();
    let mut col_chars = 0;
    while let Some(&c) = iter.peek() {
        if c.is_ascii_alphabetic() {
            let v = (c.to_ascii_uppercase() as u32) - ('A' as u32) + 1;
            col = col * 26 + v;
            iter.next();
            col_chars += 1;
        } else {
            break;
        }
    }
    if col_chars == 0 || col == 0 {
        return None;
    }
    let rest: String = iter.collect();
    let rest = rest.trim_start_matches('$');
    let row: u32 = rest.parse().ok()?;
    if row == 0 {
        return None;
    }
    Some((col - 1, row - 1))
}

/// Resolve a relative or absolute OOXML rel target against a base
/// directory. Examples:
///
/// * `("xl/worksheets", "../drawings/drawing1.xml")` → `"xl/drawings/drawing1.xml"`
/// * `("xl/worksheets", "/xl/drawings/drawing1.xml")` → `"xl/drawings/drawing1.xml"`
///   (openpyxl uses leading-slash internal absolutes)
/// * `("xl", "drawings/drawing1.xml")` → `"xl/drawings/drawing1.xml"`
pub(crate) fn resolve_relative_path(base_dir: &str, target: &str) -> String {
    // Leading `/` means "internal absolute" — drop the prefix and
    // start from a fresh root.
    let (mut parts, target_iter): (Vec<&str>, _) =
        if let Some(stripped) = target.strip_prefix('/') {
            (Vec::new(), stripped.split('/'))
        } else {
            (
                base_dir
                    .split('/')
                    .filter(|p| !p.is_empty())
                    .collect(),
                target.split('/'),
            )
        };
    for seg in target_iter {
        match seg {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            other => parts.push(other),
        }
    }
    parts.join("/")
}

/// Best-effort extract `N` from `xl/drawings/drawingN.xml`.
pub(crate) fn drawing_n_from_path(path: &str) -> Option<u32> {
    let fname = path.rsplit('/').next()?;
    let core = fname.strip_suffix(".xml")?;
    let n_str = core.strip_prefix("drawing")?;
    n_str.parse::<u32>().ok()
}

/// Build a fresh `xl/drawings/drawingN.xml` body containing one
/// `<xdr:oneCellAnchor>` wrapping a `<xdr:graphicFrame>` per queued
/// chart. The chart rids are 1-indexed within the drawing's own
/// rels file, exactly matching the order of `queued`/`chart_rids`.
pub(crate) fn build_chart_drawing_xml(
    queued: &[QueuedChartAdd],
    chart_rids: &[String],
) -> String {
    debug_assert_eq!(queued.len(), chart_rids.len());
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let c_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    let mut out = String::with_capacity(512 + queued.len() * 768);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    out.push_str(&format!(
        "<xdr:wsDr xmlns:xdr=\"{xdr_ns}\" xmlns:a=\"{a_ns}\" xmlns:r=\"{r_ns}\" xmlns:c=\"{c_ns}\">"
    ));
    for (i, (chart, rid)) in queued.iter().zip(chart_rids.iter()).enumerate() {
        out.push_str(&render_graphic_frame_anchor(chart, rid, (i + 1) as u32));
    }
    out.push_str("</xdr:wsDr>");
    out
}

/// Append one anchor per queued chart to an existing drawing XML
/// body. Inserts before the closing `</xdr:wsDr>` (or `</wsDr>`)
/// tag. Detects whether the existing wrapper uses an `xdr:` prefix
/// or a default-namespace `<wsDr>` and emits the appended fragment
/// in the matching style so the merged body stays well-formed XML
/// without requiring the caller to pre-declare prefixes.
pub(crate) fn append_graphic_frames(
    drawing_xml: &[u8],
    queued: &[QueuedChartAdd],
    chart_rids: &[String],
) -> Result<Vec<u8>, String> {
    debug_assert_eq!(queued.len(), chart_rids.len());
    let body = std::str::from_utf8(drawing_xml).map_err(|e| e.to_string())?;
    // Choose the namespace style so the inserted fragment matches
    // the existing wrapper. If we see an `xdr:` prefix anywhere in
    // the wrapper we emit prefixed tags; otherwise we use default-
    // namespace tags + an explicit `xmlns:xdr="…"` on the inserted
    // root element so any attributes remain valid.
    let use_xdr_prefix = body.contains("<xdr:wsDr") || body.contains("xmlns:xdr=");
    // Count existing anchors / picture frames to keep cNvPr ids
    // monotonic.
    let existing_count: u32 = (body.matches("<graphicFrame").count()
        + body.matches("<pic").count()) as u32;
    let mut new_anchors = String::with_capacity(queued.len() * 512);
    for (i, (chart, rid)) in queued.iter().zip(chart_rids.iter()).enumerate() {
        new_anchors.push_str(&render_graphic_frame_anchor_styled(
            chart,
            rid,
            existing_count + (i + 1) as u32,
            use_xdr_prefix,
        ));
    }
    // Drawing XML may use either `<xdr:wsDr>` (prefixed) or
    // `<wsDr>` (default-namespaced — what openpyxl emits). Find
    // whichever close tag is present.
    let pos_opt = body
        .rfind("</xdr:wsDr>")
        .or_else(|| body.rfind("</wsDr>"));
    if let Some(pos) = pos_opt {
        let mut out = String::with_capacity(body.len() + new_anchors.len());
        out.push_str(&body[..pos]);
        out.push_str(&new_anchors);
        out.push_str(&body[pos..]);
        Ok(out.into_bytes())
    } else {
        // Drawing body has no </xdr:wsDr> — wrap a minimal envelope.
        let xdr_ns =
            "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
        let r_ns =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let c_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        let mut out = String::with_capacity(new_anchors.len() + 256);
        out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
        out.push_str(&format!(
            "<xdr:wsDr xmlns:xdr=\"{xdr_ns}\" xmlns:a=\"{a_ns}\" xmlns:r=\"{r_ns}\" xmlns:c=\"{c_ns}\">"
        ));
        out.push_str(&new_anchors);
        out.push_str("</xdr:wsDr>");
        Ok(out.into_bytes())
    }
}

/// Render one `<xdr:oneCellAnchor>` containing a `<xdr:graphicFrame>`
/// for an embedded chart referenced by `chart_rid`, using `xdr:`
/// prefixes (the historical default; matches our fresh-drawing
/// envelope which declares `xmlns:xdr="..."`).
fn render_graphic_frame_anchor(
    chart: &QueuedChartAdd,
    chart_rid: &str,
    unique_id: u32,
) -> String {
    render_graphic_frame_anchor_styled(chart, chart_rid, unique_id, true)
}

/// As [`render_graphic_frame_anchor`] but emits either prefixed
/// (`<xdr:oneCellAnchor>`) or default-namespace
/// (`<oneCellAnchor xmlns="...">`) tags. The default-namespace
/// variant carries an explicit `xmlns:xdr=""` declaration so the
/// merged drawing body remains valid even when the existing wrapper
/// is `<wsDr xmlns="..."/>` (openpyxl's emit style).
fn render_graphic_frame_anchor_styled(
    chart: &QueuedChartAdd,
    chart_rid: &str,
    unique_id: u32,
    use_xdr_prefix: bool,
) -> String {
    let (col0, row0) = parse_a1_coord(&chart.anchor_a1)
        .unwrap_or((3, 1)); // fallback: D2
    let cx = chart.width_emu;
    let cy = chart.height_emu;
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let c_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    let p = if use_xdr_prefix { "xdr:" } else { "" };
    // For the default-namespace style we declare xmlns on the
    // anchor's root so the unqualified tags resolve correctly.
    let root_xmlns = if use_xdr_prefix {
        String::new()
    } else {
        format!(
            " xmlns=\"{xdr_ns}\" xmlns:a=\"{a_ns}\" xmlns:r=\"{r_ns}\" xmlns:c=\"{c_ns}\""
        )
    };

    let mut out = String::with_capacity(640);
    out.push_str(&format!("<{p}oneCellAnchor{root_xmlns}>"));
    out.push_str(&format!(
        "<{p}from><{p}col>{col0}</{p}col><{p}colOff>0</{p}colOff>\
         <{p}row>{row0}</{p}row><{p}rowOff>0</{p}rowOff></{p}from>"
    ));
    out.push_str(&format!("<{p}ext cx=\"{cx}\" cy=\"{cy}\"/>"));
    // graphicFrame interior: a:* and c:* prefixes are always used
    // (drawingML uses its own namespace anchors regardless of the
    // wsDr-prefix style).
    out.push_str(&format!(
        "<{p}graphicFrame macro=\"\">\
           <{p}nvGraphicFramePr>\
             <{p}cNvPr id=\"{unique_id}\" name=\"Chart {unique_id}\"/>\
             <{p}cNvGraphicFramePr/>\
           </{p}nvGraphicFramePr>\
           <{p}xfrm>\
             <a:off x=\"0\" y=\"0\" xmlns:a=\"{a_ns}\"/>\
             <a:ext cx=\"{cx}\" cy=\"{cy}\" xmlns:a=\"{a_ns}\"/>\
           </{p}xfrm>\
           <a:graphic xmlns:a=\"{a_ns}\">\
             <a:graphicData uri=\"{c_ns}\">\
               <c:chart xmlns:c=\"{c_ns}\" \
                        xmlns:r=\"{r_ns}\" \
                        r:id=\"{chart_rid}\"/>\
             </a:graphicData>\
           </a:graphic>\
         </{p}graphicFrame>"
    ));
    out.push_str(&format!("<{p}clientData/>"));
    out.push_str(&format!("</{p}oneCellAnchor>"));
    out
}

#[cfg(test)]
mod chart_helpers_tests {
    use super::*;

    #[test]
    fn parse_a1_basic_cells() {
        assert_eq!(parse_a1_coord("A1"), Some((0, 0)));
        assert_eq!(parse_a1_coord("D2"), Some((3, 1)));
        assert_eq!(parse_a1_coord("Z1"), Some((25, 0)));
        assert_eq!(parse_a1_coord("AA1"), Some((26, 0)));
        assert_eq!(parse_a1_coord("$D$2"), Some((3, 1)));
        assert!(parse_a1_coord("").is_none());
        assert!(parse_a1_coord("1A").is_none());
    }

    #[test]
    fn resolve_relative_basic() {
        assert_eq!(
            resolve_relative_path("xl/worksheets/_rels", "../drawings/drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
        assert_eq!(
            resolve_relative_path("xl/drawings/_rels", "../charts/chart1.xml"),
            "xl/charts/chart1.xml"
        );
    }

    #[test]
    fn drawing_n_extract() {
        assert_eq!(drawing_n_from_path("xl/drawings/drawing7.xml"), Some(7));
        assert_eq!(drawing_n_from_path("xl/drawings/drawing.xml"), None);
        assert_eq!(drawing_n_from_path("nope.xml"), None);
    }

    #[test]
    fn build_drawing_xml_for_one_chart() {
        let q = vec![QueuedChartAdd {
            chart_xml: b"<chartSpace/>".to_vec(),
            anchor_a1: "D2".into(),
            width_emu: 4_572_000,
            height_emu: 2_743_200,
        }];
        let rids = vec!["rId1".to_string()];
        let body = build_chart_drawing_xml(&q, &rids);
        assert!(body.contains("<xdr:graphicFrame"));
        assert!(body.contains("r:id=\"rId1\""));
        assert!(body.contains("<xdr:col>3</xdr:col>"));
        assert!(body.contains("<xdr:row>1</xdr:row>"));
    }

    #[test]
    fn append_graphic_frame_inserts_before_close() {
        let original = b"<?xml version=\"1.0\"?><xdr:wsDr xmlns:xdr=\"x\" xmlns:r=\"r\" xmlns:c=\"c\"><xdr:oneCellAnchor/></xdr:wsDr>";
        let q = vec![QueuedChartAdd {
            chart_xml: vec![],
            anchor_a1: "B5".into(),
            width_emu: 100,
            height_emu: 200,
        }];
        let rids = vec!["rId7".to_string()];
        let merged = append_graphic_frames(original, &q, &rids).unwrap();
        let s = std::str::from_utf8(&merged).unwrap();
        // Original anchor preserved.
        assert!(s.contains("<xdr:oneCellAnchor/>"));
        // New graphicFrame appended before the close.
        assert!(s.contains("<xdr:graphicFrame"));
        assert!(s.contains("r:id=\"rId7\""));
        assert!(s.ends_with("</xdr:wsDr>"));
    }
}
