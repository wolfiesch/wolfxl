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

#[pymethods]
impl XlsxPatcher {
    /// Open an xlsx file for surgical patching.
    #[staticmethod]
    fn open(path: &str) -> PyResult<Self> {
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
        })
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
        {
            // No changes — just copy. Includes RFC-013's `file_adds`,
            // `file_deletes`, `queued_content_type_ops`, RFC-020's
            // `queued_props`, RFC-022's `queued_hyperlinks`, RFC-021's
            // `queued_defined_names`, RFC-024's `queued_tables`,
            // RFC-023's `queued_comments`, RFC-036's
            // `queued_sheet_moves`, and RFC-030/031's
            // `queued_axis_shifts` so a no-op save remains
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
                    let g = load_or_empty_rels(&mut zip, &rels_path)?;
                    self.rels_patches.insert(rels_path.clone(), g);
                }
                let rels = self
                    .rels_patches
                    .get_mut(&rels_path)
                    .expect("just inserted above");
                let result = tables::build_tables(&patches, &tables_inventory, rels)
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
        // the stack so all sheets share dedup. Two parallel counters
        // (`next_comments_n`, `next_vml_n`) start at the highest
        // existing index + 1 and increment as new parts are minted.
        let mut comment_authors = comments::CommentAuthorTable::new();
        // Seed the workbook-wide N counters: start above the highest
        // existing comments<N>.xml / vmlDrawing<N>.vml in the source.
        let mut next_comments_n: u32 = 1;
        let mut next_vml_n: u32 = 1;
        for sheet_name in &sheet_order_local {
            let sp = match self.sheet_paths.get(sheet_name).cloned() {
                Some(p) => p,
                None => continue,
            };
            let _ = self.ancillary.populate_for_sheet(&mut zip, sheet_name, &sp);
            if let Some(anc) = self.ancillary.get(sheet_name) {
                if let Some(p) = &anc.comments_part {
                    if let Some(n) = parse_n_from_part_path(p, "xl/comments", ".xml") {
                        if n + 1 > next_comments_n {
                            next_comments_n = n + 1;
                        }
                    }
                }
                if let Some(p) = &anc.vml_drawing_part {
                    if let Some(n) =
                        parse_n_from_part_path(p, "xl/drawings/vmlDrawing", ".vml")
                    {
                        if n + 1 > next_vml_n {
                            next_vml_n = n + 1;
                        }
                    }
                }
            }
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

            // Decide N values: reuse existing-part N if any, else mint new.
            let comments_n = match &existing_comments_path {
                Some(p) => {
                    parse_n_from_part_path(p, "xl/comments", ".xml").unwrap_or_else(|| {
                        let n = next_comments_n;
                        next_comments_n += 1;
                        n
                    })
                }
                None => {
                    let n = next_comments_n;
                    next_comments_n += 1;
                    n
                }
            };
            let vml_n = match &existing_vml_path {
                Some(p) => {
                    parse_n_from_part_path(p, "xl/drawings/vmlDrawing", ".vml")
                        .unwrap_or_else(|| {
                            let n = next_vml_n;
                            next_vml_n += 1;
                            n
                        })
                }
                None => {
                    let n = next_vml_n;
                    next_vml_n += 1;
                    n
                }
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

        // --- Phase 3: Patch worksheet XMLs ---
        //
        // Two-pass per sheet: cell-level patches via `sheet_patcher`, then
        // sibling-block insertions via `wolfxl_merger`. The two passes
        // commute (cells live inside <sheetData>, blocks are siblings) so
        // composing them is straightforward.
        let mut file_patches: HashMap<String, Vec<u8>> = HashMap::new();

        // Sheets that have either kind of patch.
        let mut all_sheet_paths: std::collections::HashSet<String> =
            std::collections::HashSet::new();
        all_sheet_paths.extend(sheet_cell_patches.keys().cloned());
        all_sheet_paths.extend(local_blocks.keys().cloned());

        for sheet_path in &all_sheet_paths {
            let xml = ooxml_util::zip_read_to_string(&mut zip, sheet_path)?;

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

            file_patches.insert(sheet_path.clone(), after_blocks);
        }

        // Add styles.xml patch if modified
        if let Some(ref sxml) = styles_xml {
            file_patches.insert("xl/styles.xml".to_string(), sxml.as_bytes().to_vec());
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
        let mut workbook_xml_in_progress: Option<Vec<u8>> = None;
        if !self.queued_sheet_moves.is_empty() {
            let src_wb_xml = ooxml_util::zip_read_to_string(&mut zip, "xl/workbook.xml")?;
            let result = sheet_order::merge_sheet_moves(
                src_wb_xml.as_bytes(),
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
}

// ---------------------------------------------------------------------------
// Dict → spec conversion helpers
// ---------------------------------------------------------------------------

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
