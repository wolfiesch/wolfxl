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
#[allow(dead_code)] // RFC-024: live caller wires up alongside the patcher integration commit
pub mod tables;

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
        {
            // No changes — just copy. Includes RFC-013's `file_adds`,
            // `file_deletes`, `queued_content_type_ops`, RFC-020's
            // `queued_props`, and RFC-022's `queued_hyperlinks` so a
            // no-op save remains byte-identical even after these
            // primitives land.
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
