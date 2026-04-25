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

use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Write};

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::ooxml_util;
use sheet_patcher::{CellPatch, CellValue};
use styles::FormatSpec;
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
        for (name, rid) in sheet_rids {
            if let Some(target) = rel_targets.get(&rid) {
                sheet_paths.insert(name, ooxml_util::join_and_normalize("xl/", target));
            }
        }

        Ok(XlsxPatcher {
            file_path: path.to_string(),
            sheet_paths,
            value_patches: HashMap::new(),
            format_patches: HashMap::new(),
            rels_patches: HashMap::new(),
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
    fn sheet_names(&self) -> Vec<String> {
        self.sheet_paths.keys().cloned().collect()
    }

    /// Save patched file to a new path.
    fn save(&self, path: &str) -> PyResult<()> {
        self.do_save(path)
    }

    /// Save in-place (atomic tmp+rename).
    fn save_in_place(&self) -> PyResult<()> {
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
}

// ---------------------------------------------------------------------------
// Save implementation
// ---------------------------------------------------------------------------

impl XlsxPatcher {
    fn do_save(&self, output_path: &str) -> PyResult<()> {
        if self.value_patches.is_empty()
            && self.format_patches.is_empty()
            && self.rels_patches.is_empty()
        {
            // No changes — just copy
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

        // --- Phase 3: Patch worksheet XMLs ---
        let mut file_patches: HashMap<String, Vec<u8>> = HashMap::new();

        for (sheet_path, patches) in &sheet_cell_patches {
            let xml = ooxml_util::zip_read_to_string(&mut zip, sheet_path)?;
            let patched = sheet_patcher::patch_worksheet(&xml, patches)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("Patch failed: {e}")))?;
            file_patches.insert(sheet_path.clone(), patched.into_bytes());
        }

        // Add styles.xml patch if modified
        if let Some(ref sxml) = styles_xml {
            file_patches.insert("xl/styles.xml".to_string(), sxml.as_bytes().to_vec());
        }

        // Serialize any mutated `*.rels` graphs into file_patches. This branch
        // is dead code in the current slice (rels_patches is always empty);
        // RFC-022/023/024 will populate it.
        for (path, graph) in &self.rels_patches {
            file_patches.insert(path.clone(), graph.serialize());
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

        for i in 0..zip.len() {
            let mut file = zip
                .by_index(i)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP entry read error: {e}")))?;
            let name = file.name().to_string();

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

        out.finish()
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP finalize error: {e}")))?;

        Ok(())
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
