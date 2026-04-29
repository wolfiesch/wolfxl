//! Sheet-scoped block save phases for the surgical xlsx patcher.

use std::collections::{HashMap, HashSet};
use std::fs::File;

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use zip::ZipArchive;

use crate::ooxml_util;

use super::patcher_workbook::{load_or_empty_rels, minimal_styles_xml, sheet_rels_path_for};
use super::{autofilter, autofilter_helpers, hyperlinks, sheet_patcher, XlsxPatcher};
use sheet_patcher::CellPatch;
use wolfxl_merger::SheetBlock;

pub(super) fn apply_data_validations_phase(
    patcher: &XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    for (sheet_name, patches) in &patcher.queued_dv_patches {
        let sheet_path = match patcher.sheet_paths.get(sheet_name) {
            Some(p) => p,
            None => continue,
        };
        let xml = ooxml_util::zip_read_to_string(zip, sheet_path)?;
        let existing = super::validations::extract_existing_dv_block(&xml);
        let block_bytes =
            super::validations::build_data_validations_block(existing.as_deref(), patches);
        local_blocks
            .entry(sheet_path.clone())
            .or_default()
            .push(SheetBlock::DataValidations(block_bytes));
    }

    Ok(())
}

pub(super) fn apply_conditional_formatting_phase(
    patcher: &XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
    styles_xml: &mut Option<String>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    let mut new_dxfs_total: Vec<super::conditional_formatting::DxfPatch> = Vec::new();
    let mut styles_loaded: Option<String> = None;
    let mut running_dxf_count: u32 = 0;
    let mut cf_sheet_names: Vec<&String> = patcher.queued_cf_patches.keys().collect();
    cf_sheet_names.sort();

    for sheet_name in cf_sheet_names {
        let patches = &patcher.queued_cf_patches[sheet_name];
        let sheet_path = match patcher.sheet_paths.get(sheet_name) {
            Some(p) => p,
            None => continue,
        };
        let xml = ooxml_util::zip_read_to_string(zip, sheet_path)?;

        if styles_loaded.is_none() {
            let raw = ooxml_util::zip_read_to_string_opt(zip, "xl/styles.xml")?
                .unwrap_or_else(minimal_styles_xml);
            running_dxf_count = super::conditional_formatting::count_dxfs(&raw);
            styles_loaded = Some(raw);
        }

        let existing = super::conditional_formatting::extract_existing_cf_blocks(&xml);
        let pmax = super::conditional_formatting::scan_max_cf_priority(&xml);
        let result = super::conditional_formatting::build_cf_blocks(
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

    if !new_dxfs_total.is_empty() {
        let new_dxfs_xml: String = new_dxfs_total
            .iter()
            .map(super::conditional_formatting::dxf_to_xml)
            .collect::<Vec<_>>()
            .join("");
        let base = match styles_xml.take() {
            Some(s) => s,
            None => styles_loaded.unwrap_or_else(minimal_styles_xml),
        };
        let updated = super::conditional_formatting::ensure_dxfs_section(&base, &new_dxfs_xml);
        *styles_xml = Some(updated);
    }

    Ok(())
}

pub(super) fn apply_hyperlinks_phase(
    patcher: &mut XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    let sheet_order_local: Vec<String> = patcher.sheet_order.clone();
    for sheet_name in &sheet_order_local {
        let ops = match patcher.queued_hyperlinks.get(sheet_name) {
            Some(o) if !o.is_empty() => o.clone(),
            _ => continue,
        };
        let sheet_path = match patcher.sheet_paths.get(sheet_name).cloned() {
            Some(p) => p,
            None => continue,
        };
        let rels_path = sheet_rels_path_for(&sheet_path);
        patcher
            .ancillary
            .populate_for_sheet(zip, sheet_name, &sheet_path)
            .map_err(|e| {
                PyIOError::new_err(format!("ancillary populate for '{sheet_name}': {e}"))
            })?;
        if !patcher.rels_patches.contains_key(&rels_path) {
            let graph = load_or_empty_rels(zip, &rels_path)?;
            patcher.rels_patches.insert(rels_path.clone(), graph);
        }
        let rels = patcher
            .rels_patches
            .get_mut(&rels_path)
            .expect("just inserted above");
        let xml = ooxml_util::zip_read_to_string(zip, &sheet_path)?;
        let existing = hyperlinks::extract_hyperlinks(xml.as_bytes(), rels);
        let had_existing = !existing.is_empty();
        let (block_bytes, _deleted_rids) = hyperlinks::build_hyperlinks_block(existing, &ops, rels);

        if block_bytes.is_empty() && !had_existing {
            continue;
        }
        local_blocks
            .entry(sheet_path)
            .or_default()
            .push(SheetBlock::Hyperlinks(block_bytes));
    }

    Ok(())
}

pub(super) fn apply_sheet_setup_phase(
    patcher: &XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
) {
    let sheet_titles: Vec<String> = patcher.queued_sheet_setup.keys().cloned().collect();
    for sheet_title in &sheet_titles {
        let queued = match patcher.queued_sheet_setup.get(sheet_title) {
            Some(q) => q,
            None => continue,
        };
        let sheet_path = match patcher.sheet_paths.get(sheet_title) {
            Some(p) => p.clone(),
            None => continue,
        };
        let specs = &queued.specs;

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
            let bytes = wolfxl_writer::parse::sheet_setup::emit_sheet_protection(s);
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
        // print_titles is routed through workbook definedNames by the
        // Python coordinator. The patcher queue entry is informational.
    }
}

pub(super) fn apply_page_breaks_phase(
    patcher: &XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
) {
    let sheet_titles: Vec<String> = patcher.queued_page_breaks.keys().cloned().collect();
    for sheet_title in &sheet_titles {
        let queued = match patcher.queued_page_breaks.get(sheet_title) {
            Some(q) => q,
            None => continue,
        };
        let sheet_path = match patcher.sheet_paths.get(sheet_title) {
            Some(p) => p.clone(),
            None => continue,
        };

        if let Some(spec) = &queued.sheet_format {
            let bytes = wolfxl_writer::parse::page_breaks::emit_sheet_format_pr(spec);
            if !bytes.is_empty() {
                local_blocks
                    .entry(sheet_path.clone())
                    .or_default()
                    .push(SheetBlock::SheetFormatPr(bytes));
            }
        }
        if let Some(spec) = &queued.row_breaks {
            let bytes = wolfxl_writer::parse::page_breaks::emit_row_breaks(spec);
            if !bytes.is_empty() {
                local_blocks
                    .entry(sheet_path.clone())
                    .or_default()
                    .push(SheetBlock::RowBreaks(bytes));
            }
        }
        if let Some(spec) = &queued.col_breaks {
            let bytes = wolfxl_writer::parse::page_breaks::emit_col_breaks(spec);
            if !bytes.is_empty() {
                local_blocks
                    .entry(sheet_path.clone())
                    .or_default()
                    .push(SheetBlock::ColBreaks(bytes));
            }
        }
    }
}

pub(super) fn apply_autofilter_phase(
    patcher: &XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
    file_patches: &HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<HashMap<String, Vec<u32>>> {
    let mut autofilter_hidden_rows: HashMap<String, Vec<u32>> = HashMap::new();
    let sheet_titles: Vec<String> = patcher.queued_autofilters.keys().cloned().collect();
    for sheet_title in &sheet_titles {
        let queued = patcher
            .queued_autofilters
            .get(sheet_title)
            .cloned()
            .unwrap();
        let sheet_path = match patcher.sheet_paths.get(sheet_title) {
            Some(p) => p.clone(),
            None => continue,
        };

        let xml_bytes: Vec<u8> = if let Some(b) = file_patches.get(&sheet_path) {
            b.clone()
        } else if let Some(b) = patcher.file_adds.get(&sheet_path) {
            b.clone()
        } else {
            ooxml_util::zip_read_to_string(zip, &sheet_path)?.into_bytes()
        };

        let af_model = wolfxl_autofilter::parse::parse_autofilter(&queued.dict)
            .map_err(|e| PyErr::new::<PyValueError, _>(format!("Phase 2.5o: {e}")))?;
        let (start_row, end_row, start_col, end_col) = match af_model
            .ref_
            .as_deref()
            .and_then(autofilter_helpers::parse_a1_range)
        {
            Some(t) => t,
            None => {
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

        let data_start = start_row + 1;
        let rows_data = autofilter_helpers::extract_cell_grid(
            &xml_bytes, data_start, end_row, start_col, end_col,
        )?;

        let drain = autofilter::drain_autofilter(&queued, &rows_data, None)
            .map_err(|e| PyErr::new::<PyValueError, _>(format!("Phase 2.5o: {e}")))?;

        let abs_hidden: Vec<u32> = drain
            .hidden_offsets
            .iter()
            .map(|off| data_start + off)
            .collect();
        if !abs_hidden.is_empty() {
            autofilter_hidden_rows.insert(sheet_path.clone(), abs_hidden);
        }

        if !drain.block_bytes.is_empty() {
            local_blocks
                .entry(sheet_path.clone())
                .or_default()
                .push(SheetBlock::AutoFilter(drain.block_bytes));
        }
    }

    Ok(autofilter_hidden_rows)
}

pub(super) fn apply_worksheet_xml_patch_phase(
    patcher: &mut XlsxPatcher,
    sheet_cell_patches: &HashMap<String, Vec<CellPatch>>,
    local_blocks: &HashMap<String, Vec<SheetBlock>>,
    autofilter_hidden_rows: &HashMap<String, Vec<u32>>,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    let mut all_sheet_paths: HashSet<String> = HashSet::new();
    all_sheet_paths.extend(sheet_cell_patches.keys().cloned());
    all_sheet_paths.extend(local_blocks.keys().cloned());
    all_sheet_paths.extend(autofilter_hidden_rows.keys().cloned());

    for sheet_path in &all_sheet_paths {
        // Compose cell edits, sibling OOXML blocks, and autoFilter row hiding
        // against the newest sheet bytes, including Phase 2.7 cloned sheets.
        let xml = if let Some(bytes) = file_patches.get(sheet_path) {
            String::from_utf8_lossy(bytes).into_owned()
        } else if let Some(bytes) = patcher.file_adds.get(sheet_path) {
            String::from_utf8_lossy(bytes).into_owned()
        } else {
            ooxml_util::zip_read_to_string(zip, sheet_path)?
        };

        let after_cells: Vec<u8> = if let Some(patches) = sheet_cell_patches.get(sheet_path) {
            sheet_patcher::patch_worksheet(&xml, patches)
                .map_err(|e| PyIOError::new_err(format!("Patch failed: {e}")))?
                .into_bytes()
        } else {
            xml.into_bytes()
        };

        let after_blocks = if let Some(blocks) = local_blocks.get(sheet_path) {
            if blocks.is_empty() {
                after_cells
            } else {
                wolfxl_merger::merge_blocks(&after_cells, blocks.clone())
                    .map_err(|e| PyIOError::new_err(format!("Merge failed: {e}")))?
            }
        } else {
            after_cells
        };

        let after_blocks = if let Some(rows) = autofilter_hidden_rows.get(sheet_path) {
            if rows.is_empty() {
                after_blocks
            } else {
                autofilter_helpers::stamp_row_hidden(&after_blocks, rows)?
            }
        } else {
            after_blocks
        };

        if patcher.file_adds.contains_key(sheet_path) {
            patcher.file_adds.insert(sheet_path.clone(), after_blocks);
        } else {
            file_patches.insert(sheet_path.clone(), after_blocks);
        }
    }

    Ok(())
}
