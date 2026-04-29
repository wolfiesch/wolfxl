//! Sheet-scoped block save phases for the surgical xlsx patcher.

use std::collections::HashMap;
use std::fs::File;

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use zip::ZipArchive;

use crate::ooxml_util;

use super::{autofilter, autofilter_helpers, XlsxPatcher};
use wolfxl_merger::SheetBlock;

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
