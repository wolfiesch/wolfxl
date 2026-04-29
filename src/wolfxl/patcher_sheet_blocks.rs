//! Sheet-scoped block save phases for the surgical xlsx patcher.

use std::collections::HashMap;

use super::XlsxPatcher;
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
