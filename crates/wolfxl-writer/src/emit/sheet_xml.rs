//! `xl/worksheets/sheet{N}.xml` emitter — rows, cells, merges, freeze,
//! columns, print area, and extension hooks for CF/DV. Wave 2B.
//!
//! # Element ordering
//!
//! CT_Worksheet's `<xsd:sequence>` declares 38 ordered child elements
//! (ECMA-376 §18.3.1.99). This emitter walks them in the order pinned
//! by [`wolfxl_merger::ct_worksheet_order::ECMA_ORDER`] — the same
//! table the modify-mode merger uses to insert sibling blocks into an
//! existing sheet. Section comments below cite the slot number from
//! that table (e.g. `slot 6: <sheetData>`); if the spec is ever
//! extended, update `ECMA_ORDER` once and both this emitter and the
//! merger pick it up.
//!
//! # rId convention (must match [`crate::emit::rels::emit_sheet`])
//!
//! Sheet-level relationships are allocated in this order inside
//! `xl/worksheets/_rels/sheet{N}.xml.rels`:
//!
//! 1. **Comments** (if any): `rId1` points at `commentsN.xml`,
//!    `rId2` at `vmlDrawingN.vml`.
//! 2. **Tables**: the next contiguous block. With no comments,
//!    tables start at `rId1`; with comments, at `rId3`.
//! 3. **External hyperlinks** (targets that do not start with `#`):
//!    the tail of the rId range.
//!
//! The emitter MUST walk [`Worksheet::hyperlinks`] with the same
//! filter + iteration order as `rels::emit_sheet` uses when assigning
//! `r:id` attributes in `<hyperlink>` elements, or Excel will follow
//! mismatched rIds and silently drop hyperlink targets.
//!
//! # Extension hooks (Wave 3)
//!
//! The emitter leaves `// EXT-W3A:`, `// EXT-W3B:`, and `// EXT-W3C:`
//! marker comments at the three insertion points where Wave 3 agents
//! plug in comments/VML bridging, tables, conditional formats, and
//! data validations. Keep them even when the related collections
//! are empty — Wave 3 may need to emit structural parents.

use crate::intern::SstBuilder;
use crate::model::cell::{FormulaResult, WriteCellValue};
use crate::model::format::StylesBuilder;
use crate::model::worksheet::Worksheet;
use crate::{refs, xml_escape};

/// Emit `xl/worksheets/sheet{N}.xml` bytes for one sheet.
///
/// `sheet_idx` is zero-based; the caller converts to 1-based for any
/// user-facing references (`sheet1.xml`, `commentsN.xml`, etc.).
///
/// `sst` is mutable because string cells intern at emit time, not model
/// construction time. `styles` is immutable because all interning already
/// happened during `WriteCell` construction.
pub fn emit(
    sheet: &Worksheet,
    sheet_idx: u32,
    sst: &mut SstBuilder,
    _styles: &StylesBuilder,
) -> Vec<u8> {
    let mut out = String::with_capacity(4096);

    // XML declaration + root element
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    out.push_str("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

    // Slot 2: <dimension>
    emit_dimension(&mut out, sheet);

    // Slot 3: <sheetViews>
    emit_sheet_views(&mut out, sheet, sheet_idx);

    // Slot 4: <sheetFormatPr>
    out.push_str("<sheetFormatPr defaultRowHeight=\"15\"/>");

    // Slot 5: <cols> (only if non-empty)
    if !sheet.columns.is_empty() {
        emit_cols(&mut out, sheet);
    }

    // Slot 6: <sheetData>
    emit_sheet_data(&mut out, sheet, sst);

    // Slot 11: <autoFilter> — Sprint Ο Pod 1B (RFC-056). The bytes
    // are pre-emitted by the workbook-level coordinator from the
    // Python `ws.auto_filter.to_rust_dict()` payload via
    // `wolfxl_autofilter::emit::emit`.
    if let Some(bytes) = &sheet.auto_filter_xml {
        out.push_str(std::str::from_utf8(bytes).unwrap_or(""));
    }

    // Slot 15: <mergeCells> (only if non-empty)
    if !sheet.merges.is_empty() {
        emit_merges(&mut out, sheet);
    }

    // Slot 17: <conditionalFormatting> — EXT-W3C; 0..N elements per spec
    emit_conditional_formats(&mut out, sheet);

    // Slot 18: <dataValidations> — EXT-W3C
    emit_data_validations(&mut out, sheet);

    // Slot 19: <hyperlinks> (only if any exist)
    if !sheet.hyperlinks.is_empty() {
        emit_hyperlinks(&mut out, sheet);
    }

    // Slot 21: <pageMargins>
    out.push_str("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>");

    // Slot 30: <drawing r:id="..."/> — Sprint Λ Pod-β (RFC-045);
    // emitted iff !sheet.images.is_empty(). The rId is appended at
    // the END of the sheet's rels graph (after comments, vml, tables,
    // and external hyperlinks) so the existing rId conventions for
    // those entries are preserved.
    emit_drawing_ref(&mut out, sheet);

    // Slot 31: <legacyDrawing> — EXT-W3A; emitted iff !sheet.comments.is_empty(); rId via convention
    emit_legacy_drawing(&mut out, sheet);

    // Slot 37: <tableParts> — EXT-W3B; one <tablePart r:id=...> per table
    emit_table_parts(&mut out, sheet);

    // Slot numbers above match wolfxl_merger::ct_worksheet_order::ECMA_ORDER
    // (the merger crate's own tests assert the table is the canonical 38-slot
    // §18.3.1.99 sequence; this emitter takes those numbers as the contract).
    out.push_str("</worksheet>");

    out.into_bytes()
}

/// Compile-time assertion that the slot numbers cited in `emit`'s section
/// comments match `wolfxl_merger::ct_worksheet_order::ECMA_ORDER`. If a
/// future ECMA extension renumbers a slot, this fails to compile until both
/// this emitter and the merger are updated together.
#[allow(dead_code)]
const _PIN_SLOT_NUMBERS: () = {
    let order = wolfxl_merger::ct_worksheet_order::ECMA_ORDER;
    // Slots cited in `emit` above (zero-indexed positions in ECMA_ORDER).
    assert!(order[1].1 == 2); // dimension
    assert!(order[5].1 == 6); // sheetData
    assert!(order[14].1 == 15); // mergeCells
    assert!(order[16].1 == 17); // conditionalFormatting
    assert!(order[17].1 == 18); // dataValidations
    assert!(order[18].1 == 19); // hyperlinks
    assert!(order[20].1 == 21); // pageMargins
    assert!(order[30].1 == 31); // legacyDrawing
    assert!(order[36].1 == 37); // tableParts
};

/// Compute bounding box of all populated cells and emit `<dimension ref="…"/>`.
fn emit_dimension(out: &mut String, sheet: &Worksheet) {
    let mut min_row = u32::MAX;
    let mut max_row = 0u32;
    let mut min_col = u32::MAX;
    let mut max_col = 0u32;

    for (&row_num, row) in &sheet.rows {
        for (&col_num, cell) in &row.cells {
            if matches!(cell.value, WriteCellValue::Blank) && cell.style_id.is_none() {
                continue;
            }
            min_row = min_row.min(row_num);
            max_row = max_row.max(row_num);
            min_col = min_col.min(col_num);
            max_col = max_col.max(col_num);
        }
        // Also consider rows with custom attrs but no cells (they don't expand bounding box
        // unless they have cells — spec says dimension is the range of cell data)
    }

    if max_row == 0 {
        // No cells — emit minimal A1 reference
        out.push_str("<dimension ref=\"A1\"/>");
    } else {
        let range = refs::format_range((min_row, min_col), (max_row, max_col));
        out.push_str(&format!("<dimension ref=\"{}\"/>", range));
    }
}

/// Emit `<sheetViews><sheetView …>…</sheetView></sheetViews>`.
fn emit_sheet_views(out: &mut String, sheet: &Worksheet, sheet_idx: u32) {
    out.push_str("<sheetViews>");

    if sheet_idx == 0 {
        out.push_str("<sheetView tabSelected=\"1\" workbookViewId=\"0\">");
    } else {
        out.push_str("<sheetView workbookViewId=\"0\">");
    }

    // Emit pane — prefer freeze over split
    if let Some(freeze) = &sheet.freeze {
        emit_freeze_pane(out, freeze);
    } else if let Some(split) = &sheet.split {
        emit_split_pane(out, split);
    }

    out.push_str("</sheetView>");
    out.push_str("</sheetViews>");
}

/// Emit `<pane …/>` for a freeze pane.
///
/// OOXML pane semantics with `state="frozen"`:
///   `xSplit` = number of columns frozen
///   `ySplit` = number of rows frozen
/// Our model stores the freeze CELL coordinate ("B2" -> row=2 col=2),
/// so the spec count is `(coord - 1)`. A coord of 0 or 1 means no
/// freeze on that axis (nothing above row 1 or left of col A).
fn emit_freeze_pane(out: &mut String, freeze: &crate::model::worksheet::FreezePane) {
    let y_split = freeze.freeze_row.saturating_sub(1);
    let x_split = freeze.freeze_col.saturating_sub(1);
    let has_row = y_split > 0;
    let has_col = x_split > 0;
    if !has_row && !has_col {
        return; // freeze cell at A1 (or below) — degenerate, emit nothing.
    }

    let tl_row = freeze.top_left.map(|t| t.0).unwrap_or_else(|| {
        if has_row { freeze.freeze_row } else { 1 }
    });
    let tl_col = freeze.top_left.map(|t| t.1).unwrap_or_else(|| {
        if has_col { freeze.freeze_col } else { 1 }
    });

    let active_pane = if has_row && has_col {
        "bottomRight"
    } else if has_row {
        "bottomLeft"
    } else {
        "topRight"
    };

    out.push_str("<pane");
    if has_col {
        out.push_str(&format!(" xSplit=\"{}\"", x_split));
    }
    if has_row {
        out.push_str(&format!(" ySplit=\"{}\"", y_split));
    }
    let top_left_cell = refs::format_a1(tl_row, tl_col);
    out.push_str(&format!(" topLeftCell=\"{}\"", top_left_cell));
    out.push_str(&format!(" activePane=\"{}\"", active_pane));
    out.push_str(" state=\"frozen\"");
    out.push_str("/>");
}

/// Emit `<pane …/>` for a split (non-frozen) pane.
fn emit_split_pane(out: &mut String, split: &crate::model::worksheet::SplitPane) {
    let tl_row = split.top_left.map(|t| t.0).unwrap_or(1);
    let tl_col = split.top_left.map(|t| t.1).unwrap_or(1);

    let has_x = split.x_split != 0.0;
    let has_y = split.y_split != 0.0;

    let active_pane = if has_x && has_y {
        "bottomRight"
    } else if has_y {
        "bottomLeft"
    } else {
        "topRight"
    };

    out.push_str("<pane");

    if has_x {
        out.push_str(&format!(" xSplit=\"{:.2}\"", split.x_split));
    }
    if has_y {
        out.push_str(&format!(" ySplit=\"{:.2}\"", split.y_split));
    }

    let top_left_cell = refs::format_a1(tl_row, tl_col);
    out.push_str(&format!(" topLeftCell=\"{}\"", top_left_cell));
    out.push_str(&format!(" activePane=\"{}\"", active_pane));
    // No state="frozen" for split panes
    out.push_str("/>");
}

/// Emit `<cols>…</cols>`.
fn emit_cols(out: &mut String, sheet: &Worksheet) {
    out.push_str("<cols>");

    for (&col_idx, col) in &sheet.columns {
        // Skip completely-default columns
        if col.width.is_none() && !col.hidden && col.outline_level == 0 && col.style_id.is_none() {
            continue;
        }

        out.push_str(&format!("<col min=\"{}\" max=\"{}\"", col_idx, col_idx));

        if let Some(w) = col.width {
            out.push_str(&format!(" width=\"{}\" customWidth=\"1\"", format_f64(w)));
        }
        if col.hidden {
            out.push_str(" hidden=\"1\"");
        }
        if col.outline_level > 0 {
            out.push_str(&format!(" outlineLevel=\"{}\"", col.outline_level));
        }
        if let Some(s) = col.style_id {
            out.push_str(&format!(" style=\"{}\" customFormat=\"1\"", s));
        }

        out.push_str("/>");
    }

    out.push_str("</cols>");
}

/// Emit `<sheetData>…</sheetData>`.
fn emit_sheet_data(out: &mut String, sheet: &Worksheet, sst: &mut SstBuilder) {
    if sheet.rows.is_empty() {
        out.push_str("<sheetData/>");
        return;
    }

    out.push_str("<sheetData>");

    for (&row_num, row) in &sheet.rows {
        emit_row(out, row_num, row, sst);
    }

    out.push_str("</sheetData>");
}

/// Emit one `<row>` element with its cells.
fn emit_row(
    out: &mut String,
    row_num: u32,
    row: &crate::model::worksheet::Row,
    sst: &mut SstBuilder,
) {
    // Check if any cells have content (non-blank or styled blank)
    let has_real_cells = row.cells.values().any(|c| {
        !matches!(c.value, WriteCellValue::Blank) || c.style_id.is_some()
    });
    let has_attrs = row.custom_height.is_some() || row.hidden || row.style_id.is_some();

    // If no cells and no attributes, still emit row if cells exist
    if row.cells.is_empty() && !has_attrs {
        return;
    }

    out.push_str(&format!("<row r=\"{}\"", row_num));

    if let Some(h) = row.custom_height {
        out.push_str(&format!(" ht=\"{}\" customHeight=\"1\"", format_f64(h)));
    }
    if row.hidden {
        out.push_str(" hidden=\"1\"");
    }
    if let Some(s) = row.style_id {
        out.push_str(&format!(" s=\"{}\" customFormat=\"1\"", s));
    }

    if row.cells.is_empty() || !has_real_cells {
        // Self-close if no renderable cells
        if !has_real_cells {
            out.push_str("/>");
            return;
        }
    }

    out.push('>');

    for (&col_num, cell) in &row.cells {
        emit_cell(out, row_num, col_num, cell, sst);
    }

    out.push_str("</row>");
}

/// Emit one `<c>` element.
fn emit_cell(
    out: &mut String,
    row_num: u32,
    col_num: u32,
    cell: &crate::model::cell::WriteCell,
    sst: &mut SstBuilder,
) {
    let cell_ref = refs::format_a1(row_num, col_num);

    match &cell.value {
        WriteCellValue::Blank => {
            // Only emit if styled
            if let Some(s) = cell.style_id {
                out.push_str(&format!("<c r=\"{}\" s=\"{}\"/>", cell_ref, s));
            }
            // Otherwise skip entirely
        }

        WriteCellValue::Number(n) => {
            out.push_str(&format!("<c r=\"{}\"", cell_ref));
            if let Some(s) = cell.style_id {
                out.push_str(&format!(" s=\"{}\"", s));
            }
            out.push_str(&format!("><v>{}</v></c>", format_number(*n)));
        }

        WriteCellValue::String(s) => {
            let idx = sst.intern(s);
            out.push_str(&format!("<c r=\"{}\" t=\"s\"", cell_ref));
            if let Some(style) = cell.style_id {
                out.push_str(&format!(" s=\"{}\"", style));
            }
            out.push_str(&format!("><v>{}</v></c>", idx));
        }

        WriteCellValue::Boolean(b) => {
            out.push_str(&format!("<c r=\"{}\" t=\"b\"", cell_ref));
            if let Some(s) = cell.style_id {
                out.push_str(&format!(" s=\"{}\"", s));
            }
            let bval = if *b { 1 } else { 0 };
            out.push_str(&format!("><v>{}</v></c>", bval));
        }

        WriteCellValue::Formula { expr, result } => {
            let escaped_expr = xml_escape::text(expr);
            match result {
                None => {
                    // Emit a stub <v>0</v> so reader-only paths (calamine,
                    // xlsx2csv) see *some* cached value. Without it, calamine
                    // hands back `None` for every formula cell since it does
                    // not evaluate expressions itself. Mirrors rust_xlsxwriter.
                    out.push_str(&format!("<c r=\"{}\"", cell_ref));
                    if let Some(s) = cell.style_id {
                        out.push_str(&format!(" s=\"{}\"", s));
                    }
                    out.push_str(&format!("><f>{}</f><v>0</v></c>", escaped_expr));
                }
                Some(FormulaResult::Number(n)) => {
                    out.push_str(&format!("<c r=\"{}\"", cell_ref));
                    if let Some(s) = cell.style_id {
                        out.push_str(&format!(" s=\"{}\"", s));
                    }
                    out.push_str(&format!(
                        "><f>{}</f><v>{}</v></c>",
                        escaped_expr,
                        format_number(*n)
                    ));
                }
                Some(FormulaResult::String(s)) => {
                    // t="str" — inline formula result string (not SST-indexed)
                    out.push_str(&format!("<c r=\"{}\" t=\"str\"", cell_ref));
                    if let Some(style) = cell.style_id {
                        out.push_str(&format!(" s=\"{}\"", style));
                    }
                    out.push_str(&format!(
                        "><f>{}</f><v>{}</v></c>",
                        escaped_expr,
                        xml_escape::text(s)
                    ));
                }
                Some(FormulaResult::Boolean(b)) => {
                    out.push_str(&format!("<c r=\"{}\" t=\"b\"", cell_ref));
                    if let Some(s) = cell.style_id {
                        out.push_str(&format!(" s=\"{}\"", s));
                    }
                    let bval = if *b { 1 } else { 0 };
                    out.push_str(&format!(
                        "><f>{}</f><v>{}</v></c>",
                        escaped_expr, bval
                    ));
                }
            }
        }

        WriteCellValue::DateSerial(f) => {
            // Same serialization as Number — no type attribute
            out.push_str(&format!("<c r=\"{}\"", cell_ref));
            if let Some(s) = cell.style_id {
                out.push_str(&format!(" s=\"{}\"", s));
            }
            out.push_str(&format!("><v>{}</v></c>", format_number(*f)));
        }

        WriteCellValue::InlineRichText(runs) => {
            // Sprint Ι Pod-α: emit `<c t="inlineStr"><is>...</is></c>`
            // so the SST never gets touched (matches openpyxl's
            // rich-text emit path verbatim).
            out.push_str(&format!("<c r=\"{}\" t=\"inlineStr\"", cell_ref));
            if let Some(s) = cell.style_id {
                out.push_str(&format!(" s=\"{}\"", s));
            }
            out.push_str("><is>");
            out.push_str(&crate::rich_text::emit_runs(runs));
            out.push_str("</is></c>");
        }
    }
}

/// Emit `<mergeCells count="N">…</mergeCells>`.
fn emit_merges(out: &mut String, sheet: &Worksheet) {
    // Sort by (top_row, left_col) ascending
    let mut merges = sheet.merges.clone();
    merges.sort_by(|a, b| a.top_row.cmp(&b.top_row).then(a.left_col.cmp(&b.left_col)));

    out.push_str(&format!("<mergeCells count=\"{}\">", merges.len()));

    for merge in &merges {
        let range = refs::format_range(
            (merge.top_row, merge.left_col),
            (merge.bottom_row, merge.right_col),
        );
        out.push_str(&format!("<mergeCell ref=\"{}\"/>", range));
    }

    out.push_str("</mergeCells>");
}

/// Emit `<hyperlinks>…</hyperlinks>`.
fn emit_hyperlinks(out: &mut String, sheet: &Worksheet) {
    // Calculate the starting rId for external hyperlinks
    // Convention: comments get rId1+rId2, then tables, then external hyperlinks
    let comments_offset: u32 = if !sheet.comments.is_empty() { 2 } else { 0 };
    let tables_offset: u32 = sheet.tables.len() as u32;
    let mut rid = comments_offset + tables_offset + 1; // first ext hyperlink rId

    out.push_str("<hyperlinks>");

    for (cell_ref, hyperlink) in &sheet.hyperlinks {
        if hyperlink.is_internal {
            // Internal target: ``location`` is the bare ``Sheet2!A1`` form
            // (no leading ``#``). Source of truth is the field, NOT a string
            // prefix on ``target`` — see model::worksheet::Hyperlink docs.
            out.push_str(&format!(
                "<hyperlink ref=\"{}\" location=\"{}\"",
                xml_escape::attr(cell_ref),
                xml_escape::attr(&hyperlink.target)
            ));
        } else {
            out.push_str(&format!(
                "<hyperlink ref=\"{}\" r:id=\"rId{}\"",
                xml_escape::attr(cell_ref),
                rid
            ));
            rid += 1;
        }

        if let Some(display) = &hyperlink.display {
            out.push_str(&format!(" display=\"{}\"", xml_escape::attr(display)));
        }
        if let Some(tooltip) = &hyperlink.tooltip {
            out.push_str(&format!(" tooltip=\"{}\"", xml_escape::attr(tooltip)));
        }

        out.push_str("/>");
    }

    out.push_str("</hyperlinks>");
}

/// Format an f64 for emission in a `<v>` element.
/// If the value is a whole number that fits in i64, emit it without decimal.
fn format_number(n: f64) -> String {
    if n == (n as i64) as f64 {
        format!("{}", n as i64)
    } else {
        format!("{}", n)
    }
}

/// Format an f64 for attribute values (widths, heights).
/// Uses Rust's default Display which drops trailing zeros.
fn format_f64(n: f64) -> String {
    // If whole number, emit without decimal point for cleanliness
    if n == (n as i64) as f64 && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        format!("{}", n)
    }
}

// ---------------------------------------------------------------------------
// Wave 3 extension-point helpers
//
// Each helper guards on the relevant collection being non-empty so the
// default output stays byte-identical to Wave 2 for sheets that don't use
// the feature. Wave 3 agents fill in the bodies (scoped strictly to what's
// inside the `if` block) — the call sites in `emit()` do not move.
// ---------------------------------------------------------------------------

/// Emit `<conditionalFormatting>` elements between `</mergeCells>` and
/// `<hyperlinks>`. Filled by W3C.
fn emit_conditional_formats(out: &mut String, sheet: &Worksheet) {
    use crate::model::conditional::{CellIsOperator, ConditionalKind};
    use std::collections::BTreeSet;

    if !sheet.conditional_formats.is_empty() {
        // R4 (W4E): track which stub variants we silently dropped per call so
        // the user sees one warning per variant per sheet emit instead of a
        // silent no-op. The wildcard arm below names each variant explicitly
        // so future enum additions surface as a compiler error here, not a
        // silent third-arm-of-the-fork.
        // TODO(R4): replace with structured diagnostics + GH issue link once
        // the CF expansion wave lands. See plan W4E.R4.
        let mut dropped: BTreeSet<&'static str> = BTreeSet::new();

        for cf in &sheet.conditional_formats {
            // Buffer rule XML first; only emit the wrapper if at least one rule
            // produced output. Otherwise we'd emit an empty
            // `<conditionalFormatting sqref="..."></conditionalFormatting>`,
            // which Excel treats as invalid and "repairs" on open.
            let mut rules_buf = String::new();

            for (priority_0, rule) in cf.rules.iter().enumerate() {
                let priority = priority_0 + 1;

                match &rule.kind {
                    ConditionalKind::CellIs { operator, formula_a, formula_b } => {
                        let op_str = match operator {
                            CellIsOperator::Equal => "equal",
                            CellIsOperator::NotEqual => "notEqual",
                            CellIsOperator::GreaterThan => "greaterThan",
                            CellIsOperator::GreaterThanOrEqual => "greaterThanOrEqual",
                            CellIsOperator::LessThan => "lessThan",
                            CellIsOperator::LessThanOrEqual => "lessThanOrEqual",
                            CellIsOperator::Between => "between",
                            CellIsOperator::NotBetween => "notBetween",
                        };

                        rules_buf.push_str(&format!(
                            "<cfRule type=\"cellIs\" priority=\"{}\" operator=\"{}\"",
                            priority, op_str
                        ));
                        if let Some(dxf_id) = rule.dxf_id {
                            rules_buf.push_str(&format!(" dxfId=\"{}\"", dxf_id));
                        }
                        if rule.stop_if_true {
                            rules_buf.push_str(" stopIfTrue=\"1\"");
                        }
                        rules_buf.push('>');
                        rules_buf.push_str(&format!("<formula>{}</formula>", xml_escape::text(formula_a)));
                        let needs_second = matches!(operator, CellIsOperator::Between | CellIsOperator::NotBetween);
                        if needs_second {
                            if let Some(fb) = formula_b {
                                rules_buf.push_str(&format!("<formula>{}</formula>", xml_escape::text(fb)));
                            }
                        }
                        rules_buf.push_str("</cfRule>");
                    }

                    ConditionalKind::Expression { formula } => {
                        rules_buf.push_str(&format!(
                            "<cfRule type=\"expression\" priority=\"{}\"",
                            priority
                        ));
                        if let Some(dxf_id) = rule.dxf_id {
                            rules_buf.push_str(&format!(" dxfId=\"{}\"", dxf_id));
                        }
                        if rule.stop_if_true {
                            rules_buf.push_str(" stopIfTrue=\"1\"");
                        }
                        rules_buf.push('>');
                        rules_buf.push_str(&format!("<formula>{}</formula>", xml_escape::text(formula)));
                        rules_buf.push_str("</cfRule>");
                    }

                    ConditionalKind::DataBar { color_rgb, min, max } => {
                        rules_buf.push_str(&format!(
                            "<cfRule type=\"dataBar\" priority=\"{}\">",
                            priority
                        ));
                        rules_buf.push_str("<dataBar>");
                        emit_cfvo(&mut rules_buf, min);
                        emit_cfvo(&mut rules_buf, max);
                        rules_buf.push_str(&format!("<color rgb=\"{}\"/>", color_rgb));
                        rules_buf.push_str("</dataBar>");
                        rules_buf.push_str("</cfRule>");
                    }

                    ConditionalKind::ColorScale { stops } => {
                        rules_buf.push_str(&format!(
                            "<cfRule type=\"colorScale\" priority=\"{}\">",
                            priority
                        ));
                        rules_buf.push_str("<colorScale>");
                        // All cfvo elements first
                        for stop in stops {
                            emit_cfvo(&mut rules_buf, &stop.threshold);
                        }
                        // Then all color elements
                        for stop in stops {
                            rules_buf.push_str(&format!("<color rgb=\"{}\"/>", stop.color_rgb));
                        }
                        rules_buf.push_str("</colorScale>");
                        rules_buf.push_str("</cfRule>");
                    }

                    // Stub variants — Excel would reject the synthetic XML
                    // we'd produce for these, so skip them and remember which
                    // names were dropped. One eprintln! per variant per sheet
                    // emit (deduped via the BTreeSet above).
                    ConditionalKind::ContainsText { .. } => {
                        dropped.insert("ContainsText");
                        continue;
                    }
                    ConditionalKind::NotContainsText { .. } => {
                        dropped.insert("NotContainsText");
                        continue;
                    }
                    ConditionalKind::BeginsWith { .. } => {
                        dropped.insert("BeginsWith");
                        continue;
                    }
                    ConditionalKind::EndsWith { .. } => {
                        dropped.insert("EndsWith");
                        continue;
                    }
                    ConditionalKind::Duplicate => {
                        dropped.insert("Duplicate");
                        continue;
                    }
                    ConditionalKind::Unique => {
                        dropped.insert("Unique");
                        continue;
                    }
                    ConditionalKind::Top10 { .. } => {
                        dropped.insert("Top10");
                        continue;
                    }
                    ConditionalKind::AboveAverage { .. } => {
                        dropped.insert("AboveAverage");
                        continue;
                    }
                    ConditionalKind::IconSet { .. } => {
                        dropped.insert("IconSet");
                        continue;
                    }
                }
            }

            if !rules_buf.is_empty() {
                out.push_str(&format!(
                    "<conditionalFormatting sqref=\"{}\">",
                    xml_escape::attr(&cf.sqref)
                ));
                out.push_str(&rules_buf);
                out.push_str("</conditionalFormatting>");
            }
        }

        if !dropped.is_empty() {
            let names: Vec<&str> = dropped.iter().copied().collect();
            eprintln!(
                "wolfxl-writer: dropped {} conditional-format rule kind{} on sheet {:?} \
                 (variants: {}). Wave 3 ships only CellIs/Expression/DataBar/ColorScale; \
                 other kinds are pending a future CF expansion wave.",
                names.len(),
                if names.len() == 1 { "" } else { "s" },
                sheet.name,
                names.join(", "),
            );
        }
    }
}

/// Emit `<dataValidations count="N">…</dataValidations>` between the CF
/// block and `<hyperlinks>`. Filled by W3C.
fn emit_data_validations(out: &mut String, sheet: &Worksheet) {
    use crate::model::validation::{ErrorStyle, ValidationType, ValidationOperator};

    if !sheet.validations.is_empty() {
        out.push_str(&format!(
            "<dataValidations count=\"{}\">",
            sheet.validations.len()
        ));

        for dv in &sheet.validations {
            let type_str = match dv.validation_type {
                ValidationType::Any => "any",
                ValidationType::Whole => "whole",
                ValidationType::Decimal => "decimal",
                ValidationType::List => "list",
                ValidationType::Date => "date",
                ValidationType::Time => "time",
                ValidationType::TextLength => "textLength",
                ValidationType::Custom => "custom",
            };

            out.push_str(&format!("<dataValidation type=\"{}\"", type_str));

            // operator — omit for List and Custom
            let needs_operator = !matches!(
                dv.validation_type,
                ValidationType::List | ValidationType::Custom
            );
            if needs_operator {
                let op_str = match dv.operator {
                    ValidationOperator::Between => "between",
                    ValidationOperator::NotBetween => "notBetween",
                    ValidationOperator::Equal => "equal",
                    ValidationOperator::NotEqual => "notEqual",
                    ValidationOperator::GreaterThan => "greaterThan",
                    ValidationOperator::LessThan => "lessThan",
                    ValidationOperator::GreaterThanOrEqual => "greaterThanOrEqual",
                    ValidationOperator::LessThanOrEqual => "lessThanOrEqual",
                };
                out.push_str(&format!(" operator=\"{}\"", op_str));
            }

            if dv.allow_blank {
                out.push_str(" allowBlank=\"1\"");
            }

            // showDropDown — note: capital D in OOXML attribute name
            if dv.show_dropdown {
                out.push_str(" showDropDown=\"1\"");
            }

            if dv.show_input_message {
                out.push_str(" showInputMessage=\"1\"");
            }

            if dv.show_error_message {
                out.push_str(" showErrorMessage=\"1\"");
            }

            // errorStyle — only emit when not default (Stop)
            match dv.error_style {
                ErrorStyle::Stop => {}
                ErrorStyle::Warning => {
                    out.push_str(" errorStyle=\"warning\"");
                }
                ErrorStyle::Information => {
                    out.push_str(" errorStyle=\"information\"");
                }
            }

            if let Some(ref title) = dv.error_title {
                out.push_str(&format!(" errorTitle=\"{}\"", xml_escape::attr(title)));
            }
            if let Some(ref msg) = dv.error_message {
                out.push_str(&format!(" error=\"{}\"", xml_escape::attr(msg)));
            }
            if let Some(ref title) = dv.input_title {
                out.push_str(&format!(" promptTitle=\"{}\"", xml_escape::attr(title)));
            }
            if let Some(ref msg) = dv.input_message {
                out.push_str(&format!(" prompt=\"{}\"", xml_escape::attr(msg)));
            }

            out.push_str(&format!(" sqref=\"{}\">", xml_escape::attr(&dv.sqref)));

            // formula1 — only when formula_a is Some
            if let Some(ref fa) = dv.formula_a {
                out.push_str(&format!("<formula1>{}</formula1>", xml_escape::text(fa)));
            }

            // formula2 — only when formula_b is Some AND operator is Between or NotBetween
            let is_between = matches!(
                dv.operator,
                ValidationOperator::Between | ValidationOperator::NotBetween
            );
            if is_between {
                if let Some(ref fb) = dv.formula_b {
                    out.push_str(&format!("<formula2>{}</formula2>", xml_escape::text(fb)));
                }
            }

            out.push_str("</dataValidation>");
        }

        out.push_str("</dataValidations>");
    }
}

fn emit_cfvo(out: &mut String, threshold: &crate::model::conditional::ConditionalThreshold) {
    use crate::model::conditional::ConditionalThreshold;
    match threshold {
        ConditionalThreshold::Min => {
            out.push_str("<cfvo type=\"min\"/>");
        }
        ConditionalThreshold::Max => {
            out.push_str("<cfvo type=\"max\"/>");
        }
        ConditionalThreshold::Number(x) => {
            out.push_str(&format!("<cfvo type=\"num\" val=\"{}\"/>", format_f64(*x)));
        }
        ConditionalThreshold::Percent(x) => {
            out.push_str(&format!("<cfvo type=\"percent\" val=\"{}\"/>", format_f64(*x)));
        }
        ConditionalThreshold::Percentile(x) => {
            out.push_str(&format!("<cfvo type=\"percentile\" val=\"{}\"/>", format_f64(*x)));
        }
        ConditionalThreshold::Formula(s) => {
            out.push_str(&format!("<cfvo type=\"formula\" val=\"{}\"/>", xml_escape::attr(s)));
        }
    }
}

/// Emit `<legacyDrawing r:id="rId2"/>` when comments exist. The rId is a
/// hardcoded convention (see `rels::emit_sheet`): with comments present,
/// rId1 → commentsN.xml and rId2 → vmlDrawingN.vml. Filled by W3A.
fn emit_legacy_drawing(out: &mut String, sheet: &Worksheet) {
    if !sheet.comments.is_empty() {
        out.push_str("<legacyDrawing r:id=\"rId2\"/>");
    }
}

/// Sprint Λ Pod-β + Sprint Μ Pod-α — emit `<drawing r:id="rIdN"/>`
/// when the sheet has at least one image or chart. The rId is
/// allocated at the END of the sheet's rels graph: comments offset (2
/// if any comments) + table count + external-hyperlink count + 1.
/// This mirrors the allocation in `rels::emit_sheet` so the rId
/// numbering stays in lock-step.
fn emit_drawing_ref(out: &mut String, sheet: &Worksheet) {
    if sheet.images.is_empty() && sheet.charts.is_empty() {
        return;
    }
    let comments_offset: u32 = if !sheet.comments.is_empty() { 2 } else { 0 };
    let table_count = sheet.tables.len() as u32;
    let external_hyperlinks = sheet
        .hyperlinks
        .values()
        .filter(|h| !h.is_internal)
        .count() as u32;
    let rid = comments_offset + table_count + external_hyperlinks + 1;
    out.push_str(&format!("<drawing r:id=\"rId{rid}\"/>"));
}

/// Emit `<tableParts count="N">…<tablePart r:id="rIdX"/>…</tableParts>`.
/// rId starts after comments (offset = 2 iff comments exist, else 0),
/// one rId per table in sheet-local order. Filled by W3B.
fn emit_table_parts(out: &mut String, sheet: &Worksheet) {
    if !sheet.tables.is_empty() {
        let comments_offset: u32 = if !sheet.comments.is_empty() { 2 } else { 0 };
        out.push_str(&format!("<tableParts count=\"{}\">", sheet.tables.len()));
        for (local_idx, _) in sheet.tables.iter().enumerate() {
            let rid = comments_offset + local_idx as u32 + 1;
            out.push_str(&format!("<tablePart r:id=\"rId{}\"/>", rid));
        }
        out.push_str("</tableParts>");
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::cell::{FormulaResult, WriteCell, WriteCellValue};
    use crate::model::worksheet::{Column, FreezePane, Hyperlink, Merge, SplitPane, Worksheet};
    use quick_xml::events::Event;
    use quick_xml::Reader;

    fn parse_ok(bytes: &[u8]) {
        let text = std::str::from_utf8(bytes).expect("utf8");
        let mut reader = Reader::from_str(text);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Eof) => break,
                Err(e) => panic!("XML parse error: {e}"),
                _ => {}
            }
            buf.clear();
        }
    }

    fn emit_sheet(sheet: &Worksheet, sheet_idx: u32) -> (Vec<u8>, SstBuilder) {
        let mut sst = SstBuilder::default();
        let styles = crate::model::format::StylesBuilder::default();
        let bytes = emit(sheet, sheet_idx, &mut sst, &styles);
        (bytes, sst)
    }

    // --- 1. Empty sheet ---

    #[test]
    fn empty_sheet_emits_minimal_valid_xml() {
        let sheet = Worksheet::new("X");
        let (bytes, _sst) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("<dimension ref=\"A1\"/>"), "dimension: {text}");
        // sheetData should be empty self-close or open+close
        assert!(
            text.contains("<sheetData/>") || text.contains("<sheetData></sheetData>"),
            "empty sheetData: {text}"
        );
    }

    // --- 2. Blank cell with style ---

    #[test]
    fn blank_cell_with_style_emits_self_closing() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::Blank).with_style(3));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("<c r=\"A1\" s=\"3\"/>"), "blank+style: {text}");
    }

    // --- 3. Blank cell without style is skipped ---

    #[test]
    fn blank_cell_without_style_is_skipped() {
        let mut sheet = Worksheet::new("S");
        // Unstyled blank at A1 — should be omitted
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::Blank));
        // Another cell at B1 to ensure the row is emitted
        sheet.set_cell(1, 2, WriteCell::new(WriteCellValue::Number(5.0)));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(!text.contains("<c r=\"A1\""), "should not have A1: {text}");
        assert!(text.contains("<c r=\"B1\""), "should have B1: {text}");
    }

    // --- 4. Number whole emits as integer ---

    #[test]
    fn number_whole_emits_as_integer() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::Number(42.0)));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("<v>42</v>"), "integer value: {text}");
        assert!(!text.contains("<v>42.0</v>"), "should not have decimal: {text}");
    }

    // --- 5. Number float ---

    #[test]
    fn number_float_emits_as_float() {
        let mut sheet = Worksheet::new("S");
        // Use 1.5 — a non-integer that is not an approximation of a math constant
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::Number(1.5)));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("<v>1.5</v>"), "float value: {text}");
    }

    // --- 6. Negative number ---

    #[test]
    fn number_negative() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::Number(-17.5)));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("<v>-17.5</v>"), "negative: {text}");
    }

    // --- 7. String interns into SST ---

    #[test]
    fn string_interns_into_sst_and_emits_index() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::String("hello".into())));
        let (bytes, sst) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);

        assert_eq!(sst.total_count(), 1);
        assert_eq!(sst.unique_count(), 1);

        let text = String::from_utf8(bytes).unwrap();
        // Should emit t="s" with SST index 0
        assert!(text.contains("<c r=\"A1\" t=\"s\">"), "t=s attribute: {text}");
        assert!(text.contains("<v>0</v>"), "sst index: {text}");
    }

    // --- 8. Multiple strings in insertion order ---

    #[test]
    fn strings_intern_in_insertion_order() {
        let mut sheet = Worksheet::new("S");
        // Row 1: beta, alpha, beta
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::String("beta".into())));
        sheet.set_cell(2, 1, WriteCell::new(WriteCellValue::String("alpha".into())));
        sheet.set_cell(3, 1, WriteCell::new(WriteCellValue::String("beta".into())));
        let (bytes, sst) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);

        assert_eq!(sst.total_count(), 3);
        assert_eq!(sst.unique_count(), 2);

        // beta=0, alpha=1
        let collected: Vec<(u32, &str)> = sst.iter().collect();
        assert_eq!(collected[0], (0, "beta"));
        assert_eq!(collected[1], (1, "alpha"));

        let text = String::from_utf8(bytes).unwrap();
        // A1 (row1,col1) = beta = index 0
        assert!(text.contains("<c r=\"A1\" t=\"s\"><v>0</v></c>"), "A1 beta=0: {text}");
        // A2 (row2,col1) = alpha = index 1
        assert!(text.contains("<c r=\"A2\" t=\"s\"><v>1</v></c>"), "A2 alpha=1: {text}");
        // A3 (row3,col1) = beta = index 0 (deduped)
        assert!(text.contains("<c r=\"A3\" t=\"s\"><v>0</v></c>"), "A3 beta=0: {text}");
    }

    // --- 9. Boolean true and false ---

    #[test]
    fn boolean_true_and_false() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::Boolean(true)));
        sheet.set_cell(1, 2, WriteCell::new(WriteCellValue::Boolean(false)));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // A1 = true
        assert!(
            text.contains("<c r=\"A1\" t=\"b\"><v>1</v></c>"),
            "bool true: {text}"
        );
        // B1 = false
        assert!(
            text.contains("<c r=\"B1\" t=\"b\"><v>0</v></c>"),
            "bool false: {text}"
        );
    }

    // --- 10. Formula without result emits stub <v>0</v> ---

    #[test]
    fn formula_without_result_emits_stub_zero() {
        // Without a stub <v>, calamine and xlsx2csv show None for every
        // formula cell. rust_xlsxwriter writes <v>0</v> for the same reason;
        // we mirror that so read-back paths keep working.
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(
            5,
            3,
            WriteCell::new(WriteCellValue::Formula {
                expr: "SUM(A1:A10)".into(),
                result: None,
            }),
        );
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<f>SUM(A1:A10)</f><v>0</v>"),
            "formula+stub-v: {text}"
        );
    }

    // --- 11. Formula with numeric result ---

    #[test]
    fn formula_with_numeric_result() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(
            1,
            1,
            WriteCell::new(WriteCellValue::Formula {
                expr: "1+6".into(),
                result: Some(FormulaResult::Number(7.0)),
            }),
        );
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("<f>1+6</f><v>7</v>"), "formula+numeric result: {text}");
    }

    // --- 12. Formula with string result uses t="str" ---

    #[test]
    fn formula_with_string_result_uses_t_str() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(
            1,
            1,
            WriteCell::new(WriteCellValue::Formula {
                expr: "CONCAT(\"fo\",\"o\")".into(),
                result: Some(FormulaResult::String("foo".into())),
            }),
        );
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("t=\"str\""), "t=str: {text}");
        assert!(text.contains("<v>foo</v>"), "string result: {text}");
    }

    // --- 13. Formula with boolean result ---

    #[test]
    fn formula_with_boolean_result() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(
            1,
            1,
            WriteCell::new(WriteCellValue::Formula {
                expr: "TRUE()".into(),
                result: Some(FormulaResult::Boolean(true)),
            }),
        );
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("t=\"b\""), "t=b: {text}");
        assert!(text.contains("<v>1</v>"), "bool result 1: {text}");
    }

    // --- 14. DateSerial emits as number ---

    #[test]
    fn dateserial_emits_as_number() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::DateSerial(44927.5)));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("<v>44927.5</v>"), "date serial: {text}");
        // No t attribute for dates
        assert!(!text.contains("t=\"s\""), "no t=s for date: {text}");
        assert!(!text.contains("t=\"b\""), "no t=b for date: {text}");
    }

    // --- 15. Cell style_id emits s attribute ---

    #[test]
    fn cell_style_id_emits_s_attribute() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::Number(1.0)).with_style(5));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("s=\"5\""), "s=5 attribute: {text}");
    }

    // --- 16. Cell without style omits s attribute ---

    #[test]
    fn cell_without_style_omits_s_attribute() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::Number(1.0)));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // No s= at all (not even s="0")
        let cell_start = text.find("<c r=\"A1\"").expect("A1 cell");
        let cell_end = text[cell_start..].find(">").expect(">") + cell_start;
        let tag = &text[cell_start..=cell_end];
        assert!(!tag.contains("s="), "no s attr when no style: {tag}");
    }

    // --- 17. Merges sorted ascending ---

    #[test]
    fn merges_sorted_ascending() {
        let mut sheet = Worksheet::new("S");
        // Add in reverse order
        sheet.merge(Merge {
            top_row: 3,
            left_col: 3,
            bottom_row: 4,
            right_col: 4,
        });
        sheet.merge(Merge {
            top_row: 1,
            left_col: 1,
            bottom_row: 2,
            right_col: 2,
        });
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        let pos_a1 = text.find("ref=\"A1:B2\"").expect("A1:B2");
        let pos_c3 = text.find("ref=\"C3:D4\"").expect("C3:D4");
        assert!(pos_a1 < pos_c3, "A1:B2 should come before C3:D4: {text}");
    }

    // --- 18. Merges element omitted when empty ---

    #[test]
    fn merges_element_omitted_when_empty() {
        let sheet = Worksheet::new("S");
        let (bytes, _) = emit_sheet(&sheet, 0);
        let text = String::from_utf8(bytes).unwrap();
        assert!(!text.contains("<mergeCells"), "no mergeCells when none: {text}");
    }

    // --- 19. Freeze rows only ---

    #[test]
    fn freeze_rows_only() {
        let mut sheet = Worksheet::new("S");
        sheet.freeze = Some(FreezePane {
            freeze_row: 3,
            freeze_col: 0,
            top_left: None,
        });
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // OOXML: ySplit is the COUNT of frozen rows (= freeze_row - 1).
        assert!(text.contains("ySplit=\"2\""), "ySplit: {text}");
        assert!(text.contains("state=\"frozen\""), "state=frozen: {text}");
        assert!(text.contains("activePane=\"bottomLeft\""), "activePane: {text}");
        assert!(!text.contains("xSplit"), "no xSplit: {text}");
    }

    // --- 20. Freeze cols only ---

    #[test]
    fn freeze_cols_only() {
        let mut sheet = Worksheet::new("S");
        sheet.freeze = Some(FreezePane {
            freeze_row: 0,
            freeze_col: 2,
            top_left: None,
        });
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // OOXML: xSplit is the COUNT of frozen columns (= freeze_col - 1).
        assert!(text.contains("xSplit=\"1\""), "xSplit: {text}");
        assert!(text.contains("state=\"frozen\""), "state=frozen: {text}");
        assert!(text.contains("activePane=\"topRight\""), "activePane: {text}");
        assert!(!text.contains("ySplit"), "no ySplit: {text}");
    }

    // --- 21. Freeze both ---

    #[test]
    fn freeze_both() {
        let mut sheet = Worksheet::new("S");
        sheet.freeze = Some(FreezePane {
            freeze_row: 2,
            freeze_col: 3,
            top_left: None,
        });
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // OOXML counts: freeze_col=3 -> xSplit=2, freeze_row=2 -> ySplit=1.
        assert!(text.contains("xSplit=\"2\""), "xSplit: {text}");
        assert!(text.contains("ySplit=\"1\""), "ySplit: {text}");
        assert!(text.contains("state=\"frozen\""), "state=frozen: {text}");
        assert!(text.contains("activePane=\"bottomRight\""), "activePane: {text}");
    }

    // --- 21a. B2 freeze emits count one (W4-polish regression) ---

    #[test]
    fn emit_freeze_pane_b2_emits_count_one() {
        let mut sheet = Worksheet::new("S");
        sheet.freeze = Some(FreezePane {
            freeze_row: 2,
            freeze_col: 2,
            top_left: None,
        });
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("xSplit=\"1\""), "xSplit count: {text}");
        assert!(text.contains("ySplit=\"1\""), "ySplit count: {text}");
        assert!(text.contains("topLeftCell=\"B2\""), "topLeftCell: {text}");
        assert!(text.contains("activePane=\"bottomRight\""), "activePane: {text}");
        // Negative: must NOT emit the cell coordinate as the count.
        assert!(!text.contains("xSplit=\"2\""), "xSplit must not be 2: {text}");
        assert!(!text.contains("ySplit=\"2\""), "ySplit must not be 2: {text}");
    }

    // --- 21b. C5 freeze emits asymmetric counts ---

    #[test]
    fn emit_freeze_pane_c5_emits_correct_counts() {
        let mut sheet = Worksheet::new("S");
        sheet.freeze = Some(FreezePane {
            freeze_row: 5,
            freeze_col: 3,
            top_left: None,
        });
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("xSplit=\"2\""), "xSplit: {text}");
        assert!(text.contains("ySplit=\"4\""), "ySplit: {text}");
        assert!(text.contains("topLeftCell=\"C5\""), "topLeftCell: {text}");
    }

    // --- 21c. A1 freeze is a no-op (degenerate) ---

    #[test]
    fn emit_freeze_pane_a1_is_no_op() {
        let mut sheet = Worksheet::new("S");
        sheet.freeze = Some(FreezePane {
            freeze_row: 1,
            freeze_col: 1,
            top_left: None,
        });
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // Both splits collapse to zero — emit no <pane> at all.
        assert!(!text.contains("<pane"), "must not emit pane for A1 freeze: {text}");
    }

    // --- 22. Split pane is not frozen ---

    #[test]
    fn split_pane_is_not_frozen() {
        let mut sheet = Worksheet::new("S");
        sheet.split = Some(SplitPane {
            x_split: 1200.0,
            y_split: 600.0,
            top_left: None,
        });
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(!text.contains("state=\"frozen\""), "no frozen for split: {text}");
        assert!(text.contains("<pane"), "has pane: {text}");
    }

    // --- 23. Columns emit single min/max ---

    #[test]
    fn columns_emit_single_min_max() {
        let mut sheet = Worksheet::new("S");
        sheet.set_column(
            3,
            Column {
                width: Some(12.5),
                ..Default::default()
            },
        );
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<col min=\"3\" max=\"3\" width=\"12.5\" customWidth=\"1\"/>"),
            "col width: {text}"
        );
    }

    // --- 24. Columns hidden and outline ---

    #[test]
    fn columns_hidden_and_outline() {
        let mut sheet = Worksheet::new("S");
        sheet.set_column(
            3,
            Column {
                width: None,
                hidden: true,
                outline_level: 2,
                style_id: None,
            },
        );
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<col min=\"3\" max=\"3\" hidden=\"1\" outlineLevel=\"2\"/>"),
            "col hidden+outline: {text}"
        );
        assert!(!text.contains("width="), "no width when none: {text}");
        assert!(!text.contains("customWidth="), "no customWidth when none: {text}");
    }

    // --- 25. Columns with style_id ---

    #[test]
    fn columns_with_style_id() {
        let mut sheet = Worksheet::new("S");
        sheet.set_column(
            2,
            Column {
                width: Some(10.0),
                style_id: Some(4),
                ..Default::default()
            },
        );
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("style=\"4\" customFormat=\"1\""), "style+customFormat: {text}");
    }

    // --- 26. External hyperlink gets external rId ---

    #[test]
    fn external_hyperlink_gets_external_rid() {
        let mut sheet = Worksheet::new("S");
        sheet.hyperlinks.insert(
            "A1".to_string(),
            Hyperlink {
                target: "https://ex.com".into(),
                is_internal: false,
                display: None,
                tooltip: None,
            },
        );
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // No comments, no tables → rId1
        assert!(text.contains("r:id=\"rId1\""), "rId1 for ext hyperlink: {text}");
    }

    // --- 27. External hyperlink with comments starts at rId3 ---

    #[test]
    fn external_hyperlink_with_comments_starts_at_rid3() {
        use crate::model::comment::Comment;
        let mut sheet = Worksheet::new("S");
        // Add a comment so comments_offset = 2
        sheet.comments.insert(
            "A1".to_string(),
            Comment {
                author_id: 0,
                text: "Note".into(),
                width_pt: None,
                height_pt: None,
                visible: false,
            },
        );
        sheet.hyperlinks.insert(
            "B1".to_string(),
            Hyperlink {
                target: "https://ex.com".into(),
                is_internal: false,
                display: None,
                tooltip: None,
            },
        );
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // comments offset=2, tables=0, so first ext hyperlink = rId3
        assert!(text.contains("r:id=\"rId3\""), "rId3 with comments: {text}");
    }

    // --- 28. Internal hyperlink uses location, no r:id ---

    #[test]
    fn internal_hyperlink_uses_location_no_rid() {
        let mut sheet = Worksheet::new("S");
        sheet.hyperlinks.insert(
            "A1".to_string(),
            Hyperlink {
                target: "Sheet2!A1".into(),
                is_internal: true,
                display: None,
                tooltip: None,
            },
        );
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("location=\"Sheet2!A1\""),
            "location attr: {text}"
        );
        assert!(!text.contains("r:id="), "no r:id for internal: {text}");
    }

    // --- 29. Dimension tracks max populated ---

    #[test]
    fn dimension_tracks_max_populated() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::Number(1.0)));
        sheet.set_cell(10, 4, WriteCell::new(WriteCellValue::Number(2.0)));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<dimension ref=\"A1:D10\"/>"),
            "dimension: {text}"
        );
    }

    // --- 30. First sheet has tabSelected ---

    #[test]
    fn first_sheet_has_tab_selected() {
        let sheet = Worksheet::new("S");
        let (bytes_first, _) = emit_sheet(&sheet, 0);
        let text_first = String::from_utf8(bytes_first).unwrap();
        assert!(
            text_first.contains("tabSelected=\"1\""),
            "first sheet tabSelected: {text_first}"
        );

        let (bytes_second, _) = emit_sheet(&sheet, 1);
        let text_second = String::from_utf8(bytes_second).unwrap();
        assert!(
            !text_second.contains("tabSelected"),
            "second sheet no tabSelected: {text_second}"
        );
    }

    // --- 31. Kitchen-sink well-formed ---

    #[test]
    fn xml_well_formed_under_quick_xml() {
        let mut sheet = Worksheet::new("Kitchen");

        // Cells of each type
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::Number(42.0)));
        sheet.set_cell(1, 2, WriteCell::new(WriteCellValue::String("hello".into())));
        sheet.set_cell(1, 3, WriteCell::new(WriteCellValue::Boolean(true)));
        sheet.set_cell(
            2,
            1,
            WriteCell::new(WriteCellValue::Formula {
                expr: "SUM(A1:A1)".into(),
                result: Some(FormulaResult::Number(42.0)),
            }),
        );
        sheet.set_cell(2, 2, WriteCell::new(WriteCellValue::DateSerial(44927.0)));
        sheet.set_cell(
            2,
            3,
            WriteCell::new(WriteCellValue::Blank).with_style(1),
        );

        // Merge
        sheet.merge(Merge {
            top_row: 3,
            left_col: 1,
            bottom_row: 3,
            right_col: 2,
        });

        // Freeze
        sheet.freeze = Some(FreezePane {
            freeze_row: 1,
            freeze_col: 0,
            top_left: None,
        });

        // Column
        sheet.set_column(
            1,
            Column {
                width: Some(15.0),
                ..Default::default()
            },
        );

        // Hyperlinks (both internal and external)
        sheet.hyperlinks.insert(
            "A1".to_string(),
            Hyperlink {
                target: "https://example.com".into(),
                is_internal: false,
                display: Some("Example".into()),
                tooltip: None,
            },
        );
        sheet.hyperlinks.insert(
            "B1".to_string(),
            Hyperlink {
                target: "Sheet2!A1".into(),
                is_internal: true,
                display: None,
                tooltip: Some("Go to sheet2".into()),
            },
        );

        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
    }

    // --- 32. Dimension excludes unstyled blank cells ---

    #[test]
    fn dimension_excludes_unstyled_blank_cells() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(10, 5, WriteCell::new(WriteCellValue::Blank));
        let mut sst = SstBuilder::default();
        let styles = crate::model::format::StylesBuilder::default();
        let bytes = emit(&sheet, 0, &mut sst, &styles);
        let text = String::from_utf8(bytes).unwrap();
        // The only cell is blank+unstyled — dimension should stay A1 (empty sheet).
        assert!(text.contains("<dimension ref=\"A1\"/>"), "got: {text}");
        // No <c> should be emitted either.
        assert!(!text.contains("<c r="), "no cell should be emitted: {text}");
    }

    // --- 33. Dimension includes styled blank cells ---

    #[test]
    fn dimension_includes_styled_blank_cells() {
        let mut sheet = Worksheet::new("S");
        let styled_blank = WriteCell::new(WriteCellValue::Blank).with_style(3);
        sheet.set_cell(10, 5, styled_blank);
        let mut sst = SstBuilder::default();
        let styles = crate::model::format::StylesBuilder::default();
        let bytes = emit(&sheet, 0, &mut sst, &styles);
        let text = String::from_utf8(bytes).unwrap();
        // E10 is the single populated cell — A1:E10 bounding box (or just E10 for single-cell).
        assert!(
            text.contains("<dimension ref=\"E10\"/>") || text.contains("<dimension ref=\"A1:E10\"/>"),
            "styled blank should still count toward dimension: {text}"
        );
        // The cell MUST emit because it has a style.
        assert!(text.contains("<c r=\"E10\" s=\"3\"/>"), "got: {text}");
    }

    // --- 34. legacyDrawing emitted when comments exist ---

    #[test]
    fn legacy_drawing_emitted_when_comments_exist() {
        use crate::model::comment::Comment;
        let mut sheet = Worksheet::new("S");
        sheet.comments.insert(
            "A1".to_string(),
            Comment {
                author_id: 0,
                text: "Note".into(),
                width_pt: None,
                height_pt: None,
                visible: false,
            },
        );
        let (bytes, _) = emit_sheet(&sheet, 0);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<legacyDrawing r:id=\"rId2\"/>"),
            "legacyDrawing with rId2: {text}"
        );
    }

    // --- 35. legacyDrawing absent when no comments ---

    #[test]
    fn legacy_drawing_absent_when_no_comments() {
        let sheet = Worksheet::new("S");
        let (bytes, _) = emit_sheet(&sheet, 0);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            !text.contains("<legacyDrawing"),
            "legacyDrawing must not appear without comments: {text}"
        );
    }

    // --- 36. table_parts_absent_when_no_tables ---

    #[test]
    fn table_parts_absent_when_no_tables() {
        let sheet = Worksheet::new("S");
        let (bytes, _) = emit_sheet(&sheet, 0);
        let text = String::from_utf8(bytes).unwrap();
        assert!(!text.contains("<tableParts"), "no tableParts when none: {text}");
    }

    // --- 37. table_parts_no_comments_starts_at_rid1 ---

    #[test]
    fn table_parts_no_comments_starts_at_rid1() {
        use crate::model::table::{Table, TableColumn};
        let mut sheet = Worksheet::new("S");
        sheet.tables.push(Table {
            name: "MyTable".into(),
            display_name: None,
            range: "A1:B10".into(),
            columns: vec![TableColumn {
                name: "C1".into(),
                totals_function: None,
                totals_label: None,
            }],
            header_row: true,
            totals_row: false,
            style: None,
            autofilter: true,
        });
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts>"),
            "rId1 with no comments: {text}"
        );
    }

    // --- 38. table_parts_with_comments_starts_at_rid3 ---

    #[test]
    fn table_parts_with_comments_starts_at_rid3() {
        use crate::model::comment::Comment;
        use crate::model::table::{Table, TableColumn};
        let mut sheet = Worksheet::new("S");
        // Add a comment — comments_offset = 2 (rId1=commentsN.xml, rId2=vmlDrawingN.vml)
        sheet.comments.insert(
            "A1".to_string(),
            Comment {
                author_id: 0,
                text: "Note".into(),
                width_pt: None,
                height_pt: None,
                visible: false,
            },
        );
        sheet.tables.push(Table {
            name: "MyTable".into(),
            display_name: None,
            range: "A1:B10".into(),
            columns: vec![TableColumn {
                name: "C1".into(),
                totals_function: None,
                totals_label: None,
            }],
            header_row: true,
            totals_row: false,
            style: None,
            autofilter: true,
        });
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // rId1=comments, rId2=VML, rId3=table[0]
        assert!(
            text.contains("<tableParts count=\"1\"><tablePart r:id=\"rId3\"/></tableParts>"),
            "rId3 with comments: {text}"
        );
    }

    // --- 39. table_parts_multiple_tables_no_comments_rids_sequential ---

    #[test]
    fn table_parts_multiple_tables_no_comments_rids_sequential() {
        use crate::model::table::{Table, TableColumn};
        let mut sheet = Worksheet::new("S");
        // Two tables on the same sheet, no comments → rIds start at 1 and run sequentially.
        sheet.tables.push(Table {
            name: "TableA".into(),
            display_name: None,
            range: "A1:B10".into(),
            columns: vec![TableColumn {
                name: "C1".into(),
                totals_function: None,
                totals_label: None,
            }],
            header_row: true,
            totals_row: false,
            style: None,
            autofilter: true,
        });
        sheet.tables.push(Table {
            name: "TableB".into(),
            display_name: None,
            range: "D1:E10".into(),
            columns: vec![TableColumn {
                name: "C1".into(),
                totals_function: None,
                totals_label: None,
            }],
            header_row: true,
            totals_row: false,
            style: None,
            autofilter: true,
        });
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // Two tables, no comments → rId1, rId2 in sheet-local order.
        assert!(
            text.contains("<tableParts count=\"2\"><tablePart r:id=\"rId1\"/><tablePart r:id=\"rId2\"/></tableParts>"),
            "rId1/rId2 sequential with no comments: {text}"
        );
    }

    // --- 40. table_parts_multiple_tables_with_comments_rids_offset ---

    #[test]
    fn table_parts_multiple_tables_with_comments_rids_offset() {
        use crate::model::comment::Comment;
        use crate::model::table::{Table, TableColumn};
        let mut sheet = Worksheet::new("S");
        // Comments → comments_offset = 2 (rId1=commentsN.xml, rId2=vmlDrawingN.vml)
        sheet.comments.insert(
            "A1".to_string(),
            Comment {
                author_id: 0,
                text: "Note".into(),
                width_pt: None,
                height_pt: None,
                visible: false,
            },
        );
        sheet.tables.push(Table {
            name: "TableA".into(),
            display_name: None,
            range: "A1:B10".into(),
            columns: vec![TableColumn {
                name: "C1".into(),
                totals_function: None,
                totals_label: None,
            }],
            header_row: true,
            totals_row: false,
            style: None,
            autofilter: true,
        });
        sheet.tables.push(Table {
            name: "TableB".into(),
            display_name: None,
            range: "D1:E10".into(),
            columns: vec![TableColumn {
                name: "C1".into(),
                totals_function: None,
                totals_label: None,
            }],
            header_row: true,
            totals_row: false,
            style: None,
            autofilter: true,
        });
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // rId1=comments, rId2=VML, rId3=table[0], rId4=table[1].
        assert!(
            text.contains("<tableParts count=\"2\"><tablePart r:id=\"rId3\"/><tablePart r:id=\"rId4\"/></tableParts>"),
            "rId3/rId4 sequential with comments: {text}"
        );
    }

    // =========================================================================
    // Wave 3C — Conditional Formatting tests
    // =========================================================================

    use crate::model::conditional::{
        CellIsOperator, ColorScaleStop, ConditionalFormat, ConditionalKind, ConditionalRule,
        ConditionalThreshold,
    };
    use crate::model::validation::{DataValidation, ErrorStyle, ValidationOperator, ValidationType};

    fn make_cf(
        sqref: &str,
        rules: Vec<ConditionalRule>,
    ) -> ConditionalFormat {
        ConditionalFormat {
            sqref: sqref.to_string(),
            rules,
        }
    }

    fn make_rule(kind: ConditionalKind, dxf_id: Option<u32>, stop_if_true: bool) -> ConditionalRule {
        ConditionalRule { kind, dxf_id, stop_if_true }
    }

    // --- 34. CF absent when no conditional formats ---

    #[test]
    fn cf_absent_when_no_conditional_formats() {
        let sheet = Worksheet::new("S");
        let (bytes, _) = emit_sheet(&sheet, 0);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            !text.contains("<conditionalFormatting"),
            "no CF element on empty: {text}"
        );
    }

    // --- 35. CF cellIs greaterThan basic ---

    #[test]
    fn cf_cell_is_greater_than_basic() {
        let mut sheet = Worksheet::new("S");
        let rule = make_rule(
            ConditionalKind::CellIs {
                operator: CellIsOperator::GreaterThan,
                formula_a: "100".into(),
                formula_b: None,
            },
            Some(0),
            false,
        );
        sheet.conditional_formats.push(make_cf("A1:A10", vec![rule]));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<conditionalFormatting sqref=\"A1:A10\">"),
            "wrapper: {text}"
        );
        assert!(
            text.contains("<cfRule type=\"cellIs\" priority=\"1\" operator=\"greaterThan\" dxfId=\"0\">"),
            "cfRule attrs: {text}"
        );
        assert!(
            text.contains("<formula>100</formula>"),
            "formula child: {text}"
        );
        assert!(
            !text.contains("stopIfTrue"),
            "no stopIfTrue when false: {text}"
        );
    }

    // --- 36. CF cellIs Between emits two formulas ---

    #[test]
    fn cf_cell_is_between_emits_two_formulas() {
        let mut sheet = Worksheet::new("S");
        let rule = make_rule(
            ConditionalKind::CellIs {
                operator: CellIsOperator::Between,
                formula_a: "10".into(),
                formula_b: Some("20".into()),
            },
            Some(1),
            false,
        );
        sheet.conditional_formats.push(make_cf("B1:B5", vec![rule]));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<formula>10</formula><formula>20</formula>"),
            "two formulas for between: {text}"
        );
        assert!(
            text.contains("operator=\"between\""),
            "operator between: {text}"
        );
    }

    // --- 37. CF expression has no operator ---

    #[test]
    fn cf_expression_has_no_operator() {
        let mut sheet = Worksheet::new("S");
        let rule = make_rule(
            ConditionalKind::Expression { formula: "A1>B1".into() },
            Some(2),
            false,
        );
        sheet.conditional_formats.push(make_cf("C1", vec![rule]));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("type=\"expression\""),
            "type=expression: {text}"
        );
        // Must NOT have operator= attribute
        // Extract the cfRule tag to be precise
        let rule_start = text.find("<cfRule").expect("cfRule start");
        let rule_end = text[rule_start..].find('>').expect("cfRule end") + rule_start;
        let rule_tag = &text[rule_start..=rule_end];
        assert!(!rule_tag.contains("operator="), "no operator on expression: {rule_tag}");
        // Exactly one <formula> child
        assert_eq!(
            text.matches("<formula>").count(),
            1,
            "exactly one formula: {text}"
        );
        assert!(
            text.contains("<formula>A1&gt;B1</formula>"),
            "escaped formula: {text}"
        );
    }

    // --- 38. CF stopIfTrue emits or omits correctly ---

    #[test]
    fn cf_stop_if_true_emits_attribute() {
        let mut sheet = Worksheet::new("S");
        // Rule with stop_if_true=true
        let rule_stop = make_rule(
            ConditionalKind::Expression { formula: "A1>0".into() },
            None,
            true,
        );
        // Rule with stop_if_true=false
        let rule_no_stop = make_rule(
            ConditionalKind::Expression { formula: "A1<0".into() },
            None,
            false,
        );
        sheet.conditional_formats.push(make_cf("A1", vec![rule_stop, rule_no_stop]));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("stopIfTrue=\"1\""),
            "stopIfTrue=1 present: {text}"
        );
        // Should not contain stopIfTrue="0" anywhere
        assert!(
            !text.contains("stopIfTrue=\"0\""),
            "stopIfTrue=0 must not appear: {text}"
        );
        // Count occurrences of stopIfTrue — should be exactly 1
        assert_eq!(
            text.matches("stopIfTrue").count(),
            1,
            "exactly one stopIfTrue: {text}"
        );
    }

    // --- 39. CF dataBar has no dxfId ---

    #[test]
    fn cf_databar_has_no_dxfid() {
        let mut sheet = Worksheet::new("S");
        let rule = make_rule(
            ConditionalKind::DataBar {
                color_rgb: "FFFF0000".into(),
                min: ConditionalThreshold::Min,
                max: ConditionalThreshold::Max,
            },
            Some(99), // dxf_id should be ignored for DataBar
            false,
        );
        sheet.conditional_formats.push(make_cf("A1:A10", vec![rule]));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("type=\"dataBar\""),
            "type=dataBar: {text}"
        );
        // The cfRule element must NOT have dxfId
        let rule_start = text.find("<cfRule").expect("cfRule");
        let rule_end = text[rule_start..].find('>').expect(">") + rule_start;
        let rule_tag = &text[rule_start..=rule_end];
        assert!(!rule_tag.contains("dxfId"), "no dxfId on dataBar: {rule_tag}");
        // Inner structure
        assert!(
            text.contains("<dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FFFF0000\"/></dataBar>"),
            "dataBar structure: {text}"
        );
    }

    // --- 40. CF colorScale 2 stops ---

    #[test]
    fn cf_color_scale_2_stops() {
        let mut sheet = Worksheet::new("S");
        let rule = make_rule(
            ConditionalKind::ColorScale {
                stops: vec![
                    ColorScaleStop { threshold: ConditionalThreshold::Min, color_rgb: "FF0000FF".into() },
                    ColorScaleStop { threshold: ConditionalThreshold::Max, color_rgb: "FFFF0000".into() },
                ],
            },
            None,
            false,
        );
        sheet.conditional_formats.push(make_cf("A1:A10", vec![rule]));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("type=\"colorScale\""),
            "type=colorScale: {text}"
        );
        // cfvos before colors
        assert!(
            text.contains("<colorScale><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF0000FF\"/><color rgb=\"FFFF0000\"/></colorScale>"),
            "colorScale structure: {text}"
        );
        // No dxfId
        let rule_start = text.find("<cfRule").expect("cfRule");
        let rule_end = text[rule_start..].find('>').expect(">") + rule_start;
        let rule_tag = &text[rule_start..=rule_end];
        assert!(!rule_tag.contains("dxfId"), "no dxfId on colorScale: {rule_tag}");
    }

    // --- 41. CF colorScale 3 stops ---

    #[test]
    fn cf_color_scale_3_stops() {
        let mut sheet = Worksheet::new("S");
        let rule = make_rule(
            ConditionalKind::ColorScale {
                stops: vec![
                    ColorScaleStop { threshold: ConditionalThreshold::Min, color_rgb: "FF0000FF".into() },
                    ColorScaleStop { threshold: ConditionalThreshold::Percent(50.0), color_rgb: "FF00FF00".into() },
                    ColorScaleStop { threshold: ConditionalThreshold::Max, color_rgb: "FFFF0000".into() },
                ],
            },
            None,
            false,
        );
        sheet.conditional_formats.push(make_cf("A1:A10", vec![rule]));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // Three cfvo elements
        assert_eq!(text.matches("<cfvo").count(), 3, "three cfvo: {text}");
        // Three color elements
        assert_eq!(text.matches("<color rgb=").count(), 3, "three colors: {text}");
        // Percent threshold
        assert!(
            text.contains("<cfvo type=\"percent\" val=\"50\"/>"),
            "percent cfvo: {text}"
        );
        // All cfvos appear before all colors in the colorScale block
        let cs_start = text.find("<colorScale>").expect("colorScale");
        let cs_end = text.find("</colorScale>").expect("/colorScale");
        let cs_body = &text[cs_start..cs_end];
        let last_cfvo = cs_body.rfind("<cfvo").expect("last cfvo");
        let first_color = cs_body.find("<color rgb=").expect("first color");
        assert!(last_cfvo < first_color, "all cfvo before colors: {cs_body}");
    }

    // --- 42. CF stub variants are skipped ---

    #[test]
    fn cf_stub_variants_skipped() {
        let mut sheet = Worksheet::new("S");
        let rule = make_rule(
            ConditionalKind::Duplicate,
            None,
            false,
        );
        sheet.conditional_formats.push(make_cf("A1:A10", vec![rule]));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // Stub variants produce no cfRule element
        assert!(
            !text.contains("<cfRule type=\"duplicate\""),
            "no duplicate cfRule: {text}"
        );
        assert!(
            !text.contains("<cfRule type=\"unique\""),
            "no unique cfRule: {text}"
        );
        // The wrapper may or may not be emitted with empty children;
        // either way the document must parse and contain no invalid rules.
        // Verify no invalid cfRule appeared.
        assert!(
            !text.contains("<cfRule type=\"containsText\""),
            "no containsText cfRule: {text}"
        );
    }

    // --- 42b. CF wrapper omitted when every rule is a stub variant ---

    #[test]
    fn cf_all_stub_variants_no_wrapper() {
        // When every rule in a ConditionalFormat hits a stub arm, we must
        // skip the `<conditionalFormatting>` wrapper entirely. Excel treats
        // an empty `<conditionalFormatting sqref="..."></conditionalFormatting>`
        // as invalid and repairs the file on open.
        let mut sheet = Worksheet::new("S");
        let rules = vec![
            make_rule(ConditionalKind::Duplicate, None, false),
            make_rule(ConditionalKind::Unique, None, false),
        ];
        sheet.conditional_formats.push(make_cf("A1:A10", rules));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            !text.contains("<conditionalFormatting"),
            "no <conditionalFormatting> wrapper when all rules stubbed: {text}"
        );
        assert!(
            !text.contains("</conditionalFormatting>"),
            "no </conditionalFormatting> closing tag either: {text}"
        );
    }

    // --- 43. CF wellformed kitchen sink ---

    #[test]
    fn cf_wellformed_kitchen_sink() {
        let mut sheet = Worksheet::new("S");
        let cf = ConditionalFormat {
            sqref: "A1:D10".into(),
            rules: vec![
                make_rule(
                    ConditionalKind::CellIs {
                        operator: CellIsOperator::GreaterThan,
                        formula_a: "50".into(),
                        formula_b: None,
                    },
                    Some(0),
                    false,
                ),
                make_rule(
                    ConditionalKind::Expression { formula: "A1>B1".into() },
                    Some(1),
                    false,
                ),
                make_rule(
                    ConditionalKind::DataBar {
                        color_rgb: "FF0070C0".into(),
                        min: ConditionalThreshold::Min,
                        max: ConditionalThreshold::Max,
                    },
                    None,
                    false,
                ),
                make_rule(
                    ConditionalKind::ColorScale {
                        stops: vec![
                            ColorScaleStop { threshold: ConditionalThreshold::Min, color_rgb: "FFF8696B".into() },
                            ColorScaleStop { threshold: ConditionalThreshold::Max, color_rgb: "FF63BE7B".into() },
                        ],
                    },
                    None,
                    false,
                ),
            ],
        };
        sheet.conditional_formats.push(cf);
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
    }

    // =========================================================================
    // Wave 3C — Data Validation tests
    // =========================================================================

    fn make_dv(
        sqref: &str,
        validation_type: ValidationType,
        operator: ValidationOperator,
        formula_a: Option<&str>,
        formula_b: Option<&str>,
    ) -> DataValidation {
        DataValidation {
            sqref: sqref.to_string(),
            validation_type,
            operator,
            formula_a: formula_a.map(|s| s.to_string()),
            formula_b: formula_b.map(|s| s.to_string()),
            allow_blank: false,
            show_dropdown: false,
            show_error_message: false,
            error_style: ErrorStyle::Stop,
            error_title: None,
            error_message: None,
            show_input_message: false,
            input_title: None,
            input_message: None,
        }
    }

    // --- 44. DV absent when no validations ---

    #[test]
    fn dv_absent_when_no_validations() {
        let sheet = Worksheet::new("S");
        let (bytes, _) = emit_sheet(&sheet, 0);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            !text.contains("<dataValidations"),
            "no DV element on empty: {text}"
        );
    }

    // --- 45. DV list with literal string ---

    #[test]
    fn dv_list_with_literal_string() {
        let mut sheet = Worksheet::new("S");
        let dv = make_dv(
            "A1:A10",
            ValidationType::List,
            ValidationOperator::Between, // ignored for list
            Some("\"Red,Green,Blue\""),
            None,
        );
        sheet.validations.push(dv);
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<dataValidations count=\"1\">"),
            "count=1: {text}"
        );
        assert!(
            text.contains("<dataValidation type=\"list\""),
            "type=list: {text}"
        );
        assert!(
            text.contains("sqref=\"A1:A10\""),
            "sqref: {text}"
        );
        assert!(
            text.contains("<formula1>\"Red,Green,Blue\"</formula1>"),
            "formula1: {text}"
        );
        // List type must NOT have operator attribute
        let dv_start = text.find("<dataValidation").expect("dataValidation");
        let dv_end = text[dv_start..].find('>').expect(">") + dv_start;
        let dv_tag = &text[dv_start..=dv_end];
        assert!(!dv_tag.contains("operator="), "no operator for list: {dv_tag}");
    }

    // --- 46. DV list with range reference ---

    #[test]
    fn dv_list_with_range_reference() {
        let mut sheet = Worksheet::new("S");
        let dv = make_dv(
            "B1:B5",
            ValidationType::List,
            ValidationOperator::Between,
            Some("Sheet2!$A$1:$A$5"),
            None,
        );
        sheet.validations.push(dv);
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<formula1>Sheet2!$A$1:$A$5</formula1>"),
            "range ref formula1: {text}"
        );
    }

    // --- 47. DV whole between has two formulas ---

    #[test]
    fn dv_whole_between_has_two_formulas() {
        let mut sheet = Worksheet::new("S");
        let dv = make_dv(
            "C1",
            ValidationType::Whole,
            ValidationOperator::Between,
            Some("1"),
            Some("100"),
        );
        sheet.validations.push(dv);
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<formula1>1</formula1>"),
            "formula1: {text}"
        );
        assert!(
            text.contains("<formula2>100</formula2>"),
            "formula2: {text}"
        );
        assert!(
            text.contains("operator=\"between\""),
            "operator=between: {text}"
        );
    }

    // --- 48. DV whole greaterThan has one formula ---

    #[test]
    fn dv_whole_greater_than_has_one_formula() {
        let mut sheet = Worksheet::new("S");
        let dv = make_dv(
            "D1",
            ValidationType::Whole,
            ValidationOperator::GreaterThan,
            Some("0"),
            None,
        );
        sheet.validations.push(dv);
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<formula1>0</formula1>"),
            "formula1: {text}"
        );
        assert!(
            !text.contains("<formula2>"),
            "no formula2 for greaterThan: {text}"
        );
        assert!(
            text.contains("operator=\"greaterThan\""),
            "operator=greaterThan: {text}"
        );
    }

    // --- 49. DV custom with formula ---

    #[test]
    fn dv_custom_with_formula() {
        let mut sheet = Worksheet::new("S");
        let dv = make_dv(
            "E1",
            ValidationType::Custom,
            ValidationOperator::Between, // ignored for custom
            Some("A1>0"),
            None,
        );
        sheet.validations.push(dv);
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("type=\"custom\""),
            "type=custom: {text}"
        );
        // No operator attr for custom
        let dv_start = text.find("<dataValidation").expect("dataValidation");
        let dv_end = text[dv_start..].find('>').expect(">") + dv_start;
        let dv_tag = &text[dv_start..=dv_end];
        assert!(!dv_tag.contains("operator="), "no operator for custom: {dv_tag}");
        // > in formula is escaped as &gt;
        assert!(
            text.contains("<formula1>A1&gt;0</formula1>"),
            "escaped formula: {text}"
        );
    }

    // --- 50. DV error style warning ---

    #[test]
    fn dv_error_style_warning() {
        let mut sheet = Worksheet::new("S");
        let mut dv = make_dv("F1", ValidationType::Whole, ValidationOperator::Between, Some("0"), Some("100"));
        dv.error_style = ErrorStyle::Warning;
        dv.error_title = Some("Oops".into());
        dv.error_message = Some("Invalid".into());
        sheet.validations.push(dv);
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("errorStyle=\"warning\""),
            "errorStyle=warning: {text}"
        );
        assert!(
            text.contains("errorTitle=\"Oops\""),
            "errorTitle: {text}"
        );
        assert!(
            text.contains("error=\"Invalid\""),
            "error (not errorMessage): {text}"
        );
    }

    // --- 51. DV show flags ---

    #[test]
    fn dv_show_flags() {
        let mut sheet = Worksheet::new("S");
        let mut dv = make_dv("G1", ValidationType::Any, ValidationOperator::Between, None, None);
        dv.allow_blank = true;
        dv.show_input_message = true;
        dv.show_error_message = true;
        sheet.validations.push(dv);

        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("allowBlank=\"1\""), "allowBlank: {text}");
        assert!(text.contains("showInputMessage=\"1\""), "showInputMessage: {text}");
        assert!(text.contains("showErrorMessage=\"1\""), "showErrorMessage: {text}");

        // Now with all false
        let mut sheet2 = Worksheet::new("S");
        let dv2 = make_dv("G1", ValidationType::Any, ValidationOperator::Between, None, None);
        // all flags default to false
        sheet2.validations.push(dv2);
        let (bytes2, _) = emit_sheet(&sheet2, 0);
        let text2 = String::from_utf8(bytes2).unwrap();
        assert!(!text2.contains("allowBlank="), "no allowBlank when false: {text2}");
        assert!(!text2.contains("showInputMessage="), "no showInputMessage when false: {text2}");
        assert!(!text2.contains("showErrorMessage="), "no showErrorMessage when false: {text2}");
    }

    // --- 52. DV ordering: CF before DV before hyperlinks ---

    #[test]
    fn dv_ordering_between_cf_and_hyperlinks() {
        let mut sheet = Worksheet::new("S");
        // Add a conditional format
        let cf_rule = make_rule(
            ConditionalKind::Expression { formula: "A1>0".into() },
            None,
            false,
        );
        sheet.conditional_formats.push(make_cf("A1", vec![cf_rule]));

        // Add a data validation
        let dv = make_dv(
            "B1",
            ValidationType::Whole,
            ValidationOperator::GreaterThan,
            Some("0"),
            None,
        );
        sheet.validations.push(dv);

        // Add an external hyperlink
        sheet.hyperlinks.insert(
            "C1".to_string(),
            Hyperlink {
                target: "https://example.com".into(),
                is_internal: false,
                display: None,
                tooltip: None,
            },
        );

        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();

        let pos_cf = text.find("<conditionalFormatting").expect("CF position");
        let pos_dv = text.find("<dataValidations").expect("DV position");
        let pos_hl = text.find("<hyperlinks>").expect("hyperlinks position");

        assert!(pos_cf < pos_dv, "CF before DV: cf={pos_cf} dv={pos_dv}");
        assert!(pos_dv < pos_hl, "DV before hyperlinks: dv={pos_dv} hl={pos_hl}");
    }

}
