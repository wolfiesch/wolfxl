//! `xl/worksheets/sheet{N}.xml` emitter — rows, cells, merges, freeze,
//! columns, print area, and extension hooks for CF/DV. Wave 2B.
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

    // 1. <dimension>
    emit_dimension(&mut out, sheet);

    // 2. <sheetViews>
    emit_sheet_views(&mut out, sheet, sheet_idx);

    // 3. <sheetFormatPr>
    out.push_str("<sheetFormatPr defaultRowHeight=\"15\"/>");

    // 4. <cols> (only if non-empty)
    if !sheet.columns.is_empty() {
        emit_cols(&mut out, sheet);
    }

    // 5. <sheetData>
    emit_sheet_data(&mut out, sheet, sst);

    // 6. <mergeCells> (only if non-empty)
    if !sheet.merges.is_empty() {
        emit_merges(&mut out, sheet);
    }

    // EXT-W3C: conditional_formats — inserted between mergeCells and hyperlinks
    emit_conditional_formats(&mut out, sheet);

    // EXT-W3C: data_validations — inserted between CF and hyperlinks
    emit_data_validations(&mut out, sheet);

    // 7. <hyperlinks> (only if any exist)
    if !sheet.hyperlinks.is_empty() {
        emit_hyperlinks(&mut out, sheet);
    }

    // 8. <pageMargins>
    out.push_str("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>");

    // EXT-W3A: legacyDrawing — emitted iff !sheet.comments.is_empty(); rId via convention
    emit_legacy_drawing(&mut out, sheet);

    // EXT-W3B: tableParts — one <tablePart r:id=...> per table
    emit_table_parts(&mut out, sheet);

    out.push_str("</worksheet>");

    out.into_bytes()
}

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
fn emit_freeze_pane(out: &mut String, freeze: &crate::model::worksheet::FreezePane) {
    let has_row = freeze.freeze_row >= 1;
    let has_col = freeze.freeze_col >= 1;

    // Determine top-left cell for the active pane
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
        out.push_str(&format!(" xSplit=\"{}\"", freeze.freeze_col));
    }
    if has_row {
        out.push_str(&format!(" ySplit=\"{}\"", freeze.freeze_row));
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
                    out.push_str(&format!("<c r=\"{}\"", cell_ref));
                    if let Some(s) = cell.style_id {
                        out.push_str(&format!(" s=\"{}\"", s));
                    }
                    out.push_str(&format!("><f>{}</f></c>", escaped_expr));
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
        let is_internal = hyperlink.target.starts_with('#');

        if is_internal {
            // Strip leading '#' for location attribute
            let location = &hyperlink.target[1..];
            out.push_str(&format!(
                "<hyperlink ref=\"{}\" location=\"{}\"",
                xml_escape::attr(cell_ref),
                xml_escape::attr(location)
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
fn emit_conditional_formats(_out: &mut String, sheet: &Worksheet) {
    if !sheet.conditional_formats.is_empty() {
        // W3C fills in
    }
}

/// Emit `<dataValidations count="N">…</dataValidations>` between the CF
/// block and `<hyperlinks>`. Filled by W3C.
fn emit_data_validations(_out: &mut String, sheet: &Worksheet) {
    if !sheet.validations.is_empty() {
        // W3C fills in
    }
}

/// Emit `<legacyDrawing r:id="rId2"/>` when comments exist. The rId is a
/// hardcoded convention (see `rels::emit_sheet`): with comments present,
/// rId1 → commentsN.xml and rId2 → vmlDrawingN.vml. Filled by W3A.
fn emit_legacy_drawing(_out: &mut String, sheet: &Worksheet) {
    if !sheet.comments.is_empty() {
        // W3A fills in
    }
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

    // --- 10. Formula without result has no <v> ---

    #[test]
    fn formula_without_result_has_no_v() {
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
        assert!(text.contains("<f>SUM(A1:A10)</f>"), "formula expr: {text}");
        // The cell should not have a <v> sibling
        // Check that there's no <v> inside the C5 cell
        let cell_start = text.find("<c r=\"C5\"").expect("C5 cell");
        let cell_end = text[cell_start..].find("</c>").expect("</c>") + cell_start;
        let cell_xml = &text[cell_start..=cell_end + 3];
        assert!(!cell_xml.contains("<v>"), "no <v> for formula without result: {cell_xml}");
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
        assert!(text.contains("ySplit=\"3\""), "ySplit: {text}");
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
        assert!(text.contains("xSplit=\"2\""), "xSplit: {text}");
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
        assert!(text.contains("xSplit=\"3\""), "xSplit: {text}");
        assert!(text.contains("ySplit=\"2\""), "ySplit: {text}");
        assert!(text.contains("state=\"frozen\""), "state=frozen: {text}");
        assert!(text.contains("activePane=\"bottomRight\""), "activePane: {text}");
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
                target: "#Sheet2!A1".into(),
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
                display: Some("Example".into()),
                tooltip: None,
            },
        );
        sheet.hyperlinks.insert(
            "B1".to_string(),
            Hyperlink {
                target: "#Sheet2!A1".into(),
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

    // --- 34. table_parts_absent_when_no_tables ---

    #[test]
    fn table_parts_absent_when_no_tables() {
        let sheet = Worksheet::new("S");
        let (bytes, _) = emit_sheet(&sheet, 0);
        let text = String::from_utf8(bytes).unwrap();
        assert!(!text.contains("<tableParts"), "no tableParts when none: {text}");
    }

    // --- 35. table_parts_no_comments_starts_at_rid1 ---

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

    // --- 36. table_parts_with_comments_starts_at_rid3 ---

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
}
