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
use crate::model::format::StylesBuilder;
use crate::model::worksheet::Worksheet;
use crate::xml_escape;

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
    super::dimension::emit(&mut out, sheet);

    // Slot 3: <sheetViews>
    //
    // RFC-055 §3 (Sprint Ο Pod 1A.5): if the user set a typed
    // `views` spec on the Worksheet, prefer it; the legacy path is
    // still used otherwise so freeze/split panes set via
    // `set_freeze`/`set_split` continue to work.
    if let Some(view_spec) = sheet.views.as_ref() {
        let bytes = crate::parse::sheet_setup::emit_sheet_views(view_spec);
        if !bytes.is_empty() {
            out.push_str(std::str::from_utf8(&bytes).unwrap_or(""));
        } else {
            super::sheet_views::emit(&mut out, sheet, sheet_idx);
        }
    } else {
        super::sheet_views::emit(&mut out, sheet, sheet_idx);
    }

    // Slot 4: <sheetFormatPr>
    //
    // RFC-062: if the user set a typed `sheet_format` spec on the
    // Worksheet, prefer it; the legacy hardcoded default is still
    // emitted otherwise so unmodified sheets keep byte-stable.
    if let Some(spec) = sheet.sheet_format.as_ref() {
        let bytes = crate::parse::page_breaks::emit_sheet_format_pr(spec);
        if !bytes.is_empty() {
            out.push_str(std::str::from_utf8(&bytes).unwrap_or(""));
        } else {
            out.push_str("<sheetFormatPr defaultRowHeight=\"15\"/>");
        }
    } else {
        out.push_str("<sheetFormatPr defaultRowHeight=\"15\"/>");
    }

    // Slot 5: <cols> (only if non-empty)
    if !sheet.columns.is_empty() {
        super::columns::emit(&mut out, sheet);
    }

    // Slot 6: <sheetData>
    super::sheet_data::emit(&mut out, sheet, sst);

    // Slot 8: <sheetProtection> — Sprint Ο Pod 1A.5 (RFC-055).
    if let Some(spec) = sheet.protection.as_ref() {
        let bytes = crate::parse::sheet_setup::emit_sheet_protection(spec);
        if !bytes.is_empty() {
            out.push_str(std::str::from_utf8(&bytes).unwrap_or(""));
        }
    }

    // Slot 11: <autoFilter> — Sprint Ο Pod 1B (RFC-056). The bytes
    // are pre-emitted by the workbook-level coordinator from the
    // Python `ws.auto_filter.to_rust_dict()` payload via
    // `wolfxl_autofilter::emit::emit`.
    if let Some(bytes) = &sheet.auto_filter_xml {
        out.push_str(std::str::from_utf8(bytes).unwrap_or(""));
    }

    // Slot 15: <mergeCells> (only if non-empty)
    if !sheet.merges.is_empty() {
        super::merges::emit(&mut out, sheet);
    }

    // Slot 17: <conditionalFormatting> — EXT-W3C; 0..N elements per spec
    emit_conditional_formats(&mut out, sheet);

    // Slot 18: <dataValidations> — EXT-W3C
    super::data_validations::emit(&mut out, sheet);

    // Slot 19: <hyperlinks> (only if any exist)
    if !sheet.hyperlinks.is_empty() {
        super::hyperlinks::emit(&mut out, sheet);
    }

    // Slot 21: <pageMargins> — RFC-055 typed override or default.
    if let Some(spec) = sheet.page_margins.as_ref() {
        let bytes = crate::parse::sheet_setup::emit_page_margins(spec);
        out.push_str(std::str::from_utf8(&bytes).unwrap_or(""));
    } else {
        out.push_str("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>");
    }

    // Slot 22: <pageSetup> — RFC-055 (only emitted when set).
    if let Some(spec) = sheet.page_setup.as_ref() {
        let bytes = crate::parse::sheet_setup::emit_page_setup(spec);
        if !bytes.is_empty() {
            out.push_str(std::str::from_utf8(&bytes).unwrap_or(""));
        }
    }

    // Slot 23: <headerFooter> — RFC-055 (only emitted when set).
    if let Some(spec) = sheet.header_footer.as_ref() {
        let bytes = crate::parse::sheet_setup::emit_header_footer(spec);
        if !bytes.is_empty() {
            out.push_str(std::str::from_utf8(&bytes).unwrap_or(""));
        }
    }

    // Slot 24: <rowBreaks> — RFC-062 (only emitted when set+non-empty).
    if let Some(spec) = sheet.row_breaks.as_ref() {
        let bytes = crate::parse::page_breaks::emit_row_breaks(spec);
        if !bytes.is_empty() {
            out.push_str(std::str::from_utf8(&bytes).unwrap_or(""));
        }
    }

    // Slot 25: <colBreaks> — RFC-062 (only emitted when set+non-empty).
    if let Some(spec) = sheet.col_breaks.as_ref() {
        let bytes = crate::parse::page_breaks::emit_col_breaks(spec);
        if !bytes.is_empty() {
            out.push_str(std::str::from_utf8(&bytes).unwrap_or(""));
        }
    }

    // Slot 30: <drawing r:id="..."/> — Sprint Λ Pod-β (RFC-045);
    // emitted iff !sheet.images.is_empty(). The rId is appended at
    // the END of the sheet's rels graph (after comments, vml, tables,
    // and external hyperlinks) so the existing rId conventions for
    // those entries are preserved.
    super::drawing_refs::emit_drawing(&mut out, sheet);

    // Slot 31: <legacyDrawing> — EXT-W3A; emitted iff !sheet.comments.is_empty(); rId via convention
    super::drawing_refs::emit_legacy(&mut out, sheet);

    // Slot 37: <tableParts> — EXT-W3B; one <tablePart r:id=...> per table
    super::table_parts::emit(&mut out, sheet);

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
                    ConditionalKind::CellIs {
                        operator,
                        formula_a,
                        formula_b,
                    } => {
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
                        rules_buf.push_str(&format!(
                            "<formula>{}</formula>",
                            xml_escape::text(formula_a)
                        ));
                        let needs_second = matches!(
                            operator,
                            CellIsOperator::Between | CellIsOperator::NotBetween
                        );
                        if needs_second {
                            if let Some(fb) = formula_b {
                                rules_buf.push_str(&format!(
                                    "<formula>{}</formula>",
                                    xml_escape::text(fb)
                                ));
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
                        rules_buf
                            .push_str(&format!("<formula>{}</formula>", xml_escape::text(formula)));
                        rules_buf.push_str("</cfRule>");
                    }

                    ConditionalKind::DataBar {
                        color_rgb,
                        min,
                        max,
                    } => {
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
            out.push_str(&format!(
                "<cfvo type=\"percent\" val=\"{}\"/>",
                format_f64(*x)
            ));
        }
        ConditionalThreshold::Percentile(x) => {
            out.push_str(&format!(
                "<cfvo type=\"percentile\" val=\"{}\"/>",
                format_f64(*x)
            ));
        }
        ConditionalThreshold::Formula(s) => {
            out.push_str(&format!(
                "<cfvo type=\"formula\" val=\"{}\"/>",
                xml_escape::attr(s)
            ));
        }
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
        assert!(
            text.contains("<dimension ref=\"A1\"/>"),
            "dimension: {text}"
        );
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
        assert!(
            text.contains("<c r=\"A1\" s=\"3\"/>"),
            "blank+style: {text}"
        );
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
        assert!(
            !text.contains("<v>42.0</v>"),
            "should not have decimal: {text}"
        );
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
        assert!(
            text.contains("<c r=\"A1\" t=\"s\">"),
            "t=s attribute: {text}"
        );
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
        assert!(
            text.contains("<c r=\"A1\" t=\"s\"><v>0</v></c>"),
            "A1 beta=0: {text}"
        );
        // A2 (row2,col1) = alpha = index 1
        assert!(
            text.contains("<c r=\"A2\" t=\"s\"><v>1</v></c>"),
            "A2 alpha=1: {text}"
        );
        // A3 (row3,col1) = beta = index 0 (deduped)
        assert!(
            text.contains("<c r=\"A3\" t=\"s\"><v>0</v></c>"),
            "A3 beta=0: {text}"
        );
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
        assert!(
            text.contains("<f>1+6</f><v>7</v>"),
            "formula+numeric result: {text}"
        );
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
        sheet.set_cell(
            1,
            1,
            WriteCell::new(WriteCellValue::Number(1.0)).with_style(5),
        );
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
        assert!(
            !text.contains("<mergeCells"),
            "no mergeCells when none: {text}"
        );
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
        assert!(
            text.contains("activePane=\"bottomLeft\""),
            "activePane: {text}"
        );
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
        assert!(
            text.contains("activePane=\"topRight\""),
            "activePane: {text}"
        );
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
        assert!(
            text.contains("activePane=\"bottomRight\""),
            "activePane: {text}"
        );
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
        assert!(
            text.contains("activePane=\"bottomRight\""),
            "activePane: {text}"
        );
        // Negative: must NOT emit the cell coordinate as the count.
        assert!(
            !text.contains("xSplit=\"2\""),
            "xSplit must not be 2: {text}"
        );
        assert!(
            !text.contains("ySplit=\"2\""),
            "ySplit must not be 2: {text}"
        );
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
        assert!(
            !text.contains("<pane"),
            "must not emit pane for A1 freeze: {text}"
        );
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
        assert!(
            !text.contains("state=\"frozen\""),
            "no frozen for split: {text}"
        );
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
        assert!(
            !text.contains("customWidth="),
            "no customWidth when none: {text}"
        );
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
        assert!(
            text.contains("style=\"4\" customFormat=\"1\""),
            "style+customFormat: {text}"
        );
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
        assert!(
            text.contains("r:id=\"rId1\""),
            "rId1 for ext hyperlink: {text}"
        );
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
        sheet.set_cell(2, 3, WriteCell::new(WriteCellValue::Blank).with_style(1));

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
            text.contains("<dimension ref=\"E10\"/>")
                || text.contains("<dimension ref=\"A1:E10\"/>"),
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
        assert!(
            !text.contains("<tableParts"),
            "no tableParts when none: {text}"
        );
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
    use crate::model::validation::{
        DataValidation, ErrorStyle, ValidationOperator, ValidationType,
    };

    fn make_cf(sqref: &str, rules: Vec<ConditionalRule>) -> ConditionalFormat {
        ConditionalFormat {
            sqref: sqref.to_string(),
            rules,
        }
    }

    fn make_rule(
        kind: ConditionalKind,
        dxf_id: Option<u32>,
        stop_if_true: bool,
    ) -> ConditionalRule {
        ConditionalRule {
            kind,
            dxf_id,
            stop_if_true,
        }
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
        sheet
            .conditional_formats
            .push(make_cf("A1:A10", vec![rule]));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<conditionalFormatting sqref=\"A1:A10\">"),
            "wrapper: {text}"
        );
        assert!(
            text.contains(
                "<cfRule type=\"cellIs\" priority=\"1\" operator=\"greaterThan\" dxfId=\"0\">"
            ),
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
            ConditionalKind::Expression {
                formula: "A1>B1".into(),
            },
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
        assert!(
            !rule_tag.contains("operator="),
            "no operator on expression: {rule_tag}"
        );
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
            ConditionalKind::Expression {
                formula: "A1>0".into(),
            },
            None,
            true,
        );
        // Rule with stop_if_true=false
        let rule_no_stop = make_rule(
            ConditionalKind::Expression {
                formula: "A1<0".into(),
            },
            None,
            false,
        );
        sheet
            .conditional_formats
            .push(make_cf("A1", vec![rule_stop, rule_no_stop]));
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
        sheet
            .conditional_formats
            .push(make_cf("A1:A10", vec![rule]));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("type=\"dataBar\""), "type=dataBar: {text}");
        // The cfRule element must NOT have dxfId
        let rule_start = text.find("<cfRule").expect("cfRule");
        let rule_end = text[rule_start..].find('>').expect(">") + rule_start;
        let rule_tag = &text[rule_start..=rule_end];
        assert!(
            !rule_tag.contains("dxfId"),
            "no dxfId on dataBar: {rule_tag}"
        );
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
                    ColorScaleStop {
                        threshold: ConditionalThreshold::Min,
                        color_rgb: "FF0000FF".into(),
                    },
                    ColorScaleStop {
                        threshold: ConditionalThreshold::Max,
                        color_rgb: "FFFF0000".into(),
                    },
                ],
            },
            None,
            false,
        );
        sheet
            .conditional_formats
            .push(make_cf("A1:A10", vec![rule]));
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
        assert!(
            !rule_tag.contains("dxfId"),
            "no dxfId on colorScale: {rule_tag}"
        );
    }

    // --- 41. CF colorScale 3 stops ---

    #[test]
    fn cf_color_scale_3_stops() {
        let mut sheet = Worksheet::new("S");
        let rule = make_rule(
            ConditionalKind::ColorScale {
                stops: vec![
                    ColorScaleStop {
                        threshold: ConditionalThreshold::Min,
                        color_rgb: "FF0000FF".into(),
                    },
                    ColorScaleStop {
                        threshold: ConditionalThreshold::Percent(50.0),
                        color_rgb: "FF00FF00".into(),
                    },
                    ColorScaleStop {
                        threshold: ConditionalThreshold::Max,
                        color_rgb: "FFFF0000".into(),
                    },
                ],
            },
            None,
            false,
        );
        sheet
            .conditional_formats
            .push(make_cf("A1:A10", vec![rule]));
        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // Three cfvo elements
        assert_eq!(text.matches("<cfvo").count(), 3, "three cfvo: {text}");
        // Three color elements
        assert_eq!(
            text.matches("<color rgb=").count(),
            3,
            "three colors: {text}"
        );
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
        let rule = make_rule(ConditionalKind::Duplicate, None, false);
        sheet
            .conditional_formats
            .push(make_cf("A1:A10", vec![rule]));
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
                    ConditionalKind::Expression {
                        formula: "A1>B1".into(),
                    },
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
                            ColorScaleStop {
                                threshold: ConditionalThreshold::Min,
                                color_rgb: "FFF8696B".into(),
                            },
                            ColorScaleStop {
                                threshold: ConditionalThreshold::Max,
                                color_rgb: "FF63BE7B".into(),
                            },
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
        assert!(text.contains("sqref=\"A1:A10\""), "sqref: {text}");
        assert!(
            text.contains("<formula1>\"Red,Green,Blue\"</formula1>"),
            "formula1: {text}"
        );
        // List type must NOT have operator attribute
        let dv_start = text.find("<dataValidation").expect("dataValidation");
        let dv_end = text[dv_start..].find('>').expect(">") + dv_start;
        let dv_tag = &text[dv_start..=dv_end];
        assert!(
            !dv_tag.contains("operator="),
            "no operator for list: {dv_tag}"
        );
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
        assert!(text.contains("<formula1>1</formula1>"), "formula1: {text}");
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
        assert!(text.contains("<formula1>0</formula1>"), "formula1: {text}");
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
        assert!(text.contains("type=\"custom\""), "type=custom: {text}");
        // No operator attr for custom
        let dv_start = text.find("<dataValidation").expect("dataValidation");
        let dv_end = text[dv_start..].find('>').expect(">") + dv_start;
        let dv_tag = &text[dv_start..=dv_end];
        assert!(
            !dv_tag.contains("operator="),
            "no operator for custom: {dv_tag}"
        );
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
        let mut dv = make_dv(
            "F1",
            ValidationType::Whole,
            ValidationOperator::Between,
            Some("0"),
            Some("100"),
        );
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
        assert!(text.contains("errorTitle=\"Oops\""), "errorTitle: {text}");
        assert!(
            text.contains("error=\"Invalid\""),
            "error (not errorMessage): {text}"
        );
    }

    // --- 51. DV show flags ---

    #[test]
    fn dv_show_flags() {
        let mut sheet = Worksheet::new("S");
        let mut dv = make_dv(
            "G1",
            ValidationType::Any,
            ValidationOperator::Between,
            None,
            None,
        );
        dv.allow_blank = true;
        dv.show_input_message = true;
        dv.show_error_message = true;
        sheet.validations.push(dv);

        let (bytes, _) = emit_sheet(&sheet, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("allowBlank=\"1\""), "allowBlank: {text}");
        assert!(
            text.contains("showInputMessage=\"1\""),
            "showInputMessage: {text}"
        );
        assert!(
            text.contains("showErrorMessage=\"1\""),
            "showErrorMessage: {text}"
        );

        // Now with all false
        let mut sheet2 = Worksheet::new("S");
        let dv2 = make_dv(
            "G1",
            ValidationType::Any,
            ValidationOperator::Between,
            None,
            None,
        );
        // all flags default to false
        sheet2.validations.push(dv2);
        let (bytes2, _) = emit_sheet(&sheet2, 0);
        let text2 = String::from_utf8(bytes2).unwrap();
        assert!(
            !text2.contains("allowBlank="),
            "no allowBlank when false: {text2}"
        );
        assert!(
            !text2.contains("showInputMessage="),
            "no showInputMessage when false: {text2}"
        );
        assert!(
            !text2.contains("showErrorMessage="),
            "no showErrorMessage when false: {text2}"
        );
    }

    // --- 52. DV ordering: CF before DV before hyperlinks ---

    #[test]
    fn dv_ordering_between_cf_and_hyperlinks() {
        let mut sheet = Worksheet::new("S");
        // Add a conditional format
        let cf_rule = make_rule(
            ConditionalKind::Expression {
                formula: "A1>0".into(),
            },
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
        assert!(
            pos_dv < pos_hl,
            "DV before hyperlinks: dv={pos_dv} hl={pos_hl}"
        );
    }
}
