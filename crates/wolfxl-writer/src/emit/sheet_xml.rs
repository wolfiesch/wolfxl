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
    super::sheet_setup::emit_sheet_views(&mut out, sheet, sheet_idx);

    // Slot 4: <sheetFormatPr>
    //
    // RFC-062: if the user set a typed `sheet_format` spec on the
    // Worksheet, prefer it; the legacy hardcoded default is still
    // emitted otherwise so unmodified sheets keep byte-stable.
    super::sheet_setup::emit_sheet_format(&mut out, sheet);

    // Slot 5: <cols> (only if non-empty)
    if !sheet.columns.is_empty() {
        super::columns::emit(&mut out, sheet);
    }

    // Slot 6: <sheetData>
    super::sheet_data::emit(&mut out, sheet, sst);

    // Slot 8: <sheetProtection> — Sprint Ο Pod 1A.5 (RFC-055).
    super::sheet_setup::emit_sheet_protection(&mut out, sheet);

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
    super::conditional_formats::emit(&mut out, sheet);

    // Slot 18: <dataValidations> — EXT-W3C
    super::data_validations::emit(&mut out, sheet);

    // Slot 19: <hyperlinks> (only if any exist)
    if !sheet.hyperlinks.is_empty() {
        super::hyperlinks::emit(&mut out, sheet);
    }

    // Slot 21: <pageMargins> — RFC-055 typed override or default.
    super::sheet_setup::emit_page_margins(&mut out, sheet);

    // Slot 22: <pageSetup> — RFC-055 (only emitted when set).
    super::sheet_setup::emit_page_setup(&mut out, sheet);

    // Slot 23: <headerFooter> — RFC-055 (only emitted when set).
    super::sheet_setup::emit_header_footer(&mut out, sheet);

    // Slot 24: <rowBreaks> — RFC-062 (only emitted when set+non-empty).
    // Slot 25: <colBreaks> — RFC-062 (only emitted when set+non-empty).
    super::page_breaks::emit(&mut out, sheet);

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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::cell::{FormulaResult, WriteCell, WriteCellValue};
    use crate::model::worksheet::{Column, FreezePane, Hyperlink, Merge, Worksheet};
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
