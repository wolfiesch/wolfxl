//! `xl/workbook.xml` emitter. Wave 2C.
//!
//! Emits the workbook XML part that lists sheets, defined names,
//! and a handful of fixed metadata elements that Excel requires.

use crate::model::defined_name::BuiltinName;
use crate::model::worksheet::SheetVisibility;
use crate::model::workbook::Workbook;
use crate::refs::quote_sheet_name_if_needed;
use crate::xml_escape;

const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Return the `_xlnm.` prefixed name for a built-in defined name.
fn builtin_name_str(b: BuiltinName) -> &'static str {
    match b {
        BuiltinName::PrintArea => "_xlnm.Print_Area",
        BuiltinName::PrintTitles => "_xlnm.Print_Titles",
        BuiltinName::FilterDatabase => "_xlnm._FilterDatabase",
    }
}

/// Emit `xl/workbook.xml` as UTF-8 bytes.
pub fn emit(wb: &Workbook) -> Vec<u8> {
    let mut out = String::with_capacity(1024);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    out.push_str(&format!(
        "<workbook xmlns=\"{NS_MAIN}\" xmlns:r=\"{NS_R}\">"
    ));

    // Fixed file-version and workbook-properties elements.
    out.push_str(
        "<fileVersion appName=\"xl\" lastEdited=\"7\" lowestEdited=\"7\" rupBuild=\"10000\"/>",
    );
    out.push_str("<workbookPr date1904=\"false\"/>");
    out.push_str(
        "<bookViews>\
         <workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"20490\" windowHeight=\"7770\"/>\
         </bookViews>",
    );

    // Sheets.
    out.push_str("<sheets>");
    for (idx, sheet) in wb.sheets.iter().enumerate() {
        let n = idx + 1;
        let name_escaped = xml_escape::attr(&sheet.name);
        let state_attr = match sheet.visibility {
            SheetVisibility::Visible => String::new(),
            SheetVisibility::Hidden => " state=\"hidden\"".to_string(),
            SheetVisibility::VeryHidden => " state=\"veryHidden\"".to_string(),
        };
        out.push_str(&format!(
            "<sheet name=\"{name_escaped}\" sheetId=\"{n}\" r:id=\"rId{n}\"{state_attr}/>"
        ));
    }
    out.push_str("</sheets>");

    // Collect all defined names:
    //   1. User-defined (from wb.defined_names).
    //   2. Auto-injected print areas from sheets that have print_area set.
    //
    // Only emit the <definedNames> block when there is at least one entry.
    let has_user_names = !wb.defined_names.is_empty();
    let has_print_areas = wb.sheets.iter().any(|s| s.print_area.is_some());

    if has_user_names || has_print_areas {
        out.push_str("<definedNames>");

        // User-defined names.
        for dn in &wb.defined_names {
            let name_str = match dn.builtin {
                Some(b) => builtin_name_str(b).to_string(),
                None => dn.name.clone(),
            };
            let name_escaped = xml_escape::attr(&name_str);

            let local_attr = match dn.scope_sheet_index {
                Some(idx) => format!(" localSheetId=\"{idx}\""),
                None => String::new(),
            };

            let hidden_attr = if dn.hidden {
                " hidden=\"1\"".to_string()
            } else {
                String::new()
            };

            let formula_escaped = xml_escape::text(&dn.formula);
            out.push_str(&format!(
                "<definedName name=\"{name_escaped}\"{local_attr}{hidden_attr}>{formula_escaped}</definedName>"
            ));
        }

        // Auto-injected print areas.
        for (idx, sheet) in wb.sheets.iter().enumerate() {
            if let Some(ref range) = sheet.print_area {
                let quoted_name = quote_sheet_name_if_needed(&sheet.name);
                let formula = format!("{quoted_name}!{range}");
                let formula_escaped = xml_escape::text(&formula);
                out.push_str(&format!(
                    "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"{idx}\">{formula_escaped}</definedName>"
                ));
            }
        }

        out.push_str("</definedNames>");
    }

    out.push_str("<calcPr calcId=\"171027\"/>");
    out.push_str("</workbook>");
    out.into_bytes()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::defined_name::{BuiltinName, DefinedName};
    use crate::model::workbook::Workbook;
    use crate::model::worksheet::{SheetVisibility, Worksheet};
    use quick_xml::events::Event;
    use quick_xml::Reader;

    fn parse_ok(bytes: &[u8]) {
        let text = std::str::from_utf8(bytes).expect("utf8");
        let mut reader = Reader::from_str(text);
        reader.config_mut().check_end_names = true;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Eof) => break,
                Err(e) => panic!("parse error: {e}"),
                _ => (),
            }
            buf.clear();
        }
    }

    fn text_of(bytes: &[u8]) -> String {
        String::from_utf8(bytes.to_vec()).expect("utf8")
    }

    // 1. Single sheet minimal.
    #[test]
    fn single_sheet_minimal() {
        let mut wb = Workbook::new();
        wb.add_sheet(Worksheet::new("Sheet1"));
        let bytes = emit(&wb);
        parse_ok(&bytes);
        let text = text_of(&bytes);
        assert!(
            text.contains("<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>"),
            "{text}"
        );
        assert!(!text.contains("<definedNames>"), "no defined names expected: {text}");
    }

    // 2. Multiple sheets numbered correctly.
    #[test]
    fn multiple_sheets_numbered_correctly() {
        let mut wb = Workbook::new();
        wb.add_sheet(Worksheet::new("a"));
        wb.add_sheet(Worksheet::new("b"));
        wb.add_sheet(Worksheet::new("c"));
        let bytes = emit(&wb);
        parse_ok(&bytes);
        let text = text_of(&bytes);
        assert!(text.contains("sheetId=\"1\" r:id=\"rId1\""), "{text}");
        assert!(text.contains("sheetId=\"2\" r:id=\"rId2\""), "{text}");
        assert!(text.contains("sheetId=\"3\" r:id=\"rId3\""), "{text}");
    }

    // 3. Hidden and VeryHidden sheets emit correct state attribute.
    #[test]
    fn hidden_and_veryhidden_emit_state() {
        let mut wb = Workbook::new();
        let mut visible = Worksheet::new("Visible");
        visible.visibility = SheetVisibility::Visible;
        let mut hidden = Worksheet::new("Hidden");
        hidden.visibility = SheetVisibility::Hidden;
        let mut very_hidden = Worksheet::new("VeryHidden");
        very_hidden.visibility = SheetVisibility::VeryHidden;
        wb.add_sheet(visible);
        wb.add_sheet(hidden);
        wb.add_sheet(very_hidden);
        let bytes = emit(&wb);
        parse_ok(&bytes);
        let text = text_of(&bytes);
        // Visible: no state attribute.
        assert!(
            text.contains("<sheet name=\"Visible\" sheetId=\"1\" r:id=\"rId1\"/>"),
            "visible sheet should have no state attr: {text}"
        );
        assert!(text.contains("state=\"hidden\""), "{text}");
        assert!(text.contains("state=\"veryHidden\""), "{text}");
    }

    // 4. Workbook-scope defined name.
    #[test]
    fn defined_name_workbook_scope() {
        let mut wb = Workbook::new();
        wb.add_sheet(Worksheet::new("Sheet1"));
        wb.defined_names.push(DefinedName {
            name: "myRange".to_string(),
            formula: "Sheet1!A1:B2".to_string(),
            scope_sheet_index: None,
            builtin: None,
            hidden: false,
        });
        let bytes = emit(&wb);
        parse_ok(&bytes);
        let text = text_of(&bytes);
        assert!(
            text.contains("<definedName name=\"myRange\">Sheet1!A1:B2</definedName>"),
            "{text}"
        );
        // No localSheetId when scope is workbook.
        assert!(
            !text.contains("localSheetId"),
            "unexpected localSheetId for workbook-scope name: {text}"
        );
    }

    // 5. Sheet-scope defined name has localSheetId.
    #[test]
    fn defined_name_sheet_scope() {
        let mut wb = Workbook::new();
        wb.add_sheet(Worksheet::new("Sheet1"));
        wb.add_sheet(Worksheet::new("Sheet2"));
        wb.defined_names.push(DefinedName {
            name: "localRange".to_string(),
            formula: "Sheet2!C3:D4".to_string(),
            scope_sheet_index: Some(1),
            builtin: None,
            hidden: false,
        });
        let bytes = emit(&wb);
        parse_ok(&bytes);
        let text = text_of(&bytes);
        assert!(text.contains("localSheetId=\"1\""), "{text}");
    }

    // 6. Auto-inject print area from sheet.
    #[test]
    fn auto_inject_print_area_from_sheet() {
        let mut wb = Workbook::new();
        wb.add_sheet(Worksheet::new("Sheet1"));
        let mut sheet2 = Worksheet::new("Sheet2");
        sheet2.print_area = Some("A1:D20".to_string());
        wb.add_sheet(sheet2);
        let bytes = emit(&wb);
        parse_ok(&bytes);
        let text = text_of(&bytes);
        assert!(
            text.contains("<definedName name=\"_xlnm.Print_Area\" localSheetId=\"1\">Sheet2!A1:D20</definedName>"),
            "{text}"
        );
    }

    // 7. Print area quotes sheet name containing spaces.
    #[test]
    fn print_area_quotes_sheet_name_with_spaces() {
        let mut wb = Workbook::new();
        let mut sheet = Worksheet::new("Data Set");
        sheet.print_area = Some("A1:D20".to_string());
        wb.add_sheet(sheet);
        let bytes = emit(&wb);
        parse_ok(&bytes);
        let text = text_of(&bytes);
        assert!(
            text.contains("'Data Set'!A1:D20"),
            "expected quoted sheet name in formula: {text}"
        );
    }

    // 8. No <definedNames> block when both collections are empty.
    #[test]
    fn defined_names_block_omitted_when_both_empty() {
        let mut wb = Workbook::new();
        wb.add_sheet(Worksheet::new("Sheet1"));
        let bytes = emit(&wb);
        parse_ok(&bytes);
        let text = text_of(&bytes);
        assert!(
            !text.contains("<definedNames>"),
            "unexpected <definedNames> when both collections empty: {text}"
        );
    }

    // 9. Sheet name with XML special chars is attribute-escaped.
    #[test]
    fn sheet_name_with_xml_special_char_escaped() {
        let mut wb = Workbook::new();
        wb.add_sheet(Worksheet::new("A & B"));
        let bytes = emit(&wb);
        parse_ok(&bytes);
        let text = text_of(&bytes);
        assert!(text.contains("name=\"A &amp; B\""), "{text}");
        // Raw ampersand must not appear inside attributes.
        assert!(!text.contains("name=\"A & B\""), "raw & in attribute: {text}");
    }

    // 10. Formula text content is text-escaped.
    #[test]
    fn formula_text_escaped() {
        let mut wb = Workbook::new();
        wb.add_sheet(Worksheet::new("Sheet1"));
        wb.defined_names.push(DefinedName {
            name: "badFormula".to_string(),
            formula: "Sheet1!A1 & B1".to_string(),
            scope_sheet_index: None,
            builtin: None,
            hidden: false,
        });
        let bytes = emit(&wb);
        parse_ok(&bytes);
        let text = text_of(&bytes);
        assert!(
            text.contains("Sheet1!A1 &amp; B1"),
            "expected escaped & in formula text: {text}"
        );
    }

    // 11. Builtin prefix replaces the name attribute.
    #[test]
    fn builtin_prefix_applied_correctly() {
        let mut wb = Workbook::new();
        wb.add_sheet(Worksheet::new("Sheet1"));
        wb.defined_names.push(DefinedName {
            name: "foo".to_string(),
            formula: "Sheet1!A1:B2".to_string(),
            scope_sheet_index: Some(0),
            builtin: Some(BuiltinName::PrintArea),
            hidden: false,
        });
        let bytes = emit(&wb);
        parse_ok(&bytes);
        let text = text_of(&bytes);
        // The rendered name must be the xlnm prefix, not the user name.
        assert!(
            text.contains("name=\"_xlnm.Print_Area\""),
            "expected _xlnm.Print_Area: {text}"
        );
        // The original user name "foo" must not appear as a name= attribute.
        assert!(
            !text.contains("name=\"foo\""),
            "user name 'foo' should not appear: {text}"
        );
    }

    // 12. Hidden flag emitted.
    #[test]
    fn hidden_flag_emitted() {
        let mut wb = Workbook::new();
        wb.add_sheet(Worksheet::new("Sheet1"));
        wb.defined_names.push(DefinedName {
            name: "secret".to_string(),
            formula: "Sheet1!A1".to_string(),
            scope_sheet_index: None,
            builtin: None,
            hidden: true,
        });
        let bytes = emit(&wb);
        parse_ok(&bytes);
        let text = text_of(&bytes);
        assert!(text.contains("hidden=\"1\""), "{text}");
    }
}
