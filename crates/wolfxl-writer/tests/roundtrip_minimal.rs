//! Wave 2 integration gate: end-to-end xlsx synthesis + re-parse.
//!
//! Builds a small `Workbook` with every interesting cell type, a merge,
//! a freeze pane, a defined name, and a styled cell — runs every Wave 1+2
//! emitter + the ZIP packager, opens the resulting bytes with
//! `zip::ZipArchive`, parses every extracted XML part with
//! `quick_xml::Reader`, and asserts cell values round-trip through the SST.
//!
//! If this test fails, Wave 2 is not complete. It is deliberately verbose
//! — each assertion doubles as documentation of what one full emit cycle
//! is supposed to produce.

use std::collections::HashMap;
use std::io::{Cursor, Read};

use quick_xml::events::Event;
use quick_xml::Reader;

use wolfxl_writer::emit::{
    content_types, doc_props, rels, shared_strings_xml, sheet_xml, styles_xml, workbook_xml,
};
use wolfxl_writer::model::cell::{FormulaResult, WriteCell, WriteCellValue};
use wolfxl_writer::model::defined_name::DefinedName;
use wolfxl_writer::model::format::{FontSpec, FormatSpec};
use wolfxl_writer::model::workbook::Workbook;
use wolfxl_writer::model::worksheet::{FreezePane, Merge, Worksheet};
use wolfxl_writer::zip::{package, ZipEntry};

/// Build the fixture: two sheets, mixed cells, merge, freeze, one defined
/// name, one styled cell. Returns the assembled `Workbook` *plus* the xf
/// index of the bold-red style so the caller can cross-check the `s` attr.
fn build_fixture() -> (Workbook, u32) {
    let mut wb = Workbook::new();

    // Intern a red-bold font into the styles table so we can point a cell at it.
    let bold_red = FormatSpec {
        font: Some(FontSpec {
            bold: true,
            color_rgb: Some("FFFF0000".to_string()),
            ..Default::default()
        }),
        ..Default::default()
    };
    let style_id = wb.styles.intern_format(&bold_red);

    // Sheet 1: every cell type.
    let mut s1 = Worksheet::new("Data");
    s1.set_cell(
        1,
        1,
        WriteCell::new(WriteCellValue::String("Name".to_string())),
    );
    s1.set_cell(
        1,
        2,
        WriteCell::new(WriteCellValue::String("Count".to_string())),
    );
    s1.set_cell(
        2,
        1,
        WriteCell::new(WriteCellValue::String("apples".to_string())),
    );
    s1.set_cell(2, 2, WriteCell::new(WriteCellValue::Number(42.0)));
    s1.set_cell(
        3,
        1,
        WriteCell::new(WriteCellValue::String("pears".to_string())),
    );
    s1.set_cell(3, 2, WriteCell::new(WriteCellValue::Number(3.5)));
    s1.set_cell(4, 1, WriteCell::new(WriteCellValue::Boolean(true)));
    s1.set_cell(
        4,
        2,
        WriteCell::new(WriteCellValue::Formula {
            expr: "SUM(B2:B3)".to_string(),
            result: Some(FormulaResult::Number(45.5)),
        }),
    );
    // Styled cell — points into the xf table.
    s1.set_cell(
        5,
        1,
        WriteCell::new(WriteCellValue::String("styled".to_string())).with_style(style_id),
    );
    // A genuinely blank styled cell to prove blank+style still emits.
    s1.set_cell(5, 2, WriteCell::new(WriteCellValue::Blank).with_style(style_id));

    s1.merge(Merge {
        top_row: 1,
        left_col: 1,
        bottom_row: 1,
        right_col: 2,
    });
    s1.freeze = Some(FreezePane {
        freeze_row: 1,
        freeze_col: 0,
        top_left: None,
    });
    wb.add_sheet(s1);

    // Sheet 2: one string shared with sheet 1 ("apples") to exercise SST dedup.
    let mut s2 = Worksheet::new("Summary");
    s2.set_cell(
        1,
        1,
        WriteCell::new(WriteCellValue::String("apples".to_string())),
    );
    s2.set_cell(1, 2, WriteCell::new(WriteCellValue::Number(100.0)));
    wb.add_sheet(s2);

    // Workbook-scope defined name.
    wb.defined_names.push(DefinedName {
        name: "Grand_Total".to_string(),
        formula: "Summary!$B$1".to_string(),
        scope_sheet_index: None,
        builtin: None,
        hidden: false,
    });

    (wb, style_id)
}

/// Emit every Wave 1+2 part, package the archive, return the raw xlsx bytes.
fn emit_full_pipeline(wb: &mut Workbook) -> Vec<u8> {
    // Sheet emission mutates the SST — must run before the SST emitter.
    let mut sheet_parts: Vec<(String, Vec<u8>)> = Vec::new();
    for (idx, sheet) in wb.sheets.iter().enumerate() {
        let bytes = sheet_xml::emit(sheet, idx as u32, &mut wb.sst, &wb.styles);
        sheet_parts.push((format!("xl/worksheets/sheet{}.xml", idx + 1), bytes));
    }

    // Canonical ZIP order — matches the one documented in zip.rs.
    let mut entries: Vec<ZipEntry> = vec![
        ZipEntry {
            path: "[Content_Types].xml".to_string(),
            bytes: content_types::emit(wb),
        },
        ZipEntry {
            path: "_rels/.rels".to_string(),
            bytes: rels::emit_root(wb),
        },
        ZipEntry {
            path: "xl/workbook.xml".to_string(),
            bytes: workbook_xml::emit(wb),
        },
        ZipEntry {
            path: "xl/_rels/workbook.xml.rels".to_string(),
            bytes: rels::emit_workbook(wb),
        },
    ];
    for (path, bytes) in sheet_parts {
        entries.push(ZipEntry { path, bytes });
    }
    // Per-sheet rels — only include non-empty ones (empty means no hyperlinks,
    // comments, or tables). This fixture has none, so skip all.
    for idx in 0..wb.sheets.len() {
        let sheet_rels = rels::emit_sheet(wb, idx);
        if !sheet_rels.is_empty() {
            entries.push(ZipEntry {
                path: format!("xl/worksheets/_rels/sheet{}.xml.rels", idx + 1),
                bytes: sheet_rels,
            });
        }
    }
    entries.extend([
        ZipEntry {
            path: "xl/styles.xml".to_string(),
            bytes: styles_xml::emit(&wb.styles),
        },
        ZipEntry {
            path: "xl/sharedStrings.xml".to_string(),
            bytes: shared_strings_xml::emit(&wb.sst),
        },
        ZipEntry {
            path: "docProps/core.xml".to_string(),
            bytes: doc_props::emit_core(wb),
        },
        ZipEntry {
            path: "docProps/app.xml".to_string(),
            bytes: doc_props::emit_app(wb),
        },
    ]);

    package(&entries).expect("zip package")
}

/// Open an xlsx byte slice and return a map of path → bytes for every entry.
fn read_archive(bytes: &[u8]) -> HashMap<String, Vec<u8>> {
    let cursor = Cursor::new(bytes);
    let mut archive = zip::ZipArchive::new(cursor).expect("open zip");
    let mut out = HashMap::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip entry");
        let path = file.name().to_string();
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read entry");
        out.insert(path, buf);
    }
    out
}

/// Parse raw bytes with quick_xml end-to-end, panicking on any XML-level error.
fn assert_xml_well_formed(part: &str, bytes: &[u8]) {
    let text = std::str::from_utf8(bytes).unwrap_or_else(|e| panic!("{part} not utf8: {e}"));
    let mut reader = Reader::from_str(text);
    reader.config_mut().check_end_names = true;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Err(e) => panic!("{part} parse error: {e}"),
            _ => (),
        }
        buf.clear();
    }
}

/// Walk sharedStrings.xml and return the interned strings in index order.
fn parse_sst(bytes: &[u8]) -> Vec<String> {
    let text = std::str::from_utf8(bytes).expect("sst utf8");
    let mut reader = Reader::from_str(text);
    reader.config_mut().check_end_names = true;
    let mut buf = Vec::new();
    let mut out = Vec::new();
    let mut in_t = false;
    let mut current = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"t" => {
                in_t = true;
                current.clear();
            }
            Ok(Event::End(e)) if e.name().as_ref() == b"t" => {
                in_t = false;
                out.push(std::mem::take(&mut current));
            }
            Ok(Event::Text(t)) if in_t => {
                let decoded = t.unescape().expect("unescape sst");
                current.push_str(&decoded);
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("sst parse error: {e}"),
            _ => (),
        }
        buf.clear();
    }
    out
}

#[test]
fn wave2_full_pipeline_roundtrip() {
    let (mut wb, style_id) = build_fixture();
    let bytes = emit_full_pipeline(&mut wb);

    // 1. Archive is openable and contains every required part.
    let parts = read_archive(&bytes);
    for required in [
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/sheet2.xml",
        "xl/styles.xml",
        "xl/sharedStrings.xml",
        "docProps/core.xml",
        "docProps/app.xml",
    ] {
        assert!(
            parts.contains_key(required),
            "archive missing required part: {required}"
        );
    }

    // 2. Every XML part is well-formed.
    for (path, data) in &parts {
        assert_xml_well_formed(path, data);
    }

    // 3. SST contains exactly the distinct strings we interned, in
    //    first-use order. "apples" appears on both sheets but only once
    //    in the table; the total reference count is higher than unique.
    let sst = parse_sst(&parts["xl/sharedStrings.xml"]);
    assert!(sst.contains(&"Name".to_string()), "SST missing Name");
    assert!(sst.contains(&"Count".to_string()), "SST missing Count");
    assert!(sst.contains(&"apples".to_string()), "SST missing apples");
    assert!(sst.contains(&"pears".to_string()), "SST missing pears");
    assert!(sst.contains(&"styled".to_string()), "SST missing styled");
    assert_eq!(
        sst.iter().filter(|s| *s == "apples").count(),
        1,
        "SST should dedupe 'apples' across sheets"
    );

    // 4. Sheet1 has the right structural hooks.
    let sheet1 = std::str::from_utf8(&parts["xl/worksheets/sheet1.xml"]).unwrap();
    assert!(sheet1.contains("<mergeCell ref=\"A1:B1\""), "{sheet1}");
    // Freeze pane present.
    assert!(
        sheet1.contains("<pane "),
        "sheet1 should emit <pane> for freeze: {sheet1}"
    );
    // Formula emitted with cached result.
    assert!(sheet1.contains("<f>SUM(B2:B3)</f>"), "{sheet1}");
    assert!(sheet1.contains("<v>45.5</v>"), "{sheet1}");
    // Boolean emitted as t="b" with <v>1</v>.
    assert!(sheet1.contains("t=\"b\""), "{sheet1}");
    // Styled cell points at the interned xf id.
    let s_attr = format!("s=\"{style_id}\"");
    assert!(
        sheet1.contains(&s_attr),
        "styled cell should have {s_attr}: {sheet1}"
    );

    // 5. Sheet2 re-uses the "apples" SST index. Parse the string-typed
    //    cell and verify the index maps back to "apples".
    let sheet2 = std::str::from_utf8(&parts["xl/worksheets/sheet2.xml"]).unwrap();
    let apples_idx = sst.iter().position(|s| s == "apples").expect("apples in SST");
    let expected = format!("t=\"s\"><v>{apples_idx}</v>");
    assert!(
        sheet2.contains(&expected),
        "sheet2 should reference SST index {apples_idx} for 'apples': {sheet2}"
    );

    // 6. Workbook lists the defined name and both sheets.
    let wb_xml = std::str::from_utf8(&parts["xl/workbook.xml"]).unwrap();
    assert!(wb_xml.contains("<sheet name=\"Data\""), "{wb_xml}");
    assert!(wb_xml.contains("<sheet name=\"Summary\""), "{wb_xml}");
    assert!(
        wb_xml.contains("<definedName name=\"Grand_Total\">Summary!$B$1</definedName>"),
        "{wb_xml}"
    );

    // 7. Styles part surfaces our bold-red font. The exact font index is
    //    implementation-detail, but a bold=1 + color FFFF0000 pair must
    //    appear in the <fonts> block.
    let styles = std::str::from_utf8(&parts["xl/styles.xml"]).unwrap();
    assert!(styles.contains("<b/>"), "bold tag missing: {styles}");
    assert!(
        styles.contains("FFFF0000") || styles.contains("ffff0000"),
        "red color missing from styles: {styles}"
    );
}
