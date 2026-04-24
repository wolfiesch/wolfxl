//! `.rels` emitter — root rels, workbook rels, and per-sheet rels.
//!
//! OOXML relationships are a layer of indirection between parts: every
//! `r:id="rId5"` attribute somewhere in the workbook is resolved through
//! a `.rels` file to the actual target part path.
//!
//! Three `.rels` files are emitted here:
//!
//! | File | Emitter | Purpose |
//! |------|---------|---------|
//! | `_rels/.rels` | [`emit_root`] | workbook + docProps/core + docProps/app |
//! | `xl/_rels/workbook.xml.rels` | [`emit_workbook`] | sheets + styles + SST |
//! | `xl/worksheets/_rels/sheet{N}.xml.rels` | [`emit_sheet`] | comments, vml, tables, external hyperlinks |
//!
//! The sheet-level rels file is only emitted when the sheet has at least
//! one of: comments, external hyperlinks, tables. Otherwise [`emit_sheet`]
//! returns an empty byte vector and the caller skips the file entirely.

use crate::model::workbook::Workbook;

// -- Relationship types ------------------------------------------------------
const RT_OFFICE_DOC: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const RT_CORE_PROPS: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const RT_EXT_PROPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
const RT_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const RT_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const RT_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const RT_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const RT_VML_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const RT_TABLE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
const RT_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

const RELS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

fn rels_header() -> String {
    let mut s = String::with_capacity(512);
    s.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    s.push_str(&format!("<Relationships xmlns=\"{RELS_NS}\">"));
    s
}

fn push_relationship(s: &mut String, id: &str, rtype: &str, target: &str) {
    s.push_str(&format!(
        "<Relationship Id=\"{id}\" Type=\"{rtype}\" Target=\"{}\"/>",
        xml_attr_escape(target),
    ));
}

fn push_external_relationship(s: &mut String, id: &str, rtype: &str, target: &str) {
    s.push_str(&format!(
        "<Relationship Id=\"{id}\" Type=\"{rtype}\" Target=\"{}\" TargetMode=\"External\"/>",
        xml_attr_escape(target),
    ));
}

/// XML attribute-value escape. Relationships store URLs in `Target="…"`
/// so `&`, `<`, `>`, and `"` must all be escaped.
fn xml_attr_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}

/// `_rels/.rels` — top-level relationships (workbook, core props, app props).
pub fn emit_root(_wb: &Workbook) -> Vec<u8> {
    let mut out = rels_header();
    push_relationship(&mut out, "rId1", RT_OFFICE_DOC, "xl/workbook.xml");
    push_relationship(&mut out, "rId2", RT_CORE_PROPS, "docProps/core.xml");
    push_relationship(&mut out, "rId3", RT_EXT_PROPS, "docProps/app.xml");
    out.push_str("</Relationships>");
    out.into_bytes()
}

/// `xl/_rels/workbook.xml.rels` — workbook → sheets, styles, shared strings.
pub fn emit_workbook(wb: &Workbook) -> Vec<u8> {
    let mut out = rels_header();
    let n_sheets = wb.sheets.len();
    for idx in 0..n_sheets {
        let sheet_n = idx + 1;
        let rid = format!("rId{sheet_n}");
        let target = format!("worksheets/sheet{sheet_n}.xml");
        push_relationship(&mut out, &rid, RT_WORKSHEET, &target);
    }
    // Styles is rId{N+1}, SharedStrings is rId{N+2}.
    let styles_rid = format!("rId{}", n_sheets + 1);
    push_relationship(&mut out, &styles_rid, RT_STYLES, "styles.xml");
    let sst_rid = format!("rId{}", n_sheets + 2);
    push_relationship(&mut out, &sst_rid, RT_SHARED_STRINGS, "sharedStrings.xml");
    out.push_str("</Relationships>");
    out.into_bytes()
}

/// `xl/worksheets/_rels/sheet{N}.xml.rels` — sheet → comments, drawings,
/// tables, external hyperlinks.
///
/// Returns an empty vec when the sheet has none of those things. The
/// caller is expected to detect `is_empty()` and skip the file.
///
/// # Relationship numbering
///
/// - If comments exist: `rId1` = comments, `rId2` = vmlDrawing.
/// - Tables occupy the next contiguous block of rIds.
/// - External hyperlinks (those that do not start with `#`) occupy the
///   tail of the rId range.
///
/// # Table numbering
///
/// Tables are numbered globally across the workbook (1..N) in sheet
/// insertion order. This matches openpyxl and [`super::content_types::emit`].
pub fn emit_sheet(wb: &Workbook, sheet_idx: usize) -> Vec<u8> {
    let Some(sheet) = wb.sheets.get(sheet_idx) else {
        return Vec::new();
    };

    let has_comments = !sheet.comments.is_empty();
    let external_hyperlinks: Vec<(&String, &crate::model::worksheet::Hyperlink)> = sheet
        .hyperlinks
        .iter()
        .filter(|(_, h)| !h.target.starts_with('#'))
        .collect();
    let has_tables = !sheet.tables.is_empty();

    if !has_comments && external_hyperlinks.is_empty() && !has_tables {
        return Vec::new();
    }

    // Global table numbering: count how many tables exist in sheets before
    // this one, then assign 1-based ids starting from there + 1.
    let tables_before: usize = wb.sheets[..sheet_idx].iter().map(|s| s.tables.len()).sum();

    let mut out = rels_header();
    let mut rid_counter: u32 = 0;

    // Comments + VML.
    if has_comments {
        rid_counter += 1;
        let n = sheet_idx + 1;
        push_relationship(
            &mut out,
            &format!("rId{rid_counter}"),
            RT_COMMENTS,
            &format!("../comments/comments{n}.xml"),
        );
        rid_counter += 1;
        push_relationship(
            &mut out,
            &format!("rId{rid_counter}"),
            RT_VML_DRAWING,
            &format!("../drawings/vmlDrawing{n}.vml"),
        );
    }

    // Tables.
    for (local_idx, _table) in sheet.tables.iter().enumerate() {
        rid_counter += 1;
        let global_id = tables_before + local_idx + 1;
        push_relationship(
            &mut out,
            &format!("rId{rid_counter}"),
            RT_TABLE,
            &format!("../tables/table{global_id}.xml"),
        );
    }

    // External hyperlinks (URLs, not internal `#Sheet!A1` references).
    for (_cell_ref, hyperlink) in &external_hyperlinks {
        rid_counter += 1;
        push_external_relationship(
            &mut out,
            &format!("rId{rid_counter}"),
            RT_HYPERLINK,
            &hyperlink.target,
        );
    }

    out.push_str("</Relationships>");
    out.into_bytes()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::comment::Comment;
    use crate::model::table::{Table, TableColumn};
    use crate::model::worksheet::{Hyperlink, Worksheet};
    use quick_xml::events::Event;
    use quick_xml::Reader;

    fn wb_with_sheets(n: usize) -> Workbook {
        let mut wb = Workbook::new();
        for i in 1..=n {
            wb.add_sheet(Worksheet::new(format!("Sheet{i}")));
        }
        wb
    }

    fn parse_ok(bytes: &[u8]) {
        let text = std::str::from_utf8(bytes).expect("utf8");
        let mut reader = Reader::from_str(text);
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

    #[test]
    fn root_rels_has_three_relationships() {
        let wb = wb_with_sheets(1);
        let bytes = emit_root(&wb);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("Id=\"rId1\""));
        assert!(text.contains("Id=\"rId2\""));
        assert!(text.contains("Id=\"rId3\""));
        assert!(text.contains("Target=\"xl/workbook.xml\""));
        assert!(text.contains("Target=\"docProps/core.xml\""));
        assert!(text.contains("Target=\"docProps/app.xml\""));
    }

    #[test]
    fn workbook_rels_numbers_sheets_then_styles_then_sst() {
        let wb = wb_with_sheets(3);
        let bytes = emit_workbook(&wb);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("Id=\"rId1\"") && text.contains("worksheets/sheet1.xml"));
        assert!(text.contains("Id=\"rId2\"") && text.contains("worksheets/sheet2.xml"));
        assert!(text.contains("Id=\"rId3\"") && text.contains("worksheets/sheet3.xml"));
        assert!(text.contains("Id=\"rId4\"") && text.contains("Target=\"styles.xml\""));
        assert!(text.contains("Id=\"rId5\"") && text.contains("Target=\"sharedStrings.xml\""));
    }

    #[test]
    fn sheet_rels_empty_when_nothing_to_reference() {
        let wb = wb_with_sheets(1);
        let bytes = emit_sheet(&wb, 0);
        assert!(bytes.is_empty(), "no rels needed for bare sheet");
    }

    #[test]
    fn sheet_rels_out_of_range_returns_empty() {
        let wb = wb_with_sheets(1);
        let bytes = emit_sheet(&wb, 5);
        assert!(bytes.is_empty());
    }

    #[test]
    fn sheet_rels_with_comments_emits_two_relationships() {
        let mut wb = wb_with_sheets(1);
        wb.sheets[0].comments.insert(
            "A1".into(),
            Comment {
                text: "hello".into(),
                author_id: 0,
                width_pt: None,
                height_pt: None,
                visible: false,
            },
        );
        let bytes = emit_sheet(&wb, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("../comments/comments1.xml"));
        assert!(text.contains("../drawings/vmlDrawing1.vml"));
        assert!(text.contains("Id=\"rId1\""));
        assert!(text.contains("Id=\"rId2\""));
    }

    #[test]
    fn sheet_rels_external_hyperlink_marked_external() {
        let mut wb = wb_with_sheets(1);
        wb.sheets[0].hyperlinks.insert(
            "A1".into(),
            Hyperlink {
                target: "https://example.com".into(),
                display: None,
                tooltip: None,
            },
        );
        let bytes = emit_sheet(&wb, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("TargetMode=\"External\""));
        assert!(text.contains("https://example.com"));
    }

    #[test]
    fn sheet_rels_internal_hyperlink_skipped() {
        let mut wb = wb_with_sheets(1);
        wb.sheets[0].hyperlinks.insert(
            "A1".into(),
            Hyperlink {
                target: "#Sheet2!A1".into(),
                display: None,
                tooltip: None,
            },
        );
        // No tables, no comments, only internal link → no rels file.
        let bytes = emit_sheet(&wb, 0);
        assert!(bytes.is_empty());
    }

    #[test]
    fn sheet_rels_tables_use_global_numbering() {
        let mut wb = wb_with_sheets(2);
        let mk = || Table {
            name: "t".into(),
            display_name: None,
            range: "A1:B2".into(),
            columns: vec![TableColumn {
                name: "a".into(),
                totals_function: None,
                totals_label: None,
            }],
            header_row: true,
            totals_row: false,
            style: None,
            autofilter: false,
        };
        wb.sheets[0].tables.push(mk());
        wb.sheets[0].tables.push(mk());
        wb.sheets[1].tables.push(mk());

        let bytes0 = emit_sheet(&wb, 0);
        let text0 = String::from_utf8(bytes0).unwrap();
        assert!(text0.contains("../tables/table1.xml"));
        assert!(text0.contains("../tables/table2.xml"));

        let bytes1 = emit_sheet(&wb, 1);
        let text1 = String::from_utf8(bytes1).unwrap();
        assert!(text1.contains("../tables/table3.xml"));
        assert!(!text1.contains("../tables/table1.xml"));
    }

    #[test]
    fn sheet_rels_all_three_kinds_coexist() {
        let mut wb = wb_with_sheets(1);
        wb.sheets[0].comments.insert(
            "A1".into(),
            Comment {
                text: "n".into(),
                author_id: 0,
                width_pt: None,
                height_pt: None,
                visible: false,
            },
        );
        wb.sheets[0].tables.push(Table {
            name: "t".into(),
            display_name: None,
            range: "A1:B2".into(),
            columns: vec![TableColumn {
                name: "a".into(),
                totals_function: None,
                totals_label: None,
            }],
            header_row: true,
            totals_row: false,
            style: None,
            autofilter: false,
        });
        wb.sheets[0].hyperlinks.insert(
            "B1".into(),
            Hyperlink {
                target: "https://example.com/path?q=1&r=2".into(),
                display: None,
                tooltip: None,
            },
        );
        let bytes = emit_sheet(&wb, 0);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        // rId1 = comments, rId2 = vml, rId3 = table, rId4 = hyperlink
        assert!(text.contains("Id=\"rId1\""));
        assert!(text.contains("Id=\"rId2\""));
        assert!(text.contains("Id=\"rId3\""));
        assert!(text.contains("Id=\"rId4\""));
        // Ampersand in URL must be escaped.
        assert!(text.contains("q=1&amp;r=2"));
    }
}
