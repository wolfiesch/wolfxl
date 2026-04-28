//! `calcchain` — rebuild `xl/calcChain.xml` from the post-mutation
//! sheet bytes (Sprint Θ Pod-C3).
//!
//! Excel writes a calc-chain hint that lists every formula cell in
//! the order they should be evaluated. WolfXL historically left the
//! source file's calcChain alone (modify-mode) or omitted it entirely
//! (write-mode); Excel transparently rebuilds it on next open, so the
//! end-user impact is a one-time recompute on the first open after
//! saving.
//!
//! This module exposes the pure scanning + emission helpers; the
//! patcher's flush phase calls them after every sheet mutation has
//! settled. The output is byte-deterministic for a given (sheet
//! tab-order, formula cells per sheet) tuple — important for the
//! diff-test infrastructure.
//!
//! # Format
//!
//! ```xml
//! <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
//! <calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
//!   <c r="A1" i="1"/>
//!   <c r="B2" i="1"/>
//!   <c r="C3" i="2"/>
//! </calcChain>
//! ```
//!
//! Where `r` is the cell A1 reference and `i` is the 1-based sheet
//! index (matches `<sheet sheetId="N">` declaration order in
//! `xl/workbook.xml`).

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

/// Content type for `xl/calcChain.xml`.
pub const CT_CALC_CHAIN: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";

/// Relationship type for the workbook → calcChain edge.
pub const REL_CALC_CHAIN: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

/// One calcChain entry: cell reference + 1-based sheet index.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CalcChainEntry {
    /// A1 reference, e.g. `"A1"`, `"BC42"`. Verbatim copy of the
    /// `<c r="…">` attribute on the formula cell.
    pub cell_ref: String,
    /// 1-based sheet index — position in the workbook's tab list.
    pub sheet_index: u32,
}

/// Scan a single sheet's XML and emit one [`CalcChainEntry`] for each
/// `<c>` element that has at least one `<f>` child.
///
/// Quick-XML based; tolerant of namespaces and arbitrary whitespace.
/// Skips cells that have no formula.
pub fn scan_sheet_for_formulas(sheet_xml: &[u8], sheet_index_1based: u32) -> Vec<CalcChainEntry> {
    let mut reader = XmlReader::from_reader(sheet_xml);
    reader.config_mut().trim_text(false);

    let mut entries: Vec<CalcChainEntry> = Vec::new();
    let mut buf: Vec<u8> = Vec::new();

    // The <c> we're currently inside, if any. We commit it when we
    // see an <f> child before the closing </c>.
    let mut current_c_ref: Option<String> = None;
    let mut emitted_for_current: bool = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = name.local_name();
                match local.as_ref() {
                    b"c" => {
                        // Pull r="..." attribute.
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"r" {
                                if let Ok(v) = attr.unescape_value() {
                                    current_c_ref = Some(v.into_owned());
                                }
                            }
                        }
                        emitted_for_current = false;
                    }
                    b"f" => {
                        if let Some(r) = current_c_ref.as_ref() {
                            if !emitted_for_current {
                                entries.push(CalcChainEntry {
                                    cell_ref: r.clone(),
                                    sheet_index: sheet_index_1based,
                                });
                                emitted_for_current = true;
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                // Self-closing tags. <c r="A1" t="str"/> with no <f>
                // is a value cell, not a formula — skip. <f/> by
                // itself is unusual but we treat it the same as Start
                // for safety.
                let name = e.name();
                let local = name.local_name();
                match local.as_ref() {
                    b"c" => {
                        // Self-closing <c .../> has no children → no formula.
                        current_c_ref = None;
                        emitted_for_current = false;
                    }
                    b"f" => {
                        if let Some(r) = current_c_ref.as_ref() {
                            if !emitted_for_current {
                                entries.push(CalcChainEntry {
                                    cell_ref: r.clone(),
                                    sheet_index: sheet_index_1based,
                                });
                                emitted_for_current = true;
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                if e.name().local_name().as_ref() == b"c" {
                    current_c_ref = None;
                    emitted_for_current = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break, // best-effort: a parse failure means we
            // skip this sheet rather than abort the
            // save.
            _ => {}
        }
        buf.clear();
    }
    entries
}

/// Render `xl/calcChain.xml` from a flat list of entries. Order is
/// preserved verbatim — caller is responsible for the iteration order
/// (typically sheet tab-order, then sheet-XML scan order).
///
/// Returns `None` if there are no entries — the caller should DELETE
/// the calcChain.xml part instead of writing an empty one.
pub fn render_calc_chain(entries: &[CalcChainEntry]) -> Option<Vec<u8>> {
    if entries.is_empty() {
        return None;
    }
    let mut out = String::new();
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    out.push_str("<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
    for e in entries {
        // Escape `&`, `<`, `"` in the cell ref (defensive — Excel
        // never emits non-A1 refs but a corrupt source could).
        let escaped = e
            .cell_ref
            .replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('"', "&quot;");
        out.push_str(&format!("<c r=\"{}\" i=\"{}\"/>", escaped, e.sheet_index));
    }
    out.push_str("</calcChain>");
    Some(out.into_bytes())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scan_finds_formula_cells_only() {
        let xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1"><f>SUM(C1:D1)</f><v>10</v></c>
      <c r="C1"><v>4</v></c>
      <c r="D1"><f>A1+1</f><v>2</v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let entries = scan_sheet_for_formulas(xml.as_bytes(), 1);
        assert_eq!(entries.len(), 2, "expected 2 formula cells");
        assert_eq!(entries[0].cell_ref, "B1");
        assert_eq!(entries[0].sheet_index, 1);
        assert_eq!(entries[1].cell_ref, "D1");
        assert_eq!(entries[1].sheet_index, 1);
    }

    #[test]
    fn scan_skips_self_closing_value_cells() {
        let xml = br#"<sheetData>
            <row r="1"><c r="A1" t="str"/></row>
        </sheetData>"#;
        assert_eq!(scan_sheet_for_formulas(xml, 7), vec![]);
    }

    #[test]
    fn scan_emits_one_entry_per_cell_even_with_multiple_f() {
        // Defensive: normally a <c> contains at most one <f>, but if a
        // shared formula and a normal <f> co-exist, we shouldn't
        // double-count.
        let xml = br#"<sheetData>
            <row r="1"><c r="A1"><f t="shared" si="0"/><f>A1+1</f><v>0</v></c></row>
        </sheetData>"#;
        let entries = scan_sheet_for_formulas(xml, 1);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].cell_ref, "A1");
    }

    #[test]
    fn render_empty_returns_none() {
        assert!(render_calc_chain(&[]).is_none());
    }

    #[test]
    fn render_emits_canonical_shape() {
        let entries = vec![
            CalcChainEntry {
                cell_ref: "A1".into(),
                sheet_index: 1,
            },
            CalcChainEntry {
                cell_ref: "B2".into(),
                sheet_index: 2,
            },
        ];
        let bytes = render_calc_chain(&entries).expect("non-empty");
        let s = String::from_utf8(bytes).unwrap();
        assert!(s.contains("<?xml"));
        assert!(s.contains("<calcChain"));
        assert!(s.contains("<c r=\"A1\" i=\"1\"/>"));
        assert!(s.contains("<c r=\"B2\" i=\"2\"/>"));
        assert!(s.contains("</calcChain>"));
    }
}
