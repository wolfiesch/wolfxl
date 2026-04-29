//! `<hyperlinks>` emitter for worksheet XML.

use crate::model::worksheet::Worksheet;
use crate::xml_escape;

/// Emit `<hyperlinks>…</hyperlinks>`.
///
/// Relationship ids must stay aligned with [`crate::emit::rels::emit_sheet`]:
/// comments reserve `rId1` and `rId2`, table parts follow, then external
/// hyperlinks consume the remaining contiguous ids.
pub fn emit(out: &mut String, sheet: &Worksheet) {
    let comments_offset: u32 = if !sheet.comments.is_empty() { 2 } else { 0 };
    let tables_offset: u32 = sheet.tables.len() as u32;
    let mut rid = comments_offset + tables_offset + 1;

    out.push_str("<hyperlinks>");

    for (cell_ref, hyperlink) in &sheet.hyperlinks {
        if hyperlink.is_internal {
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
