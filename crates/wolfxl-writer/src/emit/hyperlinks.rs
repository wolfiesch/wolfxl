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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::comment::Comment;
    use crate::model::worksheet::Hyperlink;

    fn external(target: &str) -> Hyperlink {
        Hyperlink {
            target: target.into(),
            is_internal: false,
            display: None,
            tooltip: None,
        }
    }

    #[test]
    fn external_hyperlink_starts_at_rid1_without_comments_or_tables() {
        let mut sheet = Worksheet::new("S");
        sheet
            .hyperlinks
            .insert("A1".to_string(), external("https://example.com"));
        let mut out = String::new();

        emit(&mut out, &sheet);

        assert_eq!(
            out,
            "<hyperlinks><hyperlink ref=\"A1\" r:id=\"rId1\"/></hyperlinks>"
        );
    }

    #[test]
    fn external_hyperlink_starts_at_rid3_when_comments_reserve_two_rids() {
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
        sheet
            .hyperlinks
            .insert("B1".to_string(), external("https://example.com"));
        let mut out = String::new();

        emit(&mut out, &sheet);

        assert_eq!(
            out,
            "<hyperlinks><hyperlink ref=\"B1\" r:id=\"rId3\"/></hyperlinks>"
        );
    }

    #[test]
    fn internal_hyperlink_uses_location_without_relationship_id() {
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
        let mut out = String::new();

        emit(&mut out, &sheet);

        assert_eq!(
            out,
            "<hyperlinks><hyperlink ref=\"A1\" location=\"Sheet2!A1\"/></hyperlinks>"
        );
        assert!(!out.contains("r:id="), "internal link has no r:id: {out}");
    }
}
