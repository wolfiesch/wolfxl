//! `xl/drawings/vmlDrawing{N}.vml` emitter — legacy VML anchor shapes for
//! comment boxes. Wave 3A.
//!
//! VML is an ancient Microsoft XML dialect kept around specifically for
//! comment boxes and form controls. Modern OOXML uses DrawingML for
//! everything else, but comments still need VML to show Excel a
//! yellow-rectangle shape anchored to a cell.

use crate::model::worksheet::Worksheet;
use crate::refs;

pub fn emit(sheet: &Worksheet) -> Vec<u8> {
    if sheet.comments.is_empty() {
        return Vec::new();
    }

    let mut out = String::with_capacity(2048);

    // No <?xml?> declaration — VML convention; the root element is literally <xml>
    out.push_str(
        "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\
 xmlns:o=\"urn:schemas-microsoft-com:office:office\"\
 xmlns:x=\"urn:schemas-microsoft-com:office:excel\">",
    );

    // Shape layout
    out.push_str(
        "<o:shapelayout v:ext=\"edit\">\
<o:idmap v:ext=\"edit\" data=\"1\"/>\
</o:shapelayout>",
    );

    // Shape type definition (shared for all comment boxes)
    out.push_str(
        "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"\
 path=\"m,l,21600r21600,l21600,xe\">\
<v:stroke joinstyle=\"miter\"/>\
<v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\
</v:shapetype>",
    );

    // One <v:shape> per comment — BTreeMap gives A1 order
    for (idx, (cell_ref, comment)) in sheet.comments.iter().enumerate() {
        let shape_num = 1024 + (idx as u32) + 1; // starts at 1025

        // Parse row/col from the A1 key (1-based), convert to 0-based
        let (row0, col0) = match refs::parse_a1(cell_ref) {
            Some((r, c)) => (r - 1, c - 1),
            None => (0, 0), // fallback for malformed refs (should not happen)
        };

        let width = match comment.width_pt {
            Some(w) => format!("{}pt", w),
            None => "96pt".to_string(),
        };
        let height = match comment.height_pt {
            Some(h) => format!("{}pt", h),
            None => "55.5pt".to_string(),
        };

        out.push_str(&format!(
            "<v:shape id=\"_x0000_s{}\" type=\"#_x0000_t202\"\
 style=\"position:absolute; margin-left:59.25pt; margin-top:1.5pt;\
 width:{}; height:{}; z-index:1; visibility:hidden\"\
 fillcolor=\"#ffffe1\" o:insetmode=\"auto\">",
            shape_num, width, height,
        ));

        out.push_str("<v:fill color2=\"#ffffe1\"/>");
        out.push_str("<v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>");
        out.push_str("<v:path o:connecttype=\"none\"/>");
        out.push_str("<v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"/></v:textbox>");
        out.push_str("<x:ClientData ObjectType=\"Note\">");
        out.push_str("<x:MoveWithCells/>");
        out.push_str("<x:SizeWithCells/>");
        out.push_str("<x:Anchor>2, 0, 0, 0, 3, 20, 3, 15</x:Anchor>");
        out.push_str("<x:AutoFill>False</x:AutoFill>");
        out.push_str(&format!("<x:Row>{}</x:Row>", row0));
        out.push_str(&format!("<x:Column>{}</x:Column>", col0));

        // <x:Visible/> is present only when visible == true
        if comment.visible {
            out.push_str("<x:Visible/>");
        }

        out.push_str("</x:ClientData>");
        out.push_str("</v:shape>");
    }

    out.push_str("</xml>");

    out.into_bytes()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::comment::Comment;
    use crate::model::worksheet::Worksheet;
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

    fn make_comment(visible: bool, width_pt: Option<f64>, height_pt: Option<f64>) -> Comment {
        Comment {
            text: "note".to_string(),
            author_id: 0,
            width_pt,
            height_pt,
            visible,
        }
    }

    // 1. Empty sheet returns empty bytes
    #[test]
    fn empty_sheet_returns_empty_bytes() {
        let sheet = Worksheet::new("S");
        let result = emit(&sheet);
        assert!(result.is_empty(), "expected empty Vec, got {} bytes", result.len());
    }

    // 2. Root declares three namespaces
    #[test]
    fn root_declares_three_namespaces() {
        let mut sheet = Worksheet::new("S");
        sheet.comments.insert("A1".to_string(), make_comment(false, None, None));
        let bytes = emit(&sheet);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("xmlns:v="), "xmlns:v missing: {text}");
        assert!(text.contains("xmlns:o="), "xmlns:o missing: {text}");
        assert!(text.contains("xmlns:x="), "xmlns:x missing: {text}");
    }

    // 3. Shape IDs start at 1025 and increment
    #[test]
    fn shape_ids_start_at_1025_and_increment() {
        let mut sheet = Worksheet::new("S");
        // A1 < B2 — BTreeMap gives them in this order
        sheet.comments.insert("A1".to_string(), make_comment(false, None, None));
        sheet.comments.insert("B2".to_string(), make_comment(false, None, None));

        let bytes = emit(&sheet);
        let text = String::from_utf8(bytes).unwrap();

        assert!(text.contains("id=\"_x0000_s1025\""), "first shape 1025: {text}");
        assert!(text.contains("id=\"_x0000_s1026\""), "second shape 1026: {text}");
    }

    // 4. Visible true emits element; visible false omits it
    #[test]
    fn visible_true_emits_visible_element() {
        // visible=true
        let mut sheet_visible = Worksheet::new("S");
        sheet_visible.comments.insert("A1".to_string(), make_comment(true, None, None));
        let bytes_v = emit(&sheet_visible);
        let text_v = String::from_utf8(bytes_v).unwrap();
        assert!(text_v.contains("<x:Visible/>"), "visible=true needs <x:Visible/>: {text_v}");

        // visible=false
        let mut sheet_hidden = Worksheet::new("S");
        sheet_hidden.comments.insert("A1".to_string(), make_comment(false, None, None));
        let bytes_h = emit(&sheet_hidden);
        let text_h = String::from_utf8(bytes_h).unwrap();
        assert!(!text_h.contains("<x:Visible"), "visible=false must omit <x:Visible: {text_h}");
    }

    // 5. Row and column are 0-based
    #[test]
    fn row_column_are_zero_based() {
        let mut sheet = Worksheet::new("S");
        // C3 = row 3, col 3 (1-based) → <x:Row>2</x:Row>, <x:Column>2</x:Column>
        sheet.comments.insert("C3".to_string(), make_comment(false, None, None));
        let bytes = emit(&sheet);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("<x:Row>2</x:Row>"), "row 0-based: {text}");
        assert!(text.contains("<x:Column>2</x:Column>"), "col 0-based: {text}");
    }

    // 6. Width and height flow into style
    #[test]
    fn width_height_flow_into_style() {
        let mut sheet = Worksheet::new("S");
        sheet.comments.insert("A1".to_string(), make_comment(false, Some(150.0), Some(80.0)));
        let bytes = emit(&sheet);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("width:150pt"), "custom width: {text}");
        assert!(text.contains("height:80pt"), "custom height: {text}");
    }

    // 7. Well-formed under quick_xml with multiple comments
    #[test]
    fn well_formed_under_quick_xml() {
        let mut sheet = Worksheet::new("S");
        sheet.comments.insert("A1".to_string(), make_comment(true, None, None));
        sheet.comments.insert("B2".to_string(), make_comment(false, Some(120.0), Some(60.0)));
        sheet.comments.insert("C3".to_string(), make_comment(true, None, Some(45.0)));

        let bytes = emit(&sheet);
        assert!(!bytes.is_empty());
        parse_ok(&bytes);
    }
}
