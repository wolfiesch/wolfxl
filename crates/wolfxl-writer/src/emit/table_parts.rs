//! `<tableParts>` emitter for worksheet XML.

use crate::model::worksheet::Worksheet;

/// Emit `<tableParts count="N">...<tablePart r:id="rIdX"/>...</tableParts>`.
///
/// Sheet relationships reserve `rId1` and `rId2` for comments/VML when
/// comments exist, so table rIds start at `rId3` in that case and `rId1`
/// otherwise.
pub fn emit(out: &mut String, sheet: &Worksheet) {
    if sheet.tables.is_empty() {
        return;
    }

    let comments_offset: u32 = if !sheet.comments.is_empty() { 2 } else { 0 };
    out.push_str(&format!("<tableParts count=\"{}\">", sheet.tables.len()));
    for (local_idx, _) in sheet.tables.iter().enumerate() {
        let rid = comments_offset + local_idx as u32 + 1;
        out.push_str(&format!("<tablePart r:id=\"rId{}\"/>", rid));
    }
    out.push_str("</tableParts>");
}
