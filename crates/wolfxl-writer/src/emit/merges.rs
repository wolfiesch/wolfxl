//! `<mergeCells>` emitter for worksheet XML.

use crate::model::worksheet::Worksheet;
use crate::refs;

/// Emit `<mergeCells count="N">…</mergeCells>`.
///
/// The ranges are sorted by top-left coordinate to keep worksheet XML
/// deterministic even when merges were registered in a different order.
pub fn emit(out: &mut String, sheet: &Worksheet) {
    let mut merges = sheet.merges.clone();
    merges.sort_by(|a, b| a.top_row.cmp(&b.top_row).then(a.left_col.cmp(&b.left_col)));

    out.push_str(&format!("<mergeCells count=\"{}\">", merges.len()));

    for merge in &merges {
        let range = refs::format_range(
            (merge.top_row, merge.left_col),
            (merge.bottom_row, merge.right_col),
        );
        out.push_str(&format!("<mergeCell ref=\"{}\"/>", range));
    }

    out.push_str("</mergeCells>");
}
