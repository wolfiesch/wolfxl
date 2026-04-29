//! Legacy `<sheetViews>` emitter for worksheet XML.

use crate::model::worksheet::{FreezePane, SplitPane, Worksheet};
use crate::refs;

/// Emit `<sheetViews><sheetView ...>...</sheetView></sheetViews>`.
///
/// Typed sheet-view specs are emitted by `parse::sheet_setup`; this module is
/// the legacy fallback used by write-mode freeze/split pane fields.
pub fn emit(out: &mut String, sheet: &Worksheet, sheet_idx: u32) {
    out.push_str("<sheetViews>");

    if sheet_idx == 0 {
        out.push_str("<sheetView tabSelected=\"1\" workbookViewId=\"0\">");
    } else {
        out.push_str("<sheetView workbookViewId=\"0\">");
    }

    if let Some(freeze) = &sheet.freeze {
        emit_freeze_pane(out, freeze);
    } else if let Some(split) = &sheet.split {
        emit_split_pane(out, split);
    }

    out.push_str("</sheetView>");
    out.push_str("</sheetViews>");
}

/// Emit `<pane .../>` for a freeze pane.
///
/// OOXML pane semantics with `state="frozen"`:
/// `xSplit` = number of columns frozen and `ySplit` = number of rows frozen.
/// The model stores the freeze cell coordinate, so the emitted count is
/// `(coord - 1)`.
fn emit_freeze_pane(out: &mut String, freeze: &FreezePane) {
    let y_split = freeze.freeze_row.saturating_sub(1);
    let x_split = freeze.freeze_col.saturating_sub(1);
    let has_row = y_split > 0;
    let has_col = x_split > 0;
    if !has_row && !has_col {
        return;
    }

    let tl_row =
        freeze
            .top_left
            .map(|t| t.0)
            .unwrap_or_else(|| if has_row { freeze.freeze_row } else { 1 });
    let tl_col =
        freeze
            .top_left
            .map(|t| t.1)
            .unwrap_or_else(|| if has_col { freeze.freeze_col } else { 1 });

    let active_pane = if has_row && has_col {
        "bottomRight"
    } else if has_row {
        "bottomLeft"
    } else {
        "topRight"
    };

    out.push_str("<pane");
    if has_col {
        out.push_str(&format!(" xSplit=\"{}\"", x_split));
    }
    if has_row {
        out.push_str(&format!(" ySplit=\"{}\"", y_split));
    }
    let top_left_cell = refs::format_a1(tl_row, tl_col);
    out.push_str(&format!(" topLeftCell=\"{}\"", top_left_cell));
    out.push_str(&format!(" activePane=\"{}\"", active_pane));
    out.push_str(" state=\"frozen\"/>");
}

/// Emit `<pane .../>` for a split (non-frozen) pane.
fn emit_split_pane(out: &mut String, split: &SplitPane) {
    let tl_row = split.top_left.map(|t| t.0).unwrap_or(1);
    let tl_col = split.top_left.map(|t| t.1).unwrap_or(1);

    let has_x = split.x_split != 0.0;
    let has_y = split.y_split != 0.0;

    let active_pane = if has_x && has_y {
        "bottomRight"
    } else if has_y {
        "bottomLeft"
    } else {
        "topRight"
    };

    out.push_str("<pane");
    if has_x {
        out.push_str(&format!(" xSplit=\"{:.2}\"", split.x_split));
    }
    if has_y {
        out.push_str(&format!(" ySplit=\"{:.2}\"", split.y_split));
    }

    let top_left_cell = refs::format_a1(tl_row, tl_col);
    out.push_str(&format!(" topLeftCell=\"{}\"", top_left_cell));
    out.push_str(&format!(" activePane=\"{}\"/>", active_pane));
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn first_sheet_is_tab_selected() {
        let sheet = Worksheet::new("S");
        let mut out = String::new();

        emit(&mut out, &sheet, 0);

        assert!(out.contains("<sheetView tabSelected=\"1\" workbookViewId=\"0\">"));
    }

    #[test]
    fn freeze_b2_emits_one_row_and_column_split() {
        let mut sheet = Worksheet::new("S");
        sheet.set_freeze(2, 2, None);
        let mut out = String::new();

        emit(&mut out, &sheet, 0);

        assert!(out.contains(
            "<pane xSplit=\"1\" ySplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/>"
        ));
    }

    #[test]
    fn split_pane_omits_frozen_state() {
        let mut sheet = Worksheet::new("S");
        sheet.set_split(100.0, 50.0, Some((3, 4)));
        let mut out = String::new();

        emit(&mut out, &sheet, 0);

        assert!(out.contains(
            "<pane xSplit=\"100.00\" ySplit=\"50.00\" topLeftCell=\"D3\" activePane=\"bottomRight\"/>"
        ));
        assert!(!out.contains("state=\"frozen\""));
    }
}
