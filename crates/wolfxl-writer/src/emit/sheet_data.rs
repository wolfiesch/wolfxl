//! `<sheetData>` row and cell emitter for worksheet XML.

use crate::intern::SstBuilder;
use crate::model::cell::{FormulaResult, WriteCell, WriteCellValue};
use crate::model::worksheet::{Row, Worksheet};
use crate::{refs, xml_escape};

/// Emit `<sheetData>...</sheetData>` for worksheet rows and cells.
pub fn emit(out: &mut String, sheet: &Worksheet, sst: &mut SstBuilder) {
    if sheet.rows.is_empty() {
        out.push_str("<sheetData/>");
        return;
    }

    out.push_str("<sheetData>");

    for (&row_num, row) in &sheet.rows {
        emit_row(out, row_num, row, sst);
    }

    out.push_str("</sheetData>");
}

fn emit_row(out: &mut String, row_num: u32, row: &Row, sst: &mut SstBuilder) {
    let has_real_cells = row
        .cells
        .values()
        .any(|c| !matches!(c.value, WriteCellValue::Blank) || c.style_id.is_some());
    let has_attrs = row.custom_height.is_some() || row.hidden || row.style_id.is_some();

    if row.cells.is_empty() && !has_attrs {
        return;
    }

    out.push_str(&format!("<row r=\"{}\"", row_num));

    if let Some(h) = row.custom_height {
        out.push_str(&format!(" ht=\"{}\" customHeight=\"1\"", format_f64(h)));
    }
    if row.hidden {
        out.push_str(" hidden=\"1\"");
    }
    if let Some(s) = row.style_id {
        out.push_str(&format!(" s=\"{}\" customFormat=\"1\"", s));
    }

    if row.cells.is_empty() || !has_real_cells {
        if !has_real_cells {
            out.push_str("/>");
            return;
        }
    }

    out.push('>');

    for (&col_num, cell) in &row.cells {
        emit_cell(out, row_num, col_num, cell, sst);
    }

    out.push_str("</row>");
}

fn emit_cell(out: &mut String, row_num: u32, col_num: u32, cell: &WriteCell, sst: &mut SstBuilder) {
    let cell_ref = refs::format_a1(row_num, col_num);

    match &cell.value {
        WriteCellValue::Blank => {
            if let Some(s) = cell.style_id {
                out.push_str(&format!("<c r=\"{}\" s=\"{}\"/>", cell_ref, s));
            }
        }

        WriteCellValue::Number(n) => {
            out.push_str(&format!("<c r=\"{}\"", cell_ref));
            if let Some(s) = cell.style_id {
                out.push_str(&format!(" s=\"{}\"", s));
            }
            out.push_str(&format!("><v>{}</v></c>", format_number(*n)));
        }

        WriteCellValue::String(s) => {
            let idx = sst.intern(s);
            out.push_str(&format!("<c r=\"{}\" t=\"s\"", cell_ref));
            if let Some(style) = cell.style_id {
                out.push_str(&format!(" s=\"{}\"", style));
            }
            out.push_str(&format!("><v>{}</v></c>", idx));
        }

        WriteCellValue::Boolean(b) => {
            out.push_str(&format!("<c r=\"{}\" t=\"b\"", cell_ref));
            if let Some(s) = cell.style_id {
                out.push_str(&format!(" s=\"{}\"", s));
            }
            let bval = if *b { 1 } else { 0 };
            out.push_str(&format!("><v>{}</v></c>", bval));
        }

        WriteCellValue::Formula { expr, result } => {
            let escaped_expr = xml_escape::text(expr);
            match result {
                None => {
                    out.push_str(&format!("<c r=\"{}\"", cell_ref));
                    if let Some(s) = cell.style_id {
                        out.push_str(&format!(" s=\"{}\"", s));
                    }
                    out.push_str(&format!("><f>{}</f><v>0</v></c>", escaped_expr));
                }
                Some(FormulaResult::Number(n)) => {
                    out.push_str(&format!("<c r=\"{}\"", cell_ref));
                    if let Some(s) = cell.style_id {
                        out.push_str(&format!(" s=\"{}\"", s));
                    }
                    out.push_str(&format!(
                        "><f>{}</f><v>{}</v></c>",
                        escaped_expr,
                        format_number(*n)
                    ));
                }
                Some(FormulaResult::String(s)) => {
                    out.push_str(&format!("<c r=\"{}\" t=\"str\"", cell_ref));
                    if let Some(style) = cell.style_id {
                        out.push_str(&format!(" s=\"{}\"", style));
                    }
                    out.push_str(&format!(
                        "><f>{}</f><v>{}</v></c>",
                        escaped_expr,
                        xml_escape::text(s)
                    ));
                }
                Some(FormulaResult::Boolean(b)) => {
                    out.push_str(&format!("<c r=\"{}\" t=\"b\"", cell_ref));
                    if let Some(s) = cell.style_id {
                        out.push_str(&format!(" s=\"{}\"", s));
                    }
                    let bval = if *b { 1 } else { 0 };
                    out.push_str(&format!("><f>{}</f><v>{}</v></c>", escaped_expr, bval));
                }
            }
        }

        WriteCellValue::DateSerial(f) => {
            out.push_str(&format!("<c r=\"{}\"", cell_ref));
            if let Some(s) = cell.style_id {
                out.push_str(&format!(" s=\"{}\"", s));
            }
            out.push_str(&format!("><v>{}</v></c>", format_number(*f)));
        }

        WriteCellValue::InlineRichText(runs) => {
            out.push_str(&format!("<c r=\"{}\" t=\"inlineStr\"", cell_ref));
            if let Some(s) = cell.style_id {
                out.push_str(&format!(" s=\"{}\"", s));
            }
            out.push_str("><is>");
            out.push_str(&crate::rich_text::emit_runs(runs));
            out.push_str("</is></c>");
        }

        WriteCellValue::ArrayFormula { ref_range, text } => {
            out.push_str(&format!("<c r=\"{}\"", cell_ref));
            if let Some(s) = cell.style_id {
                out.push_str(&format!(" s=\"{}\"", s));
            }
            out.push_str(&format!(
                "><f t=\"array\" ref=\"{}\">{}</f></c>",
                xml_escape::attr(ref_range),
                xml_escape::text(text),
            ));
        }

        WriteCellValue::DataTableFormula {
            ref_range,
            ca,
            dt2_d,
            dtr,
            r1,
            r2,
        } => {
            out.push_str(&format!("<c r=\"{}\"", cell_ref));
            if let Some(s) = cell.style_id {
                out.push_str(&format!(" s=\"{}\"", s));
            }
            out.push_str("><f t=\"dataTable\"");
            out.push_str(&format!(" ref=\"{}\"", xml_escape::attr(ref_range)));
            if *ca {
                out.push_str(" ca=\"1\"");
            }
            if *dt2_d {
                out.push_str(" dt2D=\"1\"");
            }
            if *dtr {
                out.push_str(" dtr=\"1\"");
            }
            if let Some(r1v) = r1 {
                out.push_str(&format!(" r1=\"{}\"", xml_escape::attr(r1v)));
            }
            if let Some(r2v) = r2 {
                out.push_str(&format!(" r2=\"{}\"", xml_escape::attr(r2v)));
            }
            out.push_str("/></c>");
        }

        WriteCellValue::SpillChild => {
            out.push_str(&format!("<c r=\"{}\"", cell_ref));
            if let Some(s) = cell.style_id {
                out.push_str(&format!(" s=\"{}\"", s));
            }
            out.push_str("/>");
        }
    }
}

fn format_number(n: f64) -> String {
    if n == (n as i64) as f64 {
        format!("{}", n as i64)
    } else {
        format!("{}", n)
    }
}

fn format_f64(n: f64) -> String {
    if n == (n as i64) as f64 && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        format!("{}", n)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_sheet_data_self_closes() {
        let sheet = Worksheet::new("S");
        let mut sst = SstBuilder::default();
        let mut out = String::new();

        emit(&mut out, &sheet, &mut sst);

        assert_eq!(out, "<sheetData/>");
    }

    #[test]
    fn string_cells_intern_into_shared_string_table() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(1, 1, WriteCell::new(WriteCellValue::String("hello".into())));
        let mut sst = SstBuilder::default();
        let mut out = String::new();

        emit(&mut out, &sheet, &mut sst);

        assert!(out.contains("<c r=\"A1\" t=\"s\"><v>0</v></c>"));
        assert_eq!(sst.unique_count(), 1);
    }

    #[test]
    fn styled_blank_cells_emit_self_closing_cell() {
        let mut sheet = Worksheet::new("S");
        sheet.set_cell(2, 3, WriteCell::new(WriteCellValue::Blank).with_style(4));
        let mut sst = SstBuilder::default();
        let mut out = String::new();

        emit(&mut out, &sheet, &mut sst);

        assert!(out.contains("<c r=\"C2\" s=\"4\"/>"));
    }
}
