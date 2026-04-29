//! Format and border payload parsing for the native writer backend.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyDict};
use wolfxl_writer::model::{
    AlignmentSpec, BorderSideSpec, BorderSpec, FillSpec, FontSpec, FormatSpec, Worksheet,
    WriteCell, WriteCellValue,
};
use wolfxl_writer::Workbook;

/// Normalize a Python-side color string to OOXML's 8-char ARGB form.
pub(crate) fn parse_hex_color(input: &str) -> Option<String> {
    let s = input.strip_prefix('#').unwrap_or(input);
    let upper: String = s.chars().map(|c| c.to_ascii_uppercase()).collect();

    if !upper.chars().all(|c| c.is_ascii_hexdigit()) {
        return None;
    }

    match upper.len() {
        3 => {
            let mut expanded = String::with_capacity(8);
            expanded.push_str("FF");
            for ch in upper.chars() {
                expanded.push(ch);
                expanded.push(ch);
            }
            Some(expanded)
        }
        6 => Some(format!("FF{upper}")),
        8 => Some(upper),
        _ => None,
    }
}

fn dict_to_format_spec(dict: &Bound<'_, PyDict>) -> PyResult<FormatSpec> {
    let mut spec = FormatSpec::default();

    let mut font = FontSpec::default();
    let mut font_touched = false;

    if let Some(v) = dict.get_item("bold")? {
        if let Ok(b) = v.extract::<bool>() {
            font.bold = b;
            font_touched |= b;
        }
    }
    if let Some(v) = dict.get_item("italic")? {
        if let Ok(b) = v.extract::<bool>() {
            font.italic = b;
            font_touched |= b;
        }
    }
    if let Some(v) = dict.get_item("underline")? {
        if let Ok(s) = v.extract::<String>() {
            if !s.is_empty() {
                font.underline = Some(s);
                font_touched = true;
            }
        } else if let Ok(b) = v.extract::<bool>() {
            font.underline = if b { Some("single".to_string()) } else { None };
            font_touched |= b;
        }
    }
    if let Some(v) = dict.get_item("strikethrough")? {
        if let Ok(b) = v.extract::<bool>() {
            font.strikethrough = b;
            font_touched |= b;
        }
    }
    if let Some(v) = dict.get_item("font_name")? {
        if let Ok(s) = v.extract::<String>() {
            if !s.is_empty() {
                font.name = Some(s);
                font_touched = true;
            }
        }
    }
    if let Some(v) = dict.get_item("font_size")? {
        if let Ok(f) = v.extract::<f64>() {
            if f.is_finite() && f >= 0.0 {
                font.size = Some(f.round() as u32);
                font_touched = true;
            }
        }
    }
    if let Some(v) = dict.get_item("font_color")? {
        if let Ok(s) = v.extract::<String>() {
            if let Some(rgb) = parse_hex_color(&s) {
                font.color_rgb = Some(rgb);
                font_touched = true;
            }
        }
    }
    if font_touched {
        spec.font = Some(font);
    }

    if let Some(v) = dict.get_item("bg_color")? {
        if let Ok(s) = v.extract::<String>() {
            if let Some(rgb) = parse_hex_color(&s) {
                spec.fill = Some(FillSpec {
                    pattern_type: "solid".to_string(),
                    fg_color_rgb: Some(rgb.clone()),
                    bg_color_rgb: Some(rgb),
                });
            }
        }
    }

    if let Some(v) = dict.get_item("number_format")? {
        if let Ok(s) = v.extract::<String>() {
            if !s.is_empty() {
                spec.number_format = Some(s);
            }
        }
    }

    let mut align = AlignmentSpec::default();
    let mut align_touched = false;
    if let Some(v) = dict.get_item("h_align")? {
        if let Ok(s) = v.extract::<String>() {
            if !s.is_empty() {
                align.horizontal = Some(s);
                align_touched = true;
            }
        }
    }
    if let Some(v) = dict.get_item("v_align")? {
        if let Ok(s) = v.extract::<String>() {
            if !s.is_empty() {
                align.vertical = Some(s);
                align_touched = true;
            }
        }
    }
    if let Some(v) = dict.get_item("wrap")? {
        if let Ok(b) = v.extract::<bool>() {
            align.wrap_text = b;
            align_touched |= b;
        }
    }
    if let Some(v) = dict.get_item("rotation")? {
        if let Ok(i) = v.extract::<i32>() {
            if i >= 0 {
                align.text_rotation = i as u32;
                align_touched = true;
            }
        }
    }
    if let Some(v) = dict.get_item("indent")? {
        if let Ok(i) = v.extract::<i32>() {
            if i >= 0 {
                align.indent = i as u32;
                align_touched = true;
            }
        }
    }
    if align_touched {
        spec.alignment = Some(align);
    }

    Ok(spec)
}

fn edge_to_side_spec(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<(BorderSideSpec, bool)> {
    let mut side = BorderSideSpec::default();
    let mut touched = false;

    let Some(sub) = dict.get_item(key)? else {
        return Ok((side, false));
    };
    let Ok(d) = sub.cast::<PyDict>() else {
        return Ok((side, false));
    };

    if let Some(v) = d.get_item("style")? {
        if let Ok(s) = v.extract::<String>() {
            if !s.is_empty() {
                side.style = Some(s);
                touched = true;
            }
        }
    }
    if let Some(v) = d.get_item("color")? {
        if let Ok(s) = v.extract::<String>() {
            if let Some(rgb) = parse_hex_color(&s) {
                side.color_rgb = Some(rgb);
                touched = true;
            }
        }
    }
    Ok((side, touched))
}

fn dict_to_border_spec(dict: &Bound<'_, PyDict>) -> PyResult<BorderSpec> {
    let mut border = BorderSpec::default();

    let (top, t1) = edge_to_side_spec(dict, "top")?;
    let (bottom, t2) = edge_to_side_spec(dict, "bottom")?;
    let (left, t3) = edge_to_side_spec(dict, "left")?;
    let (right, t4) = edge_to_side_spec(dict, "right")?;
    let (diag_up, t5) = edge_to_side_spec(dict, "diagonal_up")?;
    let (diag_down, t6) = edge_to_side_spec(dict, "diagonal_down")?;

    if t1 {
        border.top = top;
    }
    if t2 {
        border.bottom = bottom;
    }
    if t3 {
        border.left = left;
    }
    if t4 {
        border.right = right;
    }
    if t5 || t6 {
        if t6 {
            border.diagonal = diag_down;
        } else {
            border.diagonal = diag_up;
        }
        border.diagonal_up = t5;
        border.diagonal_down = t6;
    }

    Ok(border)
}

pub(crate) fn intern_format_from_dict(
    wb: &mut Workbook,
    dict: &Bound<'_, PyDict>,
) -> PyResult<u32> {
    let spec = dict_to_format_spec(dict)?;
    Ok(wb.styles.intern_format(&spec))
}

pub(crate) fn intern_border_only(wb: &mut Workbook, dict: &Bound<'_, PyDict>) -> PyResult<u32> {
    let border = dict_to_border_spec(dict)?;
    let spec = FormatSpec {
        border: Some(border),
        ..Default::default()
    };
    Ok(wb.styles.intern_format(&spec))
}

pub(crate) fn apply_cell_format(
    wb: &mut Workbook,
    sheet: &str,
    row: u32,
    col: u32,
    dict: &Bound<'_, PyDict>,
) -> PyResult<()> {
    let style_id = intern_format_from_dict(wb, dict)?;
    let ws = require_sheet(wb, sheet)?;
    set_cell_style_id(ws, row, col, style_id);
    Ok(())
}

pub(crate) fn apply_cell_border(
    wb: &mut Workbook,
    sheet: &str,
    row: u32,
    col: u32,
    dict: &Bound<'_, PyDict>,
) -> PyResult<()> {
    let style_id = intern_border_only(wb, dict)?;
    let ws = require_sheet(wb, sheet)?;
    set_cell_style_id(ws, row, col, style_id);
    Ok(())
}

pub(crate) fn apply_format_grid(
    wb: &mut Workbook,
    sheet: &str,
    base_row: u32,
    base_col: u32,
    grid: &Bound<'_, PyAny>,
) -> PyResult<()> {
    let to_apply = collect_style_grid(wb, base_row, base_col, grid, intern_format_from_dict)?;
    let ws = require_sheet(wb, sheet)?;
    for (row, col, style_id) in to_apply {
        set_cell_style_id(ws, row, col, style_id);
    }
    Ok(())
}

pub(crate) fn apply_border_grid(
    wb: &mut Workbook,
    sheet: &str,
    base_row: u32,
    base_col: u32,
    grid: &Bound<'_, PyAny>,
) -> PyResult<()> {
    let to_apply = collect_style_grid(wb, base_row, base_col, grid, intern_border_only)?;
    let ws = require_sheet(wb, sheet)?;
    for (row, col, style_id) in to_apply {
        set_cell_style_id(ws, row, col, style_id);
    }
    Ok(())
}

fn collect_style_grid(
    wb: &mut Workbook,
    base_row: u32,
    base_col: u32,
    grid: &Bound<'_, PyAny>,
    intern: fn(&mut Workbook, &Bound<'_, PyDict>) -> PyResult<u32>,
) -> PyResult<Vec<(u32, u32, u32)>> {
    let rows: Vec<Bound<'_, PyAny>> = grid.extract()?;
    let mut to_apply: Vec<(u32, u32, u32)> = Vec::new();

    for (ri, row_obj) in rows.iter().enumerate() {
        let cols: Vec<Bound<'_, PyAny>> = row_obj.extract()?;
        for (ci, val) in cols.iter().enumerate() {
            if val.is_none() {
                continue;
            }
            let dict = val
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("style grid element must be dict or None"))?;
            if dict.is_empty() {
                continue;
            }
            let row = base_row + ri as u32;
            let col = base_col + ci as u32;
            let style_id = intern(wb, dict)?;
            to_apply.push((row, col, style_id));
        }
    }

    Ok(to_apply)
}

fn require_sheet<'wb>(wb: &'wb mut Workbook, name: &str) -> PyResult<&'wb mut Worksheet> {
    wb.sheet_mut_by_name(name)
        .ok_or_else(|| PyValueError::new_err(format!("Unknown sheet: {name}")))
}

fn set_cell_style_id(ws: &mut Worksheet, row: u32, col: u32, style_id: u32) {
    let cell = ws
        .rows
        .entry(row)
        .or_default()
        .cells
        .entry(col)
        .or_insert_with(|| WriteCell {
            value: WriteCellValue::Blank,
            style_id: None,
        });
    cell.style_id = Some(style_id);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_hex_color_normalizes_rgb_and_argb() {
        assert_eq!(parse_hex_color("#abc").as_deref(), Some("FFAABBCC"));
        assert_eq!(parse_hex_color("336699").as_deref(), Some("FF336699"));
        assert_eq!(parse_hex_color("#80336699").as_deref(), Some("80336699"));
    }

    #[test]
    fn parse_hex_color_rejects_invalid_shapes() {
        assert_eq!(parse_hex_color(""), None);
        assert_eq!(parse_hex_color("12345"), None);
        assert_eq!(parse_hex_color("zzzzzz"), None);
    }
}
