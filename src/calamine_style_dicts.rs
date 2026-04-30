//! Python dict emitters for calamine style fragments.

use calamine_styles::{
    Alignment, Border, BorderStyle as CalBorderStyle, Fill, FillPattern, Font, FontStyle,
    FontWeight, TextRotation,
};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use crate::calamine_format_helpers::{
    border_style_str, color_to_hex, h_align_str, underline_str, v_align_str,
};

#[derive(Clone, Debug)]
pub(crate) struct DiagonalBorderInfo {
    pub(crate) up: bool,
    pub(crate) down: bool,
    pub(crate) style: String,
    pub(crate) color: String,
}

pub(crate) fn populate_font(d: &Bound<'_, PyDict>, font: &Font) -> PyResult<()> {
    if font.weight == FontWeight::Bold {
        d.set_item("bold", true)?;
    }
    if font.style == FontStyle::Italic {
        d.set_item("italic", true)?;
    }
    if let Some(u) = underline_str(&font.underline) {
        d.set_item("underline", u)?;
    }
    if font.strikethrough {
        d.set_item("strikethrough", true)?;
    }
    if let Some(name) = &font.name {
        d.set_item("font_name", name.as_str())?;
    }
    if let Some(size) = font.size {
        d.set_item("font_size", size)?;
    }
    if let Some(color) = &font.color {
        d.set_item("font_color", color_to_hex(color))?;
    }
    Ok(())
}

pub(crate) fn populate_fill(d: &Bound<'_, PyDict>, fill: &Fill) -> PyResult<()> {
    if fill.pattern != FillPattern::None {
        if let Some(color) = fill.get_color() {
            d.set_item("bg_color", color_to_hex(&color))?;
        }
    }
    Ok(())
}

pub(crate) fn populate_alignment(d: &Bound<'_, PyDict>, align: &Alignment) -> PyResult<()> {
    if let Some(h) = h_align_str(&align.horizontal) {
        d.set_item("h_align", h)?;
    }
    if let Some(v) = v_align_str(&align.vertical) {
        d.set_item("v_align", v)?;
    }
    if align.wrap_text {
        d.set_item("wrap", true)?;
    }
    match align.text_rotation {
        TextRotation::None => {}
        TextRotation::Degrees(deg) => {
            if deg != 0 {
                d.set_item("rotation", deg)?;
            }
        }
        TextRotation::Stacked => {
            d.set_item("rotation", 255)?;
        }
    }
    if let Some(indent) = align.indent {
        if indent > 0 {
            d.set_item("indent", indent)?;
        }
    }
    Ok(())
}

pub(crate) fn maybe_set_edge(
    py: Python<'_>,
    d: &Bound<'_, PyDict>,
    key: &str,
    border: &Border,
) -> PyResult<()> {
    if border.style == CalBorderStyle::None {
        return Ok(());
    }
    let edge = PyDict::new(py);
    edge.set_item("style", border_style_str(&border.style))?;
    let color_str = border
        .color
        .as_ref()
        .map(|c| color_to_hex(c))
        .unwrap_or_else(|| "#000000".to_string());
    edge.set_item("color", color_str)?;
    d.set_item(key, edge)?;
    Ok(())
}

pub(crate) fn set_edge_from_style(
    py: Python<'_>,
    d: &Bound<'_, PyDict>,
    key: &str,
    style: &str,
    color: &str,
) -> PyResult<()> {
    if style == "none" {
        return Ok(());
    }
    let edge = PyDict::new(py);
    edge.set_item("style", style)?;
    edge.set_item("color", color)?;
    d.set_item(key, edge)?;
    Ok(())
}
