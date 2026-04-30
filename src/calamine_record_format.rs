//! Compact record-format carriers for the calamine styled read path.

use calamine_styles::{FillPattern, FontStyle, FontWeight, Style, TextRotation};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use crate::calamine_format_helpers::{
    border_style_str, color_to_hex, h_align_str, is_default_font_size, underline_str, v_align_str,
};

#[derive(Clone, Debug, Default)]
pub(crate) struct RawFontInfo {
    pub(crate) bold: bool,
    pub(crate) italic: bool,
    pub(crate) underline: Option<String>,
    pub(crate) strikethrough: bool,
    pub(crate) name: Option<String>,
    pub(crate) size: Option<f64>,
    pub(crate) color: Option<String>,
}

#[derive(Clone, Debug, Default)]
pub(crate) struct RecordFormatInfo {
    bold: bool,
    italic: bool,
    underline: Option<String>,
    strikethrough: bool,
    font_size: Option<f64>,
    bg_color: Option<String>,
    number_format: Option<String>,
    h_align: Option<&'static str>,
    v_align: Option<&'static str>,
    wrap: bool,
    rotation: Option<i64>,
    indent: Option<u32>,
    bottom_border_style: Option<&'static str>,
    is_double_underline: bool,
}

impl RecordFormatInfo {
    pub(crate) fn from_style(style: &Style) -> Self {
        let mut info = Self::default();

        if let Some(font) = &style.font {
            if font.weight == FontWeight::Bold {
                info.bold = true;
            }
            if font.style == FontStyle::Italic {
                info.italic = true;
            }
            if let Some(u) = underline_str(&font.underline) {
                info.underline = Some(u.to_string());
            }
            if font.strikethrough {
                info.strikethrough = true;
            }
            if let Some(size) = font.size {
                if !is_default_font_size(size) {
                    info.font_size = Some(size);
                }
            }
        }

        if let Some(fill) = &style.fill {
            if fill.pattern != FillPattern::None {
                if let Some(color) = fill.get_color() {
                    info.bg_color = Some(color_to_hex(&color));
                }
            }
        }

        if let Some(align) = &style.alignment {
            info.h_align = h_align_str(&align.horizontal);
            info.v_align = v_align_str(&align.vertical);
            if align.wrap_text {
                info.wrap = true;
            }
            match align.text_rotation {
                TextRotation::None => {}
                TextRotation::Degrees(deg) => {
                    if deg != 0 {
                        info.rotation = Some(deg as i64);
                    }
                }
                TextRotation::Stacked => {
                    info.rotation = Some(255);
                }
            }
            if let Some(indent) = align.indent {
                if indent > 0 {
                    info.indent = Some(indent as u32);
                }
            }
        }

        if let Some(borders) = &style.borders {
            let bottom_style = border_style_str(&borders.bottom.style);
            if bottom_style != "none" {
                info.bottom_border_style = Some(bottom_style);
                info.is_double_underline = bottom_style == "double";
            }
        }

        info
    }

    pub(crate) fn overlay_raw_font(&mut self, font: &RawFontInfo) {
        if font.bold {
            self.bold = true;
        }
        if font.italic {
            self.italic = true;
        }
        if let Some(underline) = &font.underline {
            self.underline = Some(underline.clone());
        }
        if font.strikethrough {
            self.strikethrough = true;
        }
        if let Some(size) = font.size {
            if !is_default_font_size(size) {
                self.font_size = Some(size);
            }
        }
    }

    pub(crate) fn populate_dict(&self, d: &Bound<'_, PyDict>) -> PyResult<()> {
        if self.bold {
            d.set_item("bold", true)?;
        }
        if self.italic {
            d.set_item("italic", true)?;
        }
        if let Some(underline) = &self.underline {
            d.set_item("underline", underline)?;
        }
        if self.strikethrough {
            d.set_item("strikethrough", true)?;
        }
        if let Some(size) = self.font_size {
            d.set_item("font_size", size)?;
        }
        if let Some(color) = &self.bg_color {
            d.set_item("bg_color", color)?;
        }
        if let Some(number_format) = &self.number_format {
            d.set_item("number_format", number_format)?;
        }
        if let Some(h) = self.h_align {
            d.set_item("h_align", h)?;
        }
        if let Some(v) = self.v_align {
            d.set_item("v_align", v)?;
        }
        if self.wrap {
            d.set_item("wrap", true)?;
        }
        if let Some(rotation) = self.rotation {
            d.set_item("rotation", rotation)?;
        }
        if let Some(indent) = self.indent {
            d.set_item("indent", indent)?;
        }
        if let Some(bottom_style) = self.bottom_border_style {
            d.set_item("bottom_border_style", bottom_style)?;
            d.set_item("has_bottom_border", true)?;
            if self.is_double_underline {
                d.set_item("is_double_underline", true)?;
            }
        }
        Ok(())
    }

    pub(crate) fn set_number_format(&mut self, number_format: String) {
        self.number_format = Some(number_format);
    }
}
