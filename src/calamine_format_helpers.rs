//! Formatting helpers shared by the styled calamine read backend.

use calamine_styles::{
    BorderStyle as CalBorderStyle, Color, HorizontalAlignment, UnderlineStyle, VerticalAlignment,
};

const CALIBRI_WIDTH_PADDING: f64 = 0.83203125;
const ALT_WIDTH_PADDING: f64 = 0.7109375;
const WIDTH_TOLERANCE: f64 = 0.0005;

pub(crate) fn openpyxl_builtin_num_fmt(format_id: u32) -> Option<&'static str> {
    match format_id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        5 => Some("\"$\"#,##0_);(\"$\"#,##0)"),
        6 => Some("\"$\"#,##0_);[Red](\"$\"#,##0)"),
        7 => Some("\"$\"#,##0.00_);(\"$\"#,##0.00)"),
        8 => Some("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("mm-dd-yy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        37 => Some("#,##0_);(#,##0)"),
        38 => Some("#,##0_);[Red](#,##0)"),
        39 => Some("#,##0.00_);(#,##0.00)"),
        40 => Some("#,##0.00_);[Red](#,##0.00)"),
        41 => Some(r#"_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)"#),
        42 => Some(r#"_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)"#),
        43 => Some(r#"_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)"#),
        44 => Some(r#"_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)"#),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mmss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        _ => None,
    }
}

/// Convert a calamine Color to a "#RRGGBB" hex string.
pub(crate) fn color_to_hex(c: &Color) -> String {
    format!("#{:02X}{:02X}{:02X}", c.red, c.green, c.blue)
}

/// Convert a calamine BorderStyle to the ExcelBench string token.
pub(crate) fn border_style_str(s: &CalBorderStyle) -> &'static str {
    match s {
        CalBorderStyle::None => "none",
        CalBorderStyle::Thin => "thin",
        CalBorderStyle::Medium => "medium",
        CalBorderStyle::Thick => "thick",
        CalBorderStyle::Double => "double",
        CalBorderStyle::Hair => "hair",
        CalBorderStyle::Dashed => "dashed",
        CalBorderStyle::Dotted => "dotted",
        CalBorderStyle::MediumDashed => "mediumDashed",
        CalBorderStyle::DashDot => "dashDot",
        CalBorderStyle::DashDotDot => "dashDotDot",
        CalBorderStyle::SlantDashDot => "slantDashDot",
    }
}

/// Convert a calamine HorizontalAlignment to the ExcelBench string.
pub(crate) fn h_align_str(a: &HorizontalAlignment) -> Option<&'static str> {
    match a {
        HorizontalAlignment::General => None,
        HorizontalAlignment::Left => Some("left"),
        HorizontalAlignment::Center => Some("center"),
        HorizontalAlignment::Right => Some("right"),
        HorizontalAlignment::Justify => Some("justify"),
        HorizontalAlignment::Distributed => Some("distributed"),
        HorizontalAlignment::Fill => Some("fill"),
    }
}

/// Convert a calamine VerticalAlignment to the ExcelBench string.
pub(crate) fn v_align_str(a: &VerticalAlignment) -> Option<&'static str> {
    match a {
        VerticalAlignment::Bottom => Some("bottom"),
        VerticalAlignment::Top => Some("top"),
        VerticalAlignment::Center => Some("center"),
        VerticalAlignment::Justify => Some("justify"),
        VerticalAlignment::Distributed => Some("distributed"),
    }
}

/// Convert a calamine UnderlineStyle to the ExcelBench string.
pub(crate) fn underline_str(u: &UnderlineStyle) -> Option<&'static str> {
    match u {
        UnderlineStyle::None => None,
        UnderlineStyle::Single => Some("single"),
        UnderlineStyle::Double => Some("double"),
        UnderlineStyle::SingleAccounting => Some("singleAccounting"),
        UnderlineStyle::DoubleAccounting => Some("doubleAccounting"),
    }
}

pub(crate) fn is_default_font_size(size: f64) -> bool {
    (size - 11.0).abs() < f64::EPSILON
}

// Excel stores column widths with font-metric padding included.
// These paddings match the Python-side adjustment previously used by
// `RustCalamineStyledAdapter.read_column_width()`.
pub(crate) fn strip_excel_padding(raw: f64) -> f64 {
    let frac = raw % 1.0;
    for padding in [CALIBRI_WIDTH_PADDING, ALT_WIDTH_PADDING] {
        if (frac - padding).abs() < WIDTH_TOLERANCE {
            let adjusted = raw - padding;
            if adjusted >= 0.0 {
                return (adjusted * 10000.0).round() / 10000.0;
            }
        }
    }
    (raw * 10000.0).round() / 10000.0
}
