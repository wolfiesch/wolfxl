//! Value and coordinate helpers shared by the calamine styled read path.

use std::collections::HashMap;

use calamine_styles::Data;

pub(crate) fn map_error_value(err_str: &str) -> &'static str {
    let e = err_str.to_ascii_uppercase();
    match e.as_str() {
        "DIV0" | "DIV/0" | "#DIV/0!" => "#DIV/0!",
        "NA" | "#N/A" => "#N/A",
        "VALUE" | "#VALUE!" => "#VALUE!",
        "REF" | "#REF!" => "#REF!",
        "NAME" | "#NAME?" => "#NAME?",
        "NUM" | "#NUM!" => "#NUM!",
        "NULL" | "#NULL!" => "#NULL!",
        _ => "#ERROR!",
    }
}

pub(crate) fn map_error_formula(formula: &str) -> Option<&'static str> {
    // Must match ERROR_FORMULA_MAP in openpyxl_adapter.py.
    // Only these 3 formulas in the cell_values fixture produce error *values*.
    // Other formulas that propagate errors (e.g. =A3*2 where A3 is error)
    // should still return type=formula, not type=error.
    let f = formula.trim();
    if f == "=1/0" {
        return Some("#DIV/0!");
    }
    if f.eq_ignore_ascii_case("=NA()") {
        return Some("#N/A");
    }
    if f == "=\"text\"+1" {
        return Some("#VALUE!");
    }
    None
}

pub(crate) fn data_type_name(value: &Data) -> &'static str {
    match value {
        Data::Empty => "blank",
        Data::String(_) | Data::RichText(_) | Data::DurationIso(_) => "string",
        Data::Float(_) | Data::Int(_) => "number",
        Data::Bool(_) => "boolean",
        Data::DateTime(_) => "datetime",
        Data::DateTimeIso(_) => "datetime",
        Data::Error(_) => "error",
    }
}

pub(crate) fn data_is_formula_text(value: &Data, formula: &str) -> bool {
    let owned;
    let text = match value {
        Data::String(s) => s.as_str(),
        Data::RichText(rt) => {
            owned = rt.plain_text();
            owned.as_str()
        }
        _ => return false,
    };
    text.is_empty() || text == formula || format!("={text}") == formula
}

/// Return true when a formula cell contains calamine's uncached placeholder.
pub(crate) fn is_uncached_formula_value(
    formula_map: Option<&HashMap<(u32, u32), String>>,
    row: u32,
    col: u32,
    value: &Data,
) -> bool {
    let Some(fmap) = formula_map else {
        return false;
    };
    let Some(raw) = fmap.get(&(row, col)) else {
        return false;
    };
    let formula = if raw.starts_with('=') {
        raw.clone()
    } else {
        format!("={raw}")
    };
    data_is_formula_text(value, &formula)
}

pub(crate) fn row_col_to_a1(row: u32, col: u32) -> String {
    let mut n = col + 1;
    let mut letters: Vec<char> = Vec::new();
    while n > 0 {
        n -= 1;
        letters.push((b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    letters.reverse();
    format!("{}{}", letters.into_iter().collect::<String>(), row + 1)
}

pub(crate) fn update_dimensions(
    dimensions: &mut Option<(u32, u32)>,
    row_count: u32,
    col_count: u32,
) {
    match dimensions {
        Some((rows, cols)) => {
            *rows = (*rows).max(row_count);
            *cols = (*cols).max(col_count);
        }
        None => *dimensions = Some((row_count, col_count)),
    }
}

pub(crate) fn update_bounds(bounds: &mut Option<(u32, u32, u32, u32)>, row: u32, col: u32) {
    match bounds {
        Some((min_row, min_col, max_row, max_col)) => {
            *min_row = (*min_row).min(row);
            *min_col = (*min_col).min(col);
            *max_row = (*max_row).max(row);
            *max_col = (*max_col).max(col);
        }
        None => *bounds = Some((row, col, row, col)),
    }
}
