//! Sheet-feature payload parsing for the native writer backend.

use pyo3::prelude::*;
use pyo3::types::PyDict;
use wolfxl_writer::model::{
    CellIsOperator, ColorScaleStop, Comment, CommentAuthorTable, ConditionalFormat,
    ConditionalKind, ConditionalRule, ConditionalThreshold, DataValidation, DxfRecord, ErrorStyle,
    FillSpec, Hyperlink, Person, StylesBuilder, Table, TableColumn, TableStyle, ThreadedComment,
    ValidationOperator, ValidationType,
};

use crate::native_writer_formats::parse_hex_color;

/// Unwrap an optional wrapper key, or return the original dict unchanged.
pub(crate) fn unwrap_optional_wrapper<'py>(
    dict: &'py Bound<'py, PyDict>,
    wrapper_key: &str,
) -> PyResult<Bound<'py, PyDict>> {
    if let Some(v) = dict.get_item(wrapper_key)? {
        if let Ok(inner) = v.cast::<PyDict>() {
            return Ok(inner.clone());
        }
    }
    Ok(dict.clone())
}

/// Build a `(a1_ref, Hyperlink)` pair from a cfg dict, or `None` for no-op.
pub(crate) fn dict_to_hyperlink(cfg: &Bound<'_, PyDict>) -> PyResult<Option<(String, Hyperlink)>> {
    let cell: Option<String> = cfg
        .get_item("cell")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let target: Option<String> = cfg
        .get_item("target")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });

    let (Some(cell), Some(raw_target)) = (cell, target) else {
        return Ok(None);
    };

    let display: Option<String> = cfg
        .get_item("display")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let tooltip: Option<String> = cfg
        .get_item("tooltip")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let is_internal: bool = cfg
        .get_item("internal")?
        .and_then(|v| v.extract::<bool>().ok())
        .unwrap_or(false);

    let target = if is_internal {
        raw_target.trim_start_matches('#').to_string()
    } else {
        raw_target
    };

    Ok(Some((
        cell,
        Hyperlink {
            target,
            is_internal,
            display,
            tooltip,
        },
    )))
}

/// Build a `(a1_ref, Comment)` pair from a cfg dict, or `None` for no-op.
pub(crate) fn dict_to_comment(
    cfg: &Bound<'_, PyDict>,
    authors: &mut CommentAuthorTable,
) -> PyResult<Option<(String, Comment)>> {
    let cell: Option<String> = cfg
        .get_item("cell")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let text: Option<String> = cfg
        .get_item("text")?
        .and_then(|v| v.extract::<String>().ok());

    let (Some(cell), Some(text)) = (cell, text) else {
        return Ok(None);
    };

    let author_name: String = cfg
        .get_item("author")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) })
        .unwrap_or_default();

    let author_id = authors.intern(author_name);

    Ok(Some((
        cell,
        Comment {
            text,
            author_id,
            width_pt: None,
            height_pt: None,
            visible: false,
        },
    )))
}

/// Build a `Person` from a cfg dict, or `None` for no-op.
///
/// RFC-068 / G08. The Python `PersonRegistry` allocates GUIDs eagerly so this
/// always sees an `id`; the dict-based contract keeps the Rust side
/// agnostic to Python's `Person` class shape.
pub(crate) fn dict_to_person(cfg: &Bound<'_, PyDict>) -> PyResult<Option<Person>> {
    let id: Option<String> = cfg
        .get_item("id")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let display_name: Option<String> = cfg
        .get_item("name")?
        .and_then(|v| v.extract::<String>().ok());

    let (Some(id), Some(display_name)) = (id, display_name) else {
        return Ok(None);
    };

    let user_id: String = cfg
        .get_item("user_id")?
        .and_then(|v| v.extract::<String>().ok())
        .unwrap_or_default();
    let provider_id: String = cfg
        .get_item("provider_id")?
        .and_then(|v| v.extract::<String>().ok())
        .filter(|s| !s.is_empty())
        .unwrap_or_else(|| "None".to_string());

    Ok(Some(Person {
        display_name,
        id,
        user_id,
        provider_id,
    }))
}

/// Build a `ThreadedComment` from a cfg dict, or `None` for no-op.
///
/// RFC-068 / G08. Required keys: `id`, `cell`, `person_id`, `created`, `text`.
/// Optional: `parent_id`, `done`. The Python flush layer is responsible for
/// resolving Python `Person`/`ThreadedComment` references to their GUID
/// strings before calling.
pub(crate) fn dict_to_threaded_comment(
    cfg: &Bound<'_, PyDict>,
) -> PyResult<Option<ThreadedComment>> {
    let id: Option<String> = cfg
        .get_item("id")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let cell_ref: Option<String> = cfg
        .get_item("cell")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let person_id: Option<String> = cfg
        .get_item("person_id")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let created: Option<String> = cfg
        .get_item("created")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let text: Option<String> = cfg.get_item("text")?.and_then(|v| v.extract::<String>().ok());

    let (Some(id), Some(cell_ref), Some(person_id), Some(created), Some(text)) =
        (id, cell_ref, person_id, created, text)
    else {
        return Ok(None);
    };

    let parent_id: Option<String> = cfg
        .get_item("parent_id")?
        .and_then(|v| v.extract::<String>().ok())
        .filter(|s| !s.is_empty());
    let done: bool = cfg
        .get_item("done")?
        .and_then(|v| v.extract::<bool>().ok())
        .unwrap_or(false);

    Ok(Some(ThreadedComment {
        id,
        cell_ref,
        person_id,
        created,
        parent_id,
        text,
        done,
    }))
}

/// Build a `ConditionalFormat` from a cfg dict, or `None` for no-op.
pub(crate) fn dict_to_conditional_format(
    cfg: &Bound<'_, PyDict>,
    styles: &mut StylesBuilder,
) -> PyResult<Option<ConditionalFormat>> {
    let range: Option<String> = cfg.get_item("range")?.and_then(|v| v.extract().ok());
    let rule_type: Option<String> = cfg.get_item("rule_type")?.and_then(|v| v.extract().ok());

    let (Some(range), Some(rule_type)) = (range, rule_type) else {
        return Ok(None);
    };

    let operator: Option<String> = cfg.get_item("operator")?.and_then(|v| v.extract().ok());
    let formula: Option<String> = cfg.get_item("formula")?.and_then(|v| v.extract().ok());
    let stop_if_true: bool = cfg
        .get_item("stop_if_true")?
        .and_then(|v| v.extract::<bool>().ok())
        .unwrap_or(false);

    let mut bg_color: Option<String> = None;
    if let Some(v) = cfg.get_item("format")? {
        if let Ok(fd) = v.cast::<PyDict>() {
            bg_color = fd.get_item("bg_color")?.and_then(|x| x.extract().ok());
        }
    }
    let dxf_id: Option<u32> = if let Some(ref hex) = bg_color {
        parse_hex_color(hex).map(|rgb| {
            let dxf = DxfRecord {
                font: None,
                fill: Some(FillSpec {
                    pattern_type: "solid".to_string(),
                    fg_color_rgb: Some(rgb.clone()),
                    bg_color_rgb: Some(rgb),
                    gradient: None,
                }),
                border: None,
            };
            styles.intern_dxf(&dxf)
        })
    } else {
        None
    };

    let kind = match rule_type.as_str() {
        "cellIs" | "cell_is" => {
            let op_str = operator.as_deref().unwrap_or("equal");
            let op = match op_str {
                "equal" | "==" => CellIsOperator::Equal,
                "notEqual" | "!=" => CellIsOperator::NotEqual,
                "greaterThan" | ">" => CellIsOperator::GreaterThan,
                "greaterThanOrEqual" | ">=" => CellIsOperator::GreaterThanOrEqual,
                "lessThan" | "<" => CellIsOperator::LessThan,
                "lessThanOrEqual" | "<=" => CellIsOperator::LessThanOrEqual,
                "between" => CellIsOperator::Between,
                "notBetween" => CellIsOperator::NotBetween,
                _ => CellIsOperator::Equal,
            };

            let fstr = formula.as_deref().unwrap_or("").trim_start_matches('=');
            let (formula_a, formula_b) =
                if matches!(op, CellIsOperator::Between | CellIsOperator::NotBetween) {
                    if let Some(idx) = fstr.find(',') {
                        (
                            fstr[..idx].trim().to_string(),
                            Some(fstr[idx + 1..].trim().to_string()),
                        )
                    } else {
                        (fstr.to_string(), None)
                    }
                } else {
                    (fstr.to_string(), None)
                };

            ConditionalKind::CellIs {
                operator: op,
                formula_a,
                formula_b,
            }
        }
        "expression" | "formula" => {
            let fstr = formula
                .as_deref()
                .unwrap_or("")
                .trim_start_matches('=')
                .to_string();
            ConditionalKind::Expression { formula: fstr }
        }
        "dataBar" | "data_bar" => {
            let color = cfg
                .get_item("color")?
                .and_then(|v| v.extract::<String>().ok())
                .and_then(|s| parse_hex_color(&s))
                .unwrap_or_else(|| "FF638EC6".to_string());
            ConditionalKind::DataBar {
                color_rgb: color,
                min: ConditionalThreshold::Min,
                max: ConditionalThreshold::Max,
            }
        }
        "colorScale" | "color_scale" => ConditionalKind::ColorScale {
            stops: vec![
                ColorScaleStop {
                    threshold: ConditionalThreshold::Min,
                    color_rgb: "FFAA0000".to_string(),
                },
                ColorScaleStop {
                    threshold: ConditionalThreshold::Percentile(50.0),
                    color_rgb: "FFFFFF00".to_string(),
                },
                ColorScaleStop {
                    threshold: ConditionalThreshold::Max,
                    color_rgb: "FF00AA00".to_string(),
                },
            ],
        },
        "iconSet" | "icon_set" => {
            // G11: build IconSet from the Python IconSetRule blob
            // (icon_style, value_type, values, show_value).
            let set_name: String = cfg
                .get_item("icon_style")?
                .and_then(|v| v.extract::<String>().ok())
                .filter(|s| !s.is_empty())
                .unwrap_or_else(|| "3TrafficLights1".to_string());
            let value_type: String = cfg
                .get_item("value_type")?
                .and_then(|v| v.extract::<String>().ok())
                .filter(|s| !s.is_empty())
                .unwrap_or_else(|| "percent".to_string());
            // Values may be numeric (f64/i64) or strings (for "formula"
            // value-type). Try f64 first, fall back to a string vec.
            let raw_values: Vec<String> = if let Some(v) = cfg.get_item("values")? {
                if let Ok(nums) = v.extract::<Vec<f64>>() {
                    nums.into_iter().map(|n| {
                        // Format whole numbers without trailing decimals so
                        // OOXML stays tidy.
                        if n == (n as i64) as f64 && n.abs() < 1e15 {
                            format!("{}", n as i64)
                        } else {
                            format!("{}", n)
                        }
                    }).collect()
                } else if let Ok(strs) = v.extract::<Vec<String>>() {
                    strs
                } else {
                    Vec::new()
                }
            } else {
                Vec::new()
            };
            let thresholds: Vec<ConditionalThreshold> = raw_values
                .into_iter()
                .map(|val| match value_type.as_str() {
                    "percent" => val.parse::<f64>()
                        .map(ConditionalThreshold::Percent)
                        .unwrap_or(ConditionalThreshold::Percent(0.0)),
                    "percentile" => val.parse::<f64>()
                        .map(ConditionalThreshold::Percentile)
                        .unwrap_or(ConditionalThreshold::Percentile(0.0)),
                    "num" | "number" => val.parse::<f64>()
                        .map(ConditionalThreshold::Number)
                        .unwrap_or(ConditionalThreshold::Number(0.0)),
                    "formula" => ConditionalThreshold::Formula(val),
                    _ => val.parse::<f64>()
                        .map(ConditionalThreshold::Percent)
                        .unwrap_or(ConditionalThreshold::Percent(0.0)),
                })
                .collect();
            let show_value: bool = cfg
                .get_item("show_value")?
                .and_then(|v| v.extract::<bool>().ok())
                .unwrap_or(true);
            ConditionalKind::IconSet {
                set_name,
                thresholds,
                show_value,
            }
        }
        _ => ConditionalKind::Expression {
            formula: "FALSE()".to_string(),
        },
    };

    let rule = ConditionalRule {
        kind,
        dxf_id,
        stop_if_true,
    };

    Ok(Some(ConditionalFormat {
        sqref: range,
        rules: vec![rule],
    }))
}

/// Build a `DataValidation` from a cfg dict, or `None` for no-op.
pub(crate) fn dict_to_data_validation(cfg: &Bound<'_, PyDict>) -> PyResult<Option<DataValidation>> {
    let range: Option<String> = cfg.get_item("range")?.and_then(|v| v.extract().ok());
    let validation_type: Option<String> = cfg
        .get_item("validation_type")?
        .and_then(|v| v.extract().ok());

    let (Some(range), Some(vtype_str)) = (range, validation_type) else {
        return Ok(None);
    };

    let validation_type = match vtype_str.as_str() {
        "whole" | "Whole" => ValidationType::Whole,
        "decimal" | "Decimal" => ValidationType::Decimal,
        "list" | "List" => ValidationType::List,
        "date" | "Date" => ValidationType::Date,
        "time" | "Time" => ValidationType::Time,
        "textLength" | "TextLength" | "text_length" => ValidationType::TextLength,
        "custom" | "Custom" => ValidationType::Custom,
        _ => ValidationType::Any,
    };

    let operator: Option<String> = cfg.get_item("operator")?.and_then(|v| v.extract().ok());
    let operator = match operator.as_deref().unwrap_or("between") {
        "between" | "Between" => ValidationOperator::Between,
        "notBetween" | "NotBetween" | "not_between" => ValidationOperator::NotBetween,
        "equal" | "Equal" | "==" => ValidationOperator::Equal,
        "notEqual" | "NotEqual" | "not_equal" | "!=" => ValidationOperator::NotEqual,
        "greaterThan" | "GreaterThan" | "greater_than" | ">" => ValidationOperator::GreaterThan,
        "lessThan" | "LessThan" | "less_than" | "<" => ValidationOperator::LessThan,
        "greaterThanOrEqual" | "GreaterThanOrEqual" | ">=" => {
            ValidationOperator::GreaterThanOrEqual
        }
        "lessThanOrEqual" | "LessThanOrEqual" | "<=" => ValidationOperator::LessThanOrEqual,
        _ => ValidationOperator::Between,
    };

    let formula_a: Option<String> = cfg.get_item("formula1")?.and_then(|v| v.extract().ok());
    let formula_b: Option<String> = cfg.get_item("formula2")?.and_then(|v| v.extract().ok());
    let allow_blank: bool = cfg
        .get_item("allow_blank")?
        .and_then(|v| v.extract::<bool>().ok())
        .unwrap_or(true);

    let error_title: Option<String> = cfg
        .get_item("error_title")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let error_message: Option<String> = cfg
        .get_item("error")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });

    Ok(Some(DataValidation {
        sqref: range,
        validation_type,
        operator,
        formula_a,
        formula_b,
        allow_blank,
        show_dropdown: true,
        show_error_message: true,
        error_style: ErrorStyle::Stop,
        error_title,
        error_message,
        show_input_message: false,
        input_title: None,
        input_message: None,
    }))
}

/// Build a `Table` from a cfg dict, or `None` for no-op.
pub(crate) fn dict_to_table(cfg: &Bound<'_, PyDict>) -> PyResult<Option<Table>> {
    let name: Option<String> = cfg.get_item("name")?.and_then(|v| v.extract().ok());
    let ref_range: Option<String> = cfg.get_item("ref")?.and_then(|v| v.extract().ok());

    let (Some(name), Some(ref_range)) = (name, ref_range) else {
        return Ok(None);
    };

    let style: Option<String> = cfg
        .get_item("style")?
        .and_then(|v| v.extract::<String>().ok())
        .and_then(|s| if s.is_empty() { None } else { Some(s) });
    let header_row: bool = cfg
        .get_item("header_row")?
        .and_then(|v| v.extract::<bool>().ok())
        .unwrap_or(true);
    let totals_row: bool = cfg
        .get_item("totals_row")?
        .and_then(|v| v.extract::<bool>().ok())
        .unwrap_or(false);
    let autofilter: bool = cfg
        .get_item("autofilter")?
        .and_then(|v| v.extract::<bool>().ok())
        .unwrap_or(true);

    let mut columns: Vec<TableColumn> = Vec::new();
    if let Some(v) = cfg.get_item("columns")? {
        if let Ok(list) = v.extract::<Vec<String>>() {
            for col_name in list {
                columns.push(TableColumn {
                    name: col_name,
                    totals_function: None,
                    totals_label: None,
                });
            }
        }
    }

    let table_style: Option<TableStyle> = style.map(|s| TableStyle {
        name: s,
        show_first_column: false,
        show_last_column: false,
        show_row_stripes: true,
        show_column_stripes: false,
    });

    Ok(Some(Table {
        name,
        display_name: None,
        range: ref_range,
        columns,
        header_row,
        totals_row,
        style: table_style,
        autofilter,
    }))
}
