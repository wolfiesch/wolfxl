//! Small carriers for calamine sheet-record emission.

use calamine_styles::Data;
use pyo3::prelude::*;
use pyo3::types::PyDict;

use crate::calamine_value_helpers::{
    data_is_formula_text, data_to_plain_py, data_type_name, map_error_formula, row_col_to_a1,
};

#[derive(Clone, Copy, Debug)]
pub(crate) struct SheetRecordOptions {
    pub(crate) data_only: bool,
    pub(crate) include_format: bool,
    pub(crate) include_empty: bool,
    pub(crate) include_formula_blanks: bool,
    pub(crate) include_coordinate: bool,
    pub(crate) include_style_id: bool,
    pub(crate) include_extended_format: bool,
    pub(crate) include_cached_formula_value: bool,
}

#[derive(Clone, Copy, Debug)]
pub(crate) struct SheetRecordDecision {
    pub(crate) value_is_formula_placeholder: bool,
    pub(crate) value_is_uncached_formula: bool,
    pub(crate) should_emit_formula: bool,
    pub(crate) should_emit: bool,
}

pub(crate) fn analyze_sheet_record(
    value: Option<&Data>,
    formula: Option<&str>,
    options: SheetRecordOptions,
) -> SheetRecordDecision {
    let value_is_formula_placeholder = formula
        .zip(value)
        .is_some_and(|(formula_text, v)| data_is_formula_text(v, formula_text));
    let value_is_uncached_formula = options.data_only && value_is_formula_placeholder;
    let has_value = value.is_some_and(|v| !matches!(v, Data::Empty))
        && !value_is_uncached_formula
        && !value_is_formula_placeholder;
    let has_formula_backing_entry =
        value.is_some_and(|v| !matches!(v, Data::Empty)) && !value_is_formula_placeholder;
    let should_emit_formula = formula.is_some()
        && !options.data_only
        && (options.include_formula_blanks || has_formula_backing_entry);
    let should_emit = options.include_empty || should_emit_formula || has_value;

    SheetRecordDecision {
        value_is_formula_placeholder,
        value_is_uncached_formula,
        should_emit_formula,
        should_emit,
    }
}

pub(crate) fn new_record_dict(
    py: Python<'_>,
    row: u32,
    col: u32,
    options: SheetRecordOptions,
) -> PyResult<Bound<'_, PyDict>> {
    let record = PyDict::new(py);
    record.set_item("row", row + 1)?;
    record.set_item("column", col + 1)?;
    if options.include_coordinate {
        record.set_item("coordinate", row_col_to_a1(row, col))?;
    }
    Ok(record)
}

pub(crate) fn populate_formula_fields(
    py: Python<'_>,
    record: &Bound<'_, PyDict>,
    value: Option<&Data>,
    formula_text: &str,
    options: SheetRecordOptions,
    decision: SheetRecordDecision,
) -> PyResult<()> {
    record.set_item("formula", formula_text)?;
    if options.include_cached_formula_value {
        if let Some(v) = value {
            if !matches!(v, Data::Empty) && !decision.value_is_formula_placeholder {
                record.set_item("cached_value", data_to_plain_py(py, v)?)?;
            }
        }
    }
    Ok(())
}

pub(crate) fn populate_record_value(
    py: Python<'_>,
    record: &Bound<'_, PyDict>,
    value: Option<&Data>,
    formula: Option<&str>,
    decision: SheetRecordDecision,
) -> PyResult<()> {
    if decision.should_emit_formula {
        let formula_text = formula.unwrap();
        if let Some(err_val) = map_error_formula(formula_text) {
            record.set_item("data_type", "error")?;
            record.set_item("value", err_val)?;
        } else {
            record.set_item("data_type", "formula")?;
            record.set_item("value", formula_text)?;
        }
    } else if let Some(v) = value {
        if decision.value_is_uncached_formula {
            record.set_item("data_type", "blank")?;
            record.set_item("value", py.None())?;
        } else {
            record.set_item("data_type", data_type_name(v))?;
            record.set_item("value", data_to_plain_py(py, v)?)?;
        }
    } else {
        record.set_item("data_type", "blank")?;
        record.set_item("value", py.None())?;
    }
    Ok(())
}
