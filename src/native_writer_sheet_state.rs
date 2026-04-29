//! Sheet-state adapters for the native writer backend.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyDict;
use wolfxl_writer::model::Worksheet;
use wolfxl_writer::refs;

/// Apply oracle-shaped freeze/split pane settings to a worksheet.
pub(crate) fn apply_freeze_panes(ws: &mut Worksheet, settings: &Bound<'_, PyAny>) -> PyResult<()> {
    let dict = settings
        .cast::<PyDict>()
        .map_err(|_| PyValueError::new_err("settings must be a dict"))?;

    let inner: Option<Bound<'_, PyAny>> = dict.get_item("freeze")?;
    let cfg: &Bound<'_, PyDict> = match &inner {
        Some(v) => v.cast::<PyDict>().unwrap_or(dict),
        None => dict,
    };

    let mode: String = cfg
        .get_item("mode")?
        .and_then(|v| v.extract::<String>().ok())
        .unwrap_or_else(|| "freeze".to_string());

    if mode == "freeze" {
        let top_left: Option<String> = cfg
            .get_item("top_left_cell")?
            .and_then(|v| v.extract::<String>().ok());
        if let Some(cell) = top_left {
            let (row, col) = parse_a1_to_row_col(&cell)?;
            // freeze_row/col semantics in the model: rows above `freeze_row`
            // and columns left of `freeze_col` stay pinned; the top-left
            // cell's (row, col) is the freeze split point.
            ws.set_freeze(row, col, Some((row, col)));
        }
    } else if mode == "split" {
        let x_split: f64 = cfg
            .get_item("x_split")?
            .and_then(|v| v.extract::<f64>().ok())
            .unwrap_or(0.0);
        let y_split: f64 = cfg
            .get_item("y_split")?
            .and_then(|v| v.extract::<f64>().ok())
            .unwrap_or(0.0);
        ws.set_split(x_split, y_split, None);
    }

    Ok(())
}

fn parse_a1_to_row_col(a1: &str) -> PyResult<(u32, u32)> {
    let cleaned = a1.replace('$', "");
    refs::parse_a1(&cleaned)
        .ok_or_else(|| PyValueError::new_err(format!("Invalid A1 reference: {a1}")))
}
