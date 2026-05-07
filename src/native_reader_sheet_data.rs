//! Sheet-data reader logic: cell value reads, sheet values (dict + plain),
//! formulas, cached formula values, and row-height/column-width lookups.
//! The cell-value primitives live in `native_reader_cell_helpers`; the
//! native-record dispatch lives in `native_reader_records`.

use std::collections::HashMap;

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use crate::native_reader_backend::{NativeXlsbBook, NativeXlsxBook};
use crate::native_reader_cell_helpers::{cell_to_dict, cell_to_plain, formula_to_py};
use crate::util::{a1_to_row_col, cell_blank};

type PyObject = Py<PyAny>;

// ---------- Cell value reads ----------

pub(crate) fn read_cell_value_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
    data_only: bool,
) -> PyResult<PyObject> {
    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let row = row0 + 1;
    let col = col0 + 1;
    let cell = {
        book.ensure_sheet_indexes(sheet)?;
        let index = book
            .sheet_cell_indexes
            .get(sheet)
            .and_then(|cells| cells.get(&(row, col)))
            .copied();
        let data = book.ensure_sheet(sheet)?;
        index.map(|idx| data.cells[idx].clone())
    };
    let Some(cell) = cell else {
        return cell_blank(py);
    };
    let number_format = book.number_format_for_cell(&cell);
    cell_to_dict(py, &cell, data_only, number_format, book.book.date1904())
}

pub(crate) fn read_cell_value_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
    data_only: bool,
) -> PyResult<PyObject> {
    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let row = row0 + 1;
    let col = col0 + 1;
    let cell = {
        book.ensure_sheet_indexes(sheet)?;
        let index = book
            .sheet_cell_indexes
            .get(sheet)
            .and_then(|cells| cells.get(&(row, col)))
            .copied();
        let data = book.ensure_sheet(sheet)?;
        index.map(|idx| data.cells[idx].clone())
    };
    let Some(cell) = cell else {
        return cell_blank(py);
    };
    let number_format = book.number_format_for_cell(&cell);
    cell_to_dict(py, &cell, data_only, number_format, book.book.date1904())
}

// ---------- Sheet values (dict / plain) ----------

pub(crate) fn read_sheet_values_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    cell_range: Option<&str>,
    data_only: bool,
) -> PyResult<PyObject> {
    let (min_row, min_col, max_row, max_col) = match book.resolve_window(sheet, cell_range)? {
        Some(bounds) => bounds,
        None => return Ok(PyList::empty(py).into()),
    };
    book.ensure_sheet_indexes(sheet)?;
    let cell_index = book
        .sheet_cell_indexes
        .get(sheet)
        .cloned()
        .unwrap_or_default();
    let cells = book.ensure_sheet(sheet)?.cells.clone();
    let outer = PyList::empty(py);
    let date1904 = book.book.date1904();
    for row in min_row..=max_row {
        let inner = PyList::empty(py);
        for col in min_col..=max_col {
            if let Some(cell) = cell_index.get(&(row, col)).map(|idx| &cells[*idx]) {
                let number_format = book.number_format_for_cell(cell);
                inner.append(cell_to_dict(py, cell, data_only, number_format, date1904)?)?;
            } else {
                inner.append(cell_blank(py)?)?;
            }
        }
        outer.append(inner)?;
    }
    Ok(outer.into())
}

pub(crate) fn read_sheet_values_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
    cell_range: Option<&str>,
    data_only: bool,
) -> PyResult<PyObject> {
    let window = book.resolve_window(sheet, cell_range)?;
    let Some((min_row, min_col, max_row, max_col)) = window else {
        return Ok(PyList::empty(py).into());
    };
    book.ensure_sheet_indexes(sheet)?;
    let cell_index = book
        .sheet_cell_indexes
        .get(sheet)
        .cloned()
        .unwrap_or_default();
    let data = book.ensure_sheet(sheet)?.clone();
    let date1904 = book.book.date1904();
    let outer = PyList::empty(py);
    for row in min_row..=max_row {
        let inner = PyList::empty(py);
        for col in min_col..=max_col {
            let cell = cell_index
                .get(&(row, col))
                .map(|idx| data.cells[*idx].clone());
            match cell {
                Some(cell) => {
                    let number_format = book.number_format_for_cell(&cell);
                    inner.append(cell_to_dict(py, &cell, data_only, number_format, date1904)?)?;
                }
                None => inner.append(cell_blank(py)?)?,
            }
        }
        outer.append(inner)?;
    }
    Ok(outer.into())
}

pub(crate) fn read_sheet_values_plain_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    cell_range: Option<&str>,
    data_only: bool,
) -> PyResult<PyObject> {
    let (min_row, min_col, max_row, max_col) = match book.resolve_window(sheet, cell_range)? {
        Some(bounds) => bounds,
        None => return Ok(PyList::empty(py).into()),
    };
    book.ensure_sheet_indexes(sheet)?;
    let cell_index = book
        .sheet_cell_indexes
        .get(sheet)
        .cloned()
        .unwrap_or_default();
    let cells = book.ensure_sheet(sheet)?.cells.clone();
    let date1904 = book.book.date1904();
    let outer = PyList::empty(py);
    for row in min_row..=max_row {
        let inner = PyList::empty(py);
        for col in min_col..=max_col {
            if let Some(cell) = cell_index.get(&(row, col)).map(|idx| &cells[*idx]) {
                let number_format = book.number_format_for_cell(cell);
                inner.append(cell_to_plain(py, cell, data_only, number_format, date1904)?)?;
            } else {
                inner.append(py.None())?;
            }
        }
        outer.append(inner)?;
    }
    Ok(outer.into())
}

pub(crate) fn read_sheet_values_plain_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
    cell_range: Option<&str>,
    data_only: bool,
) -> PyResult<PyObject> {
    let window = book.resolve_window(sheet, cell_range)?;
    let Some((min_row, min_col, max_row, max_col)) = window else {
        return Ok(PyList::empty(py).into());
    };
    book.ensure_sheet_indexes(sheet)?;
    let cell_index = book
        .sheet_cell_indexes
        .get(sheet)
        .cloned()
        .unwrap_or_default();
    let data = book.ensure_sheet(sheet)?.clone();
    let date1904 = book.book.date1904();
    let outer = PyList::empty(py);
    for row in min_row..=max_row {
        let inner = PyList::empty(py);
        for col in min_col..=max_col {
            let cell = cell_index
                .get(&(row, col))
                .map(|idx| data.cells[*idx].clone());
            match cell {
                Some(cell) => {
                    let number_format = book.number_format_for_cell(&cell);
                    inner.append(cell_to_plain(
                        py,
                        &cell,
                        data_only,
                        number_format,
                        date1904,
                    )?)?;
                }
                None => inner.append(py.None())?,
            }
        }
        outer.append(inner)?;
    }
    Ok(outer.into())
}

// ---------- Formulas ----------

pub(crate) fn read_sheet_formulas_xlsx(
    book: &mut NativeXlsxBook,
    sheet: &str,
) -> PyResult<HashMap<(u32, u32), String>> {
    let data = book.ensure_sheet(sheet)?;
    Ok(data
        .cells
        .iter()
        .filter_map(|c| {
            c.formula
                .as_ref()
                .map(|f| ((c.row - 1, c.col - 1), f.clone()))
        })
        .collect())
}

pub(crate) fn read_sheet_formulas_xlsb(
    book: &mut NativeXlsbBook,
    sheet: &str,
) -> PyResult<HashMap<(u32, u32), String>> {
    let data = book.ensure_sheet(sheet)?;
    Ok(data
        .cells
        .iter()
        .filter_map(|c| {
            c.formula
                .as_ref()
                .map(|f| ((c.row - 1, c.col - 1), f.clone()))
        })
        .collect())
}

pub(crate) fn read_cell_formula_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
) -> PyResult<PyObject> {
    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let row = row0 + 1;
    let col = col0 + 1;
    let formula = book
        .ensure_sheet(sheet)?
        .cells
        .iter()
        .find(|cell| cell.row == row && cell.col == col)
        .and_then(|cell| cell.formula.as_deref())
        .map(str::to_string);
    match formula {
        Some(formula) => formula_to_py(py, &formula),
        None => Ok(py.None()),
    }
}

pub(crate) fn read_cached_formula_values_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let cells = book.ensure_sheet(sheet)?.cells.clone();
    let date1904 = book.book.date1904();
    let out = PyDict::new(py);
    for cell in &cells {
        if cell.formula.is_some() {
            let number_format = book.number_format_for_cell(cell);
            out.set_item(
                &cell.coordinate,
                cell_to_plain(py, cell, true, number_format, date1904)?,
            )?;
        }
    }
    Ok(out.into())
}

pub(crate) fn read_cached_formula_values_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let cells = book.ensure_sheet(sheet)?.cells.clone();
    let date1904 = book.book.date1904();
    let out = PyDict::new(py);
    for cell in &cells {
        if cell.formula.is_some() {
            let number_format = book.number_format_for_cell(cell);
            out.set_item(
                &cell.coordinate,
                cell_to_plain(py, cell, true, number_format, date1904)?,
            )?;
        }
    }
    Ok(out.into())
}

// ---------- Row height / column width ----------

pub(crate) fn read_row_height_xlsx(
    book: &mut NativeXlsxBook,
    sheet: &str,
    row: i64,
) -> PyResult<Option<f64>> {
    if row < 1 {
        return Ok(None);
    }
    Ok(book
        .ensure_sheet(sheet)?
        .row_heights
        .get(&(row as u32))
        .filter(|height| height.custom_height)
        .map(|height| height.height))
}

pub(crate) fn read_row_height_xlsb(
    book: &mut NativeXlsbBook,
    sheet: &str,
    row: i64,
) -> PyResult<Option<f64>> {
    if row < 1 {
        return Ok(None);
    }
    Ok(book
        .ensure_sheet(sheet)?
        .row_heights
        .get(&(row as u32))
        .filter(|height| height.custom_height)
        .map(|height| height.height))
}

pub(crate) fn read_column_width_xlsx(
    book: &mut NativeXlsxBook,
    sheet: &str,
    col_letter: &str,
) -> PyResult<Option<f64>> {
    let col = crate::native_reader_dimensions::col_letter_to_index_1based(col_letter)?;
    Ok(book
        .ensure_sheet(sheet)?
        .column_widths
        .iter()
        .find(|width| width.custom_width && col >= width.min && col <= width.max)
        .map(|width| crate::native_reader_dimensions::strip_excel_padding(width.width)))
}

pub(crate) fn read_column_width_xlsb(
    book: &mut NativeXlsbBook,
    sheet: &str,
    col_letter: &str,
) -> PyResult<Option<f64>> {
    let col = crate::native_reader_dimensions::col_letter_to_index_1based(col_letter)?;
    Ok(book
        .ensure_sheet(sheet)?
        .column_widths
        .iter()
        .find(|width| width.custom_width && col >= width.min && col <= width.max)
        .map(|width| crate::native_reader_dimensions::strip_excel_padding(width.width)))
}

// ---------- Re-export for the `native_reader_styles` array-formula fallback ----------

pub(crate) use crate::native_reader_cell_helpers::ensure_formula_prefix;
