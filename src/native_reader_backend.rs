//! PyO3 bridge for the native WolfXL XLSX reader.
//!
//! This class is opt-in while the native reader grows to parity. It preserves
//! the reader method shape Python already calls, but only the value/topology
//! subset is implemented in this first slice.

use std::collections::HashMap;

use pyo3::exceptions::{PyIOError, PyNotImplementedError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use pyo3::IntoPyObjectExt;

type PyObject = Py<PyAny>;

use crate::util::{a1_to_row_col, cell_blank, cell_with_value};
use wolfxl_reader::{
    Cell, CellDataType, CellValue, NativeXlsxBook as NativeReaderBook, WorksheetData,
};

#[pyclass(unsendable, module = "wolfxl._rust")]
pub struct NativeXlsxBook {
    book: NativeReaderBook,
    sheet_names: Vec<String>,
    sheet_cache: HashMap<String, WorksheetData>,
    opened_from_bytes: bool,
    source_path: Option<String>,
}

#[pymethods]
impl NativeXlsxBook {
    /// Open an XLSX/XLSM workbook from a filesystem path.
    #[staticmethod]
    #[pyo3(signature = (path, _permissive = false))]
    pub fn open(path: &str, _permissive: bool) -> PyResult<Self> {
        let book = NativeReaderBook::open_path(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("native xlsx open failed: {e}")))?;
        let sheet_names = book.sheet_names().into_iter().map(str::to_string).collect();
        Ok(Self {
            book,
            sheet_names,
            sheet_cache: HashMap::new(),
            opened_from_bytes: false,
            source_path: Some(path.to_string()),
        })
    }

    /// Open an XLSX/XLSM workbook from raw bytes.
    #[staticmethod]
    #[pyo3(signature = (data, _permissive = false))]
    pub fn open_from_bytes(data: &[u8], _permissive: bool) -> PyResult<Self> {
        let book = NativeReaderBook::open_bytes(data.to_vec())
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("native xlsx open failed: {e}")))?;
        let sheet_names = book.sheet_names().into_iter().map(str::to_string).collect();
        Ok(Self {
            book,
            sheet_names,
            sheet_cache: HashMap::new(),
            opened_from_bytes: true,
            source_path: None,
        })
    }

    pub fn sheet_names(&self) -> Vec<String> {
        self.sheet_names.clone()
    }

    pub fn opened_from_bytes(&self) -> bool {
        self.opened_from_bytes
    }

    pub fn source_path(&self) -> Option<String> {
        self.source_path.clone()
    }

    pub fn read_sheet_dimensions(&mut self, sheet: &str) -> PyResult<Option<(u32, u32)>> {
        let Some((_, _, max_row, max_col)) = self.read_bounds_1based(sheet)? else {
            return Ok(None);
        };
        Ok(Some((max_row, max_col)))
    }

    pub fn read_sheet_bounds(&mut self, sheet: &str) -> PyResult<Option<(u32, u32, u32, u32)>> {
        self.read_bounds_1based(sheet)
    }

    #[pyo3(signature = (sheet, a1, data_only = false))]
    pub fn read_cell_value(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
        data_only: bool,
    ) -> PyResult<PyObject> {
        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let row = row0 + 1;
        let col = col0 + 1;
        let data = self.ensure_sheet(sheet)?;
        let Some(cell) = data.cells.iter().find(|c| c.row == row && c.col == col) else {
            return cell_blank(py);
        };
        cell_to_dict(py, cell, data_only)
    }

    #[pyo3(signature = (sheet, cell_range = None, data_only = false))]
    pub fn read_sheet_values(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        cell_range: Option<&str>,
        data_only: bool,
    ) -> PyResult<PyObject> {
        let (min_row, min_col, max_row, max_col) = match self.resolve_window(sheet, cell_range)? {
            Some(bounds) => bounds,
            None => return Ok(PyList::empty(py).into()),
        };
        let data = self.ensure_sheet(sheet)?;
        let outer = PyList::empty(py);
        for row in min_row..=max_row {
            let inner = PyList::empty(py);
            for col in min_col..=max_col {
                if let Some(cell) = data.cells.iter().find(|c| c.row == row && c.col == col) {
                    inner.append(cell_to_dict(py, cell, data_only)?)?;
                } else {
                    inner.append(cell_blank(py)?)?;
                }
            }
            outer.append(inner)?;
        }
        Ok(outer.into())
    }

    #[pyo3(signature = (sheet, cell_range = None, data_only = false))]
    pub fn read_sheet_values_plain(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        cell_range: Option<&str>,
        data_only: bool,
    ) -> PyResult<PyObject> {
        let (min_row, min_col, max_row, max_col) = match self.resolve_window(sheet, cell_range)? {
            Some(bounds) => bounds,
            None => return Ok(PyList::empty(py).into()),
        };
        let data = self.ensure_sheet(sheet)?;
        let outer = PyList::empty(py);
        for row in min_row..=max_row {
            let inner = PyList::empty(py);
            for col in min_col..=max_col {
                if let Some(cell) = data.cells.iter().find(|c| c.row == row && c.col == col) {
                    inner.append(cell_to_plain(py, cell, data_only)?)?;
                } else {
                    inner.append(py.None())?;
                }
            }
            outer.append(inner)?;
        }
        Ok(outer.into())
    }

    #[pyo3(signature = (
        sheet,
        cell_range = None,
        data_only = false,
        _include_format = true,
        include_empty = false,
        _include_formula_blanks = true,
        include_coordinate = true,
        include_style_id = true,
        _include_extended_format = true,
        _include_cached_formula_value = false,
    ))]
    #[allow(clippy::too_many_arguments)]
    pub fn read_sheet_records(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        cell_range: Option<&str>,
        data_only: bool,
        _include_format: bool,
        include_empty: bool,
        _include_formula_blanks: bool,
        include_coordinate: bool,
        include_style_id: bool,
        _include_extended_format: bool,
        _include_cached_formula_value: bool,
    ) -> PyResult<PyObject> {
        let window = self.resolve_window(sheet, cell_range)?;
        let cells = self.ensure_sheet(sheet)?.cells.clone();
        let out = PyList::empty(py);
        for cell in &cells {
            if let Some((min_row, min_col, max_row, max_col)) = window {
                if cell.row < min_row
                    || cell.row > max_row
                    || cell.col < min_col
                    || cell.col > max_col
                {
                    continue;
                }
            }
            if !include_empty && matches!(cell.value, CellValue::Empty) && cell.formula.is_none() {
                continue;
            }
            let record = PyDict::new(py);
            record.set_item("row", cell.row)?;
            record.set_item("column", cell.col)?;
            record.set_item("value", cell_to_plain(py, cell, data_only)?)?;
            record.set_item("data_type", native_data_type(cell, data_only))?;
            if include_coordinate {
                record.set_item("coordinate", &cell.coordinate)?;
            }
            if include_style_id {
                record.set_item("style_id", cell.style_id)?;
            }
            if let Some(number_format) = self.number_format_for_cell(cell) {
                record.set_item("number_format", number_format)?;
            }
            if let Some(formula) = &cell.formula {
                record.set_item("formula", ensure_formula_prefix(formula))?;
            }
            out.append(record)?;
        }
        Ok(out.into())
    }

    pub fn read_sheet_formulas(&mut self, sheet: &str) -> PyResult<HashMap<(u32, u32), String>> {
        let data = self.ensure_sheet(sheet)?;
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

    pub fn read_cached_formula_values(
        &mut self,
        py: Python<'_>,
        sheet: &str,
    ) -> PyResult<PyObject> {
        let data = self.ensure_sheet(sheet)?;
        let out = PyDict::new(py);
        for cell in &data.cells {
            if cell.formula.is_some() {
                out.set_item(&cell.coordinate, cell_to_plain(py, cell, true)?)?;
            }
        }
        Ok(out.into())
    }

    pub fn read_merged_ranges(&mut self, sheet: &str) -> PyResult<Vec<String>> {
        Ok(self.ensure_sheet(sheet)?.merged_ranges.clone())
    }

    pub fn read_sheet_visibility(&self, py: Python<'_>, _sheet: &str) -> PyResult<PyObject> {
        let d = PyDict::new(py);
        d.set_item("hidden_rows", PyList::empty(py))?;
        d.set_item("hidden_columns", PyList::empty(py))?;
        d.set_item("row_outline_levels", PyList::empty(py))?;
        d.set_item("column_outline_levels", PyList::empty(py))?;
        Ok(d.into())
    }

    pub fn read_hyperlinks(&self, py: Python<'_>, _sheet: &str) -> PyResult<PyObject> {
        Ok(PyList::empty(py).into())
    }

    pub fn read_comments(&self, py: Python<'_>, _sheet: &str) -> PyResult<PyObject> {
        Ok(PyList::empty(py).into())
    }

    pub fn read_freeze_panes(&self, py: Python<'_>, _sheet: &str) -> PyResult<PyObject> {
        Ok(py.None())
    }

    pub fn read_conditional_formats(&self, py: Python<'_>, _sheet: &str) -> PyResult<PyObject> {
        Ok(PyList::empty(py).into())
    }

    pub fn read_data_validations(&self, py: Python<'_>, _sheet: &str) -> PyResult<PyObject> {
        Ok(PyList::empty(py).into())
    }

    pub fn read_named_ranges(&self, py: Python<'_>, _sheet: &str) -> PyResult<PyObject> {
        Ok(PyList::empty(py).into())
    }

    pub fn read_tables(&self, py: Python<'_>, _sheet: &str) -> PyResult<PyObject> {
        Ok(PyList::empty(py).into())
    }

    pub fn read_doc_properties(&self, py: Python<'_>) -> PyResult<PyObject> {
        Ok(PyDict::new(py).into())
    }

    pub fn read_row_height(&self, _sheet: &str, _row: i64) -> Option<f64> {
        None
    }

    pub fn read_column_width(&self, _sheet: &str, _col_letter: &str) -> Option<f64> {
        None
    }

    pub fn read_cell_array_formula(&self, py: Python<'_>, _sheet: &str, _a1: &str) -> PyObject {
        py.None()
    }

    pub fn read_cell_rich_text(&self, py: Python<'_>, _sheet: &str, _a1: &str) -> PyObject {
        py.None()
    }

    pub fn read_cell_format(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let row = row0 + 1;
        let col = col0 + 1;
        let style_id = {
            let data = self.ensure_sheet(sheet)?;
            data.cells
                .iter()
                .find(|c| c.row == row && c.col == col)
                .and_then(|cell| cell.style_id)
        };
        let d = PyDict::new(py);
        if let Some(style_id) = style_id {
            if let Some(number_format) = self.book.number_format_for_style_id(style_id) {
                d.set_item("number_format", number_format)?;
            }
        }
        Ok(d.into())
    }

    pub fn read_cell_border(&self, _sheet: &str, _a1: &str) -> PyResult<()> {
        Err(PyErr::new::<PyNotImplementedError, _>(
            "native XLSX border reads are not implemented yet; unset WOLFXL_NATIVE_READER",
        ))
    }
}

impl NativeXlsxBook {
    fn ensure_sheet(&mut self, sheet: &str) -> PyResult<&WorksheetData> {
        if !self.sheet_names.iter().any(|name| name == sheet) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }
        if !self.sheet_cache.contains_key(sheet) {
            let data = self.book.worksheet(sheet).map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("native sheet read failed: {e}"))
            })?;
            self.sheet_cache.insert(sheet.to_string(), data);
        }
        Ok(self.sheet_cache.get(sheet).unwrap())
    }

    fn read_bounds_1based(&mut self, sheet: &str) -> PyResult<Option<(u32, u32, u32, u32)>> {
        let data = self.ensure_sheet(sheet)?;
        let mut bounds: Option<(u32, u32, u32, u32)> = None;
        for cell in &data.cells {
            update_bounds(&mut bounds, cell.row, cell.col);
        }
        if bounds.is_none() {
            bounds = data.dimension.as_deref().and_then(parse_range_1based);
        }
        Ok(bounds)
    }

    fn resolve_window(
        &mut self,
        sheet: &str,
        cell_range: Option<&str>,
    ) -> PyResult<Option<(u32, u32, u32, u32)>> {
        if let Some(range) = cell_range.filter(|s| !s.is_empty()) {
            return parse_range_1based(range)
                .ok_or_else(|| {
                    PyErr::new::<PyValueError, _>(format!("Invalid cell range: {range}"))
                })
                .map(Some);
        }
        self.read_bounds_1based(sheet)
    }

    fn number_format_for_cell(&self, cell: &Cell) -> Option<&str> {
        cell.style_id
            .and_then(|style_id| self.book.number_format_for_style_id(style_id))
    }
}

fn cell_to_dict(py: Python<'_>, cell: &Cell, data_only: bool) -> PyResult<PyObject> {
    if !data_only {
        if let Some(formula) = &cell.formula {
            let formula = ensure_formula_prefix(formula);
            let d = PyDict::new(py);
            d.set_item("type", "formula")?;
            d.set_item("formula", &formula)?;
            d.set_item("value", &formula)?;
            return Ok(d.into());
        }
    }
    match &cell.value {
        CellValue::Empty => cell_blank(py),
        CellValue::String(s) => cell_with_value(py, "string", s),
        CellValue::Number(n) => cell_with_value(py, "number", *n),
        CellValue::Bool(b) => cell_with_value(py, "boolean", *b),
        CellValue::Error(e) => cell_with_value(py, "error", e),
    }
}

fn cell_to_plain(py: Python<'_>, cell: &Cell, data_only: bool) -> PyResult<PyObject> {
    if !data_only {
        if let Some(formula) = &cell.formula {
            return Ok(ensure_formula_prefix(formula).into_py_any(py)?);
        }
    }
    match &cell.value {
        CellValue::Empty => Ok(py.None()),
        CellValue::String(s) => Ok(s.clone().into_py_any(py)?),
        CellValue::Number(n) => Ok((*n).into_py_any(py)?),
        CellValue::Bool(b) => Ok((*b).into_py_any(py)?),
        CellValue::Error(e) => Ok(e.clone().into_py_any(py)?),
    }
}

fn native_data_type(cell: &Cell, data_only: bool) -> &'static str {
    if !data_only && cell.formula.is_some() {
        return "formula";
    }
    match cell.data_type {
        CellDataType::Bool => "boolean",
        CellDataType::Error => "error",
        CellDataType::InlineString | CellDataType::SharedString | CellDataType::FormulaString => {
            "string"
        }
        CellDataType::Number => {
            if matches!(cell.value, CellValue::Empty) {
                "blank"
            } else {
                "number"
            }
        }
    }
}

fn ensure_formula_prefix(formula: &str) -> String {
    if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    }
}

fn parse_range_1based(range: &str) -> Option<(u32, u32, u32, u32)> {
    let clean = range.replace('$', "").to_ascii_uppercase();
    let mut parts = clean.split(':');
    let start = parts.next()?;
    let end = parts.next().unwrap_or(start);
    let (start_row0, start_col0) = a1_to_row_col(start).ok()?;
    let (end_row0, end_col0) = a1_to_row_col(end).ok()?;
    let start_row = start_row0 + 1;
    let start_col = start_col0 + 1;
    let end_row = end_row0 + 1;
    let end_col = end_col0 + 1;
    Some((
        start_row.min(end_row),
        start_col.min(end_col),
        start_row.max(end_row),
        start_col.max(end_col),
    ))
}

fn update_bounds(bounds: &mut Option<(u32, u32, u32, u32)>, row: u32, col: u32) {
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
