//! PyO3 bridge for the native WolfXL XLSX reader.
//!
//! This class is opt-in while the native reader grows to parity. It preserves
//! the reader method shape Python already calls, but only the value/topology
//! subset is implemented in this first slice.

use std::collections::HashMap;

use chrono::{Datelike, Duration, NaiveDate, Timelike};
use pyo3::exceptions::{PyIOError, PyNotImplementedError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDateTime, PyDict, PyList};
use pyo3::IntoPyObjectExt;

type PyObject = Py<PyAny>;

use crate::util::{a1_to_row_col, cell_blank, cell_with_value};
use wolfxl_reader::{
    Cell, CellDataType, CellValue, NativeXlsxBook as NativeReaderBook, PaneMode, WorksheetData,
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
        let cell = {
            let data = self.ensure_sheet(sheet)?;
            data.cells
                .iter()
                .find(|c| c.row == row && c.col == col)
                .cloned()
        };
        let Some(cell) = cell else {
            return cell_blank(py);
        };
        let number_format = self.number_format_for_cell(&cell);
        cell_to_dict(py, &cell, data_only, number_format, self.book.date1904())
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
        let cells = self.ensure_sheet(sheet)?.cells.clone();
        let outer = PyList::empty(py);
        for row in min_row..=max_row {
            let inner = PyList::empty(py);
            for col in min_col..=max_col {
                if let Some(cell) = cells.iter().find(|c| c.row == row && c.col == col) {
                    let number_format = self.number_format_for_cell(cell);
                    inner.append(cell_to_dict(
                        py,
                        cell,
                        data_only,
                        number_format,
                        self.book.date1904(),
                    )?)?;
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
        let cells = self.ensure_sheet(sheet)?.cells.clone();
        let outer = PyList::empty(py);
        for row in min_row..=max_row {
            let inner = PyList::empty(py);
            for col in min_col..=max_col {
                if let Some(cell) = cells.iter().find(|c| c.row == row && c.col == col) {
                    let number_format = self.number_format_for_cell(cell);
                    inner.append(cell_to_plain(
                        py,
                        cell,
                        data_only,
                        number_format,
                        self.book.date1904(),
                    )?)?;
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
            let number_format = self.number_format_for_cell(cell);
            record.set_item(
                "value",
                cell_to_plain(py, cell, data_only, number_format, self.book.date1904())?,
            )?;
            record.set_item(
                "data_type",
                native_data_type(cell, data_only, number_format),
            )?;
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
        let cells = self.ensure_sheet(sheet)?.cells.clone();
        let date1904 = self.book.date1904();
        let out = PyDict::new(py);
        for cell in &cells {
            if cell.formula.is_some() {
                let number_format = self.number_format_for_cell(cell);
                out.set_item(
                    &cell.coordinate,
                    cell_to_plain(py, cell, true, number_format, date1904)?,
                )?;
            }
        }
        Ok(out.into())
    }

    pub fn read_merged_ranges(&mut self, sheet: &str) -> PyResult<Vec<String>> {
        Ok(self.ensure_sheet(sheet)?.merged_ranges.clone())
    }

    pub fn read_sheet_visibility(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let data = self.ensure_sheet(sheet)?;
        let d = PyDict::new(py);
        d.set_item("hidden_rows", data.hidden_rows.clone())?;
        d.set_item("hidden_columns", data.hidden_columns.clone())?;
        let row_levels = PyDict::new(py);
        for (row, level) in &data.row_outline_levels {
            row_levels.set_item(*row, *level)?;
        }
        d.set_item("row_outline_levels", row_levels)?;
        let column_levels = PyDict::new(py);
        for (col, level) in &data.column_outline_levels {
            column_levels.set_item(*col, *level)?;
        }
        d.set_item("column_outline_levels", column_levels)?;
        Ok(d.into())
    }

    pub fn read_hyperlinks(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let links = self.ensure_sheet(sheet)?.hyperlinks.clone();
        let result = PyList::empty(py);
        for link in &links {
            let d = PyDict::new(py);
            d.set_item("cell", &link.cell)?;
            d.set_item("target", &link.target)?;
            d.set_item("display", &link.display)?;
            match &link.tooltip {
                Some(tooltip) => d.set_item("tooltip", tooltip)?,
                None => d.set_item("tooltip", py.None())?,
            }
            d.set_item("internal", link.internal)?;
            result.append(d)?;
        }
        Ok(result.into())
    }

    pub fn read_comments(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let comments = self.ensure_sheet(sheet)?.comments.clone();
        let result = PyList::empty(py);
        for comment in &comments {
            let d = PyDict::new(py);
            d.set_item("cell", &comment.cell)?;
            d.set_item("text", &comment.text)?;
            d.set_item("author", &comment.author)?;
            d.set_item("threaded", comment.threaded)?;
            result.append(d)?;
        }
        Ok(result.into())
    }

    pub fn read_freeze_panes(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let d = PyDict::new(py);
        let Some(info) = self.ensure_sheet(sheet)?.freeze_panes.clone() else {
            return Ok(d.into());
        };
        d.set_item(
            "mode",
            match info.mode {
                PaneMode::Freeze => "freeze",
                PaneMode::Split => "split",
            },
        )?;
        if let Some(top_left_cell) = info.top_left_cell {
            d.set_item("top_left_cell", top_left_cell)?;
        }
        if let Some(x_split) = info.x_split {
            d.set_item("x_split", x_split)?;
        }
        if let Some(y_split) = info.y_split {
            d.set_item("y_split", y_split)?;
        }
        if let Some(active_pane) = info.active_pane {
            d.set_item("active_pane", active_pane)?;
        }
        Ok(d.into())
    }

    pub fn read_conditional_formats(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let rules = self.ensure_sheet(sheet)?.conditional_formats.clone();
        let result = PyList::empty(py);
        for rule in &rules {
            let d = PyDict::new(py);
            d.set_item("range", &rule.range)?;
            d.set_item("rule_type", &rule.rule_type)?;
            if let Some(operator) = &rule.operator {
                d.set_item("operator", operator)?;
            }
            if let Some(formula) = &rule.formula {
                d.set_item("formula", formula)?;
            }
            if let Some(priority) = rule.priority {
                d.set_item("priority", priority)?;
            }
            if let Some(stop_if_true) = rule.stop_if_true {
                d.set_item("stop_if_true", stop_if_true)?;
            }
            result.append(d)?;
        }
        Ok(result.into())
    }

    pub fn read_data_validations(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let validations = self.ensure_sheet(sheet)?.data_validations.clone();
        let result = PyList::empty(py);
        for validation in &validations {
            let d = PyDict::new(py);
            d.set_item("range", &validation.range)?;
            d.set_item("validation_type", &validation.validation_type)?;
            if let Some(operator) = &validation.operator {
                d.set_item("operator", operator)?;
            }
            if let Some(formula1) = &validation.formula1 {
                d.set_item("formula1", formula1)?;
            }
            if let Some(formula2) = &validation.formula2 {
                d.set_item("formula2", formula2)?;
            }
            d.set_item("allow_blank", validation.allow_blank)?;
            if let Some(error_title) = &validation.error_title {
                d.set_item("error_title", error_title)?;
            }
            if let Some(error) = &validation.error {
                d.set_item("error", error)?;
            }
            result.append(d)?;
        }
        Ok(result.into())
    }

    pub fn read_named_ranges(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        if !self.sheet_names.iter().any(|name| name == sheet) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }
        let result = PyList::empty(py);
        for named_range in self.book.named_ranges() {
            if named_range.scope == "sheet" {
                let refers_to = named_range.refers_to.trim_start_matches('=');
                let Some((sheet_part, _addr)) = refers_to.split_once('!') else {
                    continue;
                };
                if sheet_part.trim_matches('\'') != sheet {
                    continue;
                }
            }
            let d = PyDict::new(py);
            d.set_item("name", &named_range.name)?;
            d.set_item("scope", &named_range.scope)?;
            d.set_item("refers_to", &named_range.refers_to)?;
            result.append(d)?;
        }
        Ok(result.into())
    }

    pub fn read_tables(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let tables = self.ensure_sheet(sheet)?.tables.clone();
        let result = PyList::empty(py);
        for table in &tables {
            let d = PyDict::new(py);
            d.set_item("name", &table.name)?;
            d.set_item("ref", &table.ref_range)?;
            d.set_item("header_row", table.header_row)?;
            d.set_item("totals_row", table.totals_row)?;
            match &table.style {
                Some(style) => d.set_item("style", style)?,
                None => d.set_item("style", py.None())?,
            }
            d.set_item("columns", table.columns.clone())?;
            d.set_item("autofilter", table.autofilter)?;
            result.append(d)?;
        }
        Ok(result.into())
    }

    pub fn read_doc_properties(&self, py: Python<'_>) -> PyResult<PyObject> {
        let d = PyDict::new(py);
        for (key, value) in self.book.doc_properties() {
            d.set_item(key, value)?;
        }
        Ok(d.into())
    }

    pub fn read_row_height(&mut self, sheet: &str, row: i64) -> PyResult<Option<f64>> {
        if row < 1 {
            return Ok(None);
        }
        Ok(self
            .ensure_sheet(sheet)?
            .row_heights
            .get(&(row as u32))
            .filter(|height| height.custom_height)
            .map(|height| height.height))
    }

    pub fn read_column_width(&mut self, sheet: &str, col_letter: &str) -> PyResult<Option<f64>> {
        let col = col_letter_to_index_1based(col_letter)?;
        Ok(self
            .ensure_sheet(sheet)?
            .column_widths
            .iter()
            .find(|width| width.custom_width && col >= width.min && col <= width.max)
            .map(|width| strip_excel_padding(width.width)))
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
        for range in &data.merged_ranges {
            if let Some((min_row, min_col, max_row, max_col)) = parse_range_1based(range) {
                update_bounds(&mut bounds, min_row, min_col);
                update_bounds(&mut bounds, max_row, max_col);
            }
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

fn cell_to_dict(
    py: Python<'_>,
    cell: &Cell,
    data_only: bool,
    number_format: Option<&str>,
    date1904: bool,
) -> PyResult<PyObject> {
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
        CellValue::Number(n) if is_date_format(number_format) => {
            let dt = excel_serial_to_datetime(*n, date1904);
            let midnight = chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap();
            if dt.time() == midnight {
                cell_with_value(py, "date", dt.date().format("%Y-%m-%d").to_string())
            } else {
                cell_with_value(py, "datetime", dt.format("%Y-%m-%dT%H:%M:%S").to_string())
            }
        }
        CellValue::Number(n) => cell_with_value(py, "number", *n),
        CellValue::Bool(b) => cell_with_value(py, "boolean", *b),
        CellValue::Error(e) => cell_with_value(py, "error", e),
    }
}

fn cell_to_plain(
    py: Python<'_>,
    cell: &Cell,
    data_only: bool,
    number_format: Option<&str>,
    date1904: bool,
) -> PyResult<PyObject> {
    if !data_only {
        if let Some(formula) = &cell.formula {
            return Ok(ensure_formula_prefix(formula).into_py_any(py)?);
        }
    }
    match &cell.value {
        CellValue::Empty => Ok(py.None()),
        CellValue::String(s) => Ok(s.clone().into_py_any(py)?),
        CellValue::Number(n) if is_date_format(number_format) => {
            let dt = excel_serial_to_datetime(*n, date1904);
            let py_dt = PyDateTime::new(
                py,
                dt.year(),
                dt.month() as u8,
                dt.day() as u8,
                dt.hour() as u8,
                dt.minute() as u8,
                dt.second() as u8,
                0,
                None,
            )?;
            Ok(py_dt.into_any().unbind())
        }
        CellValue::Number(n) => Ok((*n).into_py_any(py)?),
        CellValue::Bool(b) => Ok((*b).into_py_any(py)?),
        CellValue::Error(e) => Ok(e.clone().into_py_any(py)?),
    }
}

fn native_data_type(cell: &Cell, data_only: bool, number_format: Option<&str>) -> &'static str {
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
            } else if is_date_format(number_format) {
                "datetime"
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

fn col_letter_to_index_1based(col: &str) -> PyResult<u32> {
    let mut idx = 0u32;
    for ch in col.chars() {
        if !ch.is_ascii_alphabetic() {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Invalid column letter: {col}"
            )));
        }
        idx = idx
            .checked_mul(26)
            .and_then(|value| value.checked_add((ch.to_ascii_uppercase() as u8 - b'A' + 1) as u32))
            .ok_or_else(|| {
                PyErr::new::<PyValueError, _>(format!("Invalid column letter: {col}"))
            })?;
    }
    if idx == 0 {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Invalid column letter: {col}"
        )));
    }
    Ok(idx)
}

fn strip_excel_padding(raw: f64) -> f64 {
    const CALIBRI_WIDTH_PADDING: f64 = 0.83203125;
    const ALT_WIDTH_PADDING: f64 = 0.7109375;
    const WIDTH_TOLERANCE: f64 = 0.0005;

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

fn is_date_format(format: Option<&str>) -> bool {
    let Some(format) = format else {
        return false;
    };
    let first = format.split(';').next().unwrap_or(format);
    let mut in_quote = false;
    let chars: Vec<char> = first.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let ch = chars[i];
        if ch == '"' {
            in_quote = !in_quote;
            i += 1;
            continue;
        }
        if in_quote {
            i += 1;
            continue;
        }
        if ch == '[' {
            let mut j = i + 1;
            while j < chars.len() && chars[j] != ']' {
                j += 1;
            }
            if j < chars.len() {
                let bracket: String = chars[i + 1..j].iter().collect();
                let lower = bracket.to_ascii_lowercase();
                if lower != "h"
                    && lower != "hh"
                    && lower != "m"
                    && lower != "mm"
                    && lower != "s"
                    && lower != "ss"
                {
                    i = j + 1;
                    continue;
                }
            }
        }
        if matches!(
            ch,
            'd' | 'D' | 'm' | 'M' | 'h' | 'H' | 'y' | 'Y' | 's' | 'S'
        ) {
            let prev = i.checked_sub(1).and_then(|idx| chars.get(idx)).copied();
            if prev != Some('_') && prev != Some('\\') {
                return true;
            }
        }
        i += 1;
    }
    false
}

fn excel_serial_to_datetime(serial: f64, date1904: bool) -> chrono::NaiveDateTime {
    let epoch = if date1904 {
        NaiveDate::from_ymd_opt(1904, 1, 1).unwrap()
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30).unwrap()
    };
    let mut days = serial.trunc() as i64;
    if !date1904 && serial > 0.0 && serial < 60.0 {
        days += 1;
    }
    let fraction = serial - serial.trunc();
    let millis = (fraction * 86_400_000.0).round() as i64;
    epoch.and_hms_opt(0, 0, 0).unwrap() + Duration::days(days) + Duration::milliseconds(millis)
}
