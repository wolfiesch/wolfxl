//! PyO3 bridge for the native WolfXL XLSX reader.
//!
//! This class is opt-in while the native reader grows to parity. It preserves
//! the reader method shape Python already calls, but only the value/topology
//! subset is implemented in this first slice.

use std::collections::HashMap;

use chrono::{Datelike, Duration, NaiveDate, Timelike};
use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDateTime, PyDict, PyList};
use pyo3::IntoPyObjectExt;

type PyObject = Py<PyAny>;

use crate::util::{a1_to_row_col, cell_blank, cell_with_value};
use wolfxl_reader::{
    AlignmentInfo, AnchorExtentInfo, AnchorMarkerInfo, AnchorPositionInfo, ArrayFormulaInfo,
    BookViewInfo, BorderInfo, BreakInfo, CalcPropertiesInfo, Cell, CellDataType, CellValue,
    ChartAxisInfo, ChartDataLabelsInfo, ChartErrorBarsInfo, ChartInfo, ChartSeriesInfo,
    ChartTrendlineInfo, CustomPropertyInfo, DateGroupItemInfo, FillInfo, FilterColumnInfo,
    FilterInfo, FontInfo, HeaderFooterInfo, HeaderFooterItemInfo, ImageAnchorInfo, ImageInfo,
    InlineFontProps, NativeXlsxBook as NativeReaderBook, PageBreakListInfo, PageMarginsInfo,
    PageSetupInfo, PaneMode, SelectionInfo, SheetFormatInfo, SheetPropertiesInfo, SheetProtection,
    SheetState, SheetViewInfo, SortConditionInfo, SortStateInfo, WorkbookPropertiesInfo,
    WorkbookSecurity, WorksheetData,
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

    pub fn read_sheet_state(&self, sheet: &str) -> PyResult<&'static str> {
        if !self.sheet_names.iter().any(|name| name == sheet) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }
        let state = self.book.sheet_state(sheet).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("native sheet state read failed: {e}"))
        })?;
        Ok(match state {
            SheetState::Visible => "visible",
            SheetState::Hidden => "hidden",
            SheetState::VeryHidden => "veryHidden",
        })
    }

    pub fn read_print_area(&self, sheet: &str) -> PyResult<Option<String>> {
        if !self.sheet_names.iter().any(|name| name == sheet) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }
        Ok(self.book.print_area(sheet).map(str::to_string))
    }

    pub fn read_print_titles(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        if !self.sheet_names.iter().any(|name| name == sheet) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }
        let Some(titles) = self.book.print_titles(sheet) else {
            return Ok(py.None());
        };
        let d = PyDict::new(py);
        d.set_item("rows", titles.rows.as_deref())?;
        d.set_item("cols", titles.cols.as_deref())?;
        Ok(d.into())
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
        include_format = true,
        include_empty = false,
        include_formula_blanks = true,
        include_coordinate = true,
        include_style_id = true,
        include_extended_format = true,
        include_cached_formula_value = false,
    ))]
    #[allow(clippy::too_many_arguments)]
    pub fn read_sheet_records(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        cell_range: Option<&str>,
        data_only: bool,
        include_format: bool,
        include_empty: bool,
        include_formula_blanks: bool,
        include_coordinate: bool,
        include_style_id: bool,
        include_extended_format: bool,
        include_cached_formula_value: bool,
    ) -> PyResult<PyObject> {
        let window = self.resolve_window(sheet, cell_range)?;
        let data = self.ensure_sheet(sheet)?.clone();
        let cells_by_coord: HashMap<(u32, u32), Cell> = data
            .cells
            .iter()
            .cloned()
            .map(|cell| ((cell.row, cell.col), cell))
            .collect();
        let options = NativeRecordOptions {
            data_only,
            include_format,
            include_empty,
            include_formula_blanks,
            include_coordinate,
            include_style_id,
            include_extended_format,
            include_cached_formula_value,
        };
        let out = PyList::empty(py);

        if include_empty {
            if let Some((min_row, min_col, max_row, max_col)) = window {
                for row in min_row..=max_row {
                    for col in min_col..=max_col {
                        append_native_record(
                            py,
                            &out,
                            &self.book,
                            &data.merged_ranges,
                            cells_by_coord.get(&(row, col)),
                            row,
                            col,
                            options,
                        )?;
                    }
                }
                return Ok(out.into());
            }
        }

        for cell in &data.cells {
            if let Some((min_row, min_col, max_row, max_col)) = window {
                if cell.row < min_row
                    || cell.row > max_row
                    || cell.col < min_col
                    || cell.col > max_col
                {
                    continue;
                }
            }
            if !native_record_should_emit(cell, options) {
                continue;
            }
            append_native_record(
                py,
                &out,
                &self.book,
                &data.merged_ranges,
                Some(cell),
                cell.row,
                cell.col,
                options,
            )?;
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

    pub fn read_cell_formula(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let row = row0 + 1;
        let col = col0 + 1;
        let formula = self
            .ensure_sheet(sheet)?
            .cells
            .iter()
            .find(|cell| cell.row == row && cell.col == col)
            .and_then(|cell| cell.formula.as_deref());
        match formula {
            Some(formula) => formula_to_py(py, formula),
            None => Ok(py.None()),
        }
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

    pub fn read_sheet_view(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        match &self.ensure_sheet(sheet)?.sheet_view {
            Some(sheet_view) => sheet_view_to_py(py, sheet_view),
            None => Ok(py.None()),
        }
    }

    pub fn read_sheet_properties(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        match &self.ensure_sheet(sheet)?.sheet_properties {
            Some(properties) => sheet_properties_to_py(py, properties),
            None => Ok(py.None()),
        }
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

    pub fn read_sheet_protection(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let protection = self.ensure_sheet(sheet)?.sheet_protection.clone();
        match protection {
            Some(protection) => sheet_protection_to_py(py, &protection),
            None => Ok(py.None()),
        }
    }

    pub fn read_auto_filter(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let auto_filter = self.ensure_sheet(sheet)?.auto_filter.clone();
        match auto_filter {
            Some(auto_filter) => {
                let d = PyDict::new(py);
                d.set_item("ref", auto_filter.ref_range)?;
                let columns = PyList::empty(py);
                for column in &auto_filter.filter_columns {
                    columns.append(filter_column_to_py(py, column)?)?;
                }
                d.set_item("filter_columns", columns)?;
                match &auto_filter.sort_state {
                    Some(sort_state) => {
                        d.set_item("sort_state", sort_state_to_py(py, sort_state)?)?
                    }
                    None => d.set_item("sort_state", py.None())?,
                }
                Ok(d.into())
            }
            None => Ok(py.None()),
        }
    }

    pub fn read_images(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let images = self.ensure_sheet(sheet)?.images.clone();
        let result = PyList::empty(py);
        for image in &images {
            result.append(image_to_py(py, image)?)?;
        }
        Ok(result.into())
    }

    pub fn read_charts(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let charts = self.ensure_sheet(sheet)?.charts.clone();
        let result = PyList::empty(py);
        for chart in &charts {
            result.append(chart_to_py(py, chart)?)?;
        }
        Ok(result.into())
    }

    pub fn read_page_margins(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        match self.ensure_sheet(sheet)?.page_margins {
            Some(margins) => page_margins_to_py(py, &margins),
            None => Ok(py.None()),
        }
    }

    pub fn read_page_setup(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        match &self.ensure_sheet(sheet)?.page_setup {
            Some(setup) => page_setup_to_py(py, setup),
            None => Ok(py.None()),
        }
    }

    pub fn read_header_footer(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        match &self.ensure_sheet(sheet)?.header_footer {
            Some(header_footer) => header_footer_to_py(py, header_footer),
            None => Ok(py.None()),
        }
    }

    pub fn read_row_breaks(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        match &self.ensure_sheet(sheet)?.row_breaks {
            Some(breaks) => page_breaks_to_py(py, breaks),
            None => Ok(py.None()),
        }
    }

    pub fn read_column_breaks(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        match &self.ensure_sheet(sheet)?.column_breaks {
            Some(breaks) => page_breaks_to_py(py, breaks),
            None => Ok(py.None()),
        }
    }

    pub fn read_sheet_format(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        match &self.ensure_sheet(sheet)?.sheet_format {
            Some(format) => sheet_format_to_py(py, format),
            None => Ok(py.None()),
        }
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
            d.set_item("comment", table.comment.clone())?;
            d.set_item("table_type", table.table_type.clone())?;
            d.set_item("totals_row_shown", table.totals_row_shown)?;
            match &table.style {
                Some(style) => d.set_item("style", style)?,
                None => d.set_item("style", py.None())?,
            }
            d.set_item("show_first_column", table.show_first_column)?;
            d.set_item("show_last_column", table.show_last_column)?;
            d.set_item("show_row_stripes", table.show_row_stripes)?;
            d.set_item("show_column_stripes", table.show_column_stripes)?;
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

    pub fn read_custom_doc_properties(&self, py: Python<'_>) -> PyResult<PyObject> {
        let result = PyList::empty(py);
        for property in self.book.custom_doc_properties() {
            result.append(custom_property_to_py(py, property)?)?;
        }
        Ok(result.into())
    }

    pub fn read_workbook_security(&self, py: Python<'_>) -> PyResult<PyObject> {
        workbook_security_to_py(py, self.book.workbook_security())
    }

    pub fn read_workbook_properties(&self, py: Python<'_>) -> PyResult<PyObject> {
        match self.book.workbook_properties() {
            Some(properties) => workbook_properties_to_py(py, properties),
            None => Ok(py.None()),
        }
    }

    pub fn read_calc_properties(&self, py: Python<'_>) -> PyResult<PyObject> {
        match self.book.calc_properties() {
            Some(properties) => calc_properties_to_py(py, properties),
            None => Ok(py.None()),
        }
    }

    pub fn read_workbook_views(&self, py: Python<'_>) -> PyResult<PyObject> {
        let result = PyList::empty(py);
        for view in self.book.workbook_views() {
            result.append(book_view_to_py(py, view)?)?;
        }
        Ok(result.into())
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

    pub fn read_cell_array_formula(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let row = row0 + 1;
        let col = col0 + 1;
        let Some(info) = self
            .ensure_sheet(sheet)?
            .array_formulas
            .get(&(row, col))
            .cloned()
        else {
            return Ok(py.None());
        };
        let d = PyDict::new(py);
        match info {
            ArrayFormulaInfo::Array { ref_range, text } => {
                d.set_item("kind", "array")?;
                d.set_item("ref", ref_range)?;
                d.set_item("text", text)?;
            }
            ArrayFormulaInfo::DataTable {
                ref_range,
                ca,
                dt2_d,
                dtr,
                r1,
                r2,
            } => {
                d.set_item("kind", "data_table")?;
                d.set_item("ref", ref_range)?;
                d.set_item("ca", ca)?;
                d.set_item("dt2D", dt2_d)?;
                d.set_item("dtr", dtr)?;
                d.set_item("r1", r1)?;
                d.set_item("r2", r2)?;
            }
            ArrayFormulaInfo::SpillChild => {
                d.set_item("kind", "spill_child")?;
            }
        }
        Ok(d.into())
    }

    pub fn read_cell_rich_text(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let row = row0 + 1;
        let col = col0 + 1;
        let runs = {
            let data = self.ensure_sheet(sheet)?;
            data.cells
                .iter()
                .find(|c| c.row == row && c.col == col)
                .and_then(|cell| cell.rich_text.clone())
        };
        let Some(runs) = runs else {
            return Ok(py.None());
        };
        let out = PyList::empty(py);
        for run in runs {
            let item = PyList::empty(py);
            item.append(run.text)?;
            match run.font {
                Some(font) => item.append(rich_font_to_py(py, &font)?)?,
                None => item.append(py.None())?,
            }
            out.append(item)?;
        }
        Ok(out.into())
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
            if is_merged_subordinate(&data.merged_ranges, row, col) {
                return Ok(PyDict::new(py).into());
            }
            data.cells
                .iter()
                .find(|c| c.row == row && c.col == col)
                .and_then(|cell| cell.style_id)
        };
        let d = PyDict::new(py);
        if let Some(style_id) = style_id {
            if style_id == 0 {
                return Ok(d.into());
            }
            if let Some(font) = self.book.font_for_style_id(style_id) {
                populate_font(&d, font)?;
            }
            if let Some(fill) = self.book.fill_for_style_id(style_id) {
                populate_fill(&d, fill)?;
            }
            if let Some(number_format) = self.book.number_format_for_style_id(style_id) {
                d.set_item("number_format", number_format)?;
            }
            if let Some(alignment) = self.book.alignment_for_style_id(style_id) {
                populate_alignment(&d, alignment)?;
            }
        }
        Ok(d.into())
    }

    pub fn read_cell_border(
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
            if is_merged_subordinate(&data.merged_ranges, row, col) {
                return Ok(PyDict::new(py).into());
            }
            data.cells
                .iter()
                .find(|c| c.row == row && c.col == col)
                .and_then(|cell| cell.style_id)
        };
        let d = PyDict::new(py);
        if let Some(border) = style_id.and_then(|id| self.book.border_for_style_id(id)) {
            populate_border(py, &d, border)?;
        }
        Ok(d.into())
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

#[derive(Clone, Copy)]
struct NativeRecordOptions {
    data_only: bool,
    include_format: bool,
    include_empty: bool,
    include_formula_blanks: bool,
    include_coordinate: bool,
    include_style_id: bool,
    include_extended_format: bool,
    include_cached_formula_value: bool,
}

fn native_record_should_emit(cell: &Cell, options: NativeRecordOptions) -> bool {
    let has_formula = cell.formula.is_some();
    let has_value = !matches!(cell.value, CellValue::Empty);
    let should_emit_formula =
        has_formula && !options.data_only && (options.include_formula_blanks || has_value);
    options.include_empty || should_emit_formula || has_value
}

fn append_native_record(
    py: Python<'_>,
    out: &Bound<'_, PyList>,
    book: &NativeReaderBook,
    merged_ranges: &[String],
    cell: Option<&Cell>,
    row: u32,
    col: u32,
    options: NativeRecordOptions,
) -> PyResult<()> {
    let Some(cell) = cell else {
        if !options.include_empty {
            return Ok(());
        }
        let record = PyDict::new(py);
        record.set_item("row", row)?;
        record.set_item("column", col)?;
        record.set_item("data_type", "blank")?;
        record.set_item("value", py.None())?;
        if options.include_coordinate {
            record.set_item("coordinate", row_col_to_a1_1based(row, col))?;
        }
        out.append(record)?;
        return Ok(());
    };

    if !native_record_should_emit(cell, options) {
        return Ok(());
    }

    let is_merged_subordinate = is_merged_subordinate(merged_ranges, cell.row, cell.col);
    let number_format = if is_merged_subordinate {
        None
    } else {
        cell.style_id
            .and_then(|style_id| book.number_format_for_style_id(style_id))
    };
    let has_formula = cell.formula.is_some();
    let has_cached_value = !matches!(cell.value, CellValue::Empty);
    let should_emit_formula =
        has_formula && !options.data_only && (options.include_formula_blanks || has_cached_value);

    let record = PyDict::new(py);
    record.set_item("row", cell.row)?;
    record.set_item("column", cell.col)?;
    if options.include_coordinate {
        record.set_item("coordinate", &cell.coordinate)?;
    }
    if let Some(formula) = &cell.formula {
        record.set_item("formula", ensure_formula_prefix(formula))?;
        if options.include_cached_formula_value && has_cached_value {
            record.set_item(
                "cached_value",
                cell_to_plain(py, cell, true, number_format, book.date1904())?,
            )?;
        }
    }
    if should_emit_formula {
        let formula = ensure_formula_prefix(cell.formula.as_deref().unwrap_or_default());
        record.set_item("data_type", "formula")?;
        record.set_item("value", formula)?;
    } else if options.data_only && has_formula && !has_cached_value {
        record.set_item("data_type", "blank")?;
        record.set_item("value", py.None())?;
    } else {
        record.set_item(
            "data_type",
            native_data_type(cell, options.data_only, number_format),
        )?;
        record.set_item(
            "value",
            cell_to_plain(py, cell, options.data_only, number_format, book.date1904())?,
        )?;
    }

    if options.include_format && !is_merged_subordinate {
        populate_record_format(book, &record, cell.style_id, options)?;
    }
    out.append(record)?;
    Ok(())
}

fn populate_record_format(
    book: &NativeReaderBook,
    record: &Bound<'_, PyDict>,
    style_id: Option<u32>,
    options: NativeRecordOptions,
) -> PyResult<()> {
    let Some(style_id) = style_id else {
        return Ok(());
    };
    if options.include_style_id {
        record.set_item("style_id", style_id)?;
    }
    if style_id == 0 {
        return Ok(());
    }
    if options.include_extended_format {
        if let Some(font) = book.font_for_style_id(style_id) {
            if font.bold {
                record.set_item("bold", true)?;
            }
            if font.italic {
                record.set_item("italic", true)?;
            }
            if let Some(value) = &font.underline {
                record.set_item("underline", value)?;
            }
            if font.strikethrough {
                record.set_item("strikethrough", true)?;
            }
            if let Some(value) = font
                .size
                .filter(|value| (*value - 11.0).abs() > f64::EPSILON)
            {
                record.set_item("font_size", value)?;
            }
        }
        if let Some(fill) = book.fill_for_style_id(style_id) {
            if let Some(value) = &fill.bg_color {
                record.set_item("bg_color", value)?;
            }
        }
        if let Some(alignment) = book.alignment_for_style_id(style_id) {
            if let Some(value) = &alignment.horizontal {
                record.set_item("h_align", value)?;
            }
            if let Some(value) = &alignment.vertical {
                record.set_item("v_align", value)?;
            }
            if alignment.wrap_text {
                record.set_item("wrap", true)?;
            }
            if let Some(value) = alignment.text_rotation.filter(|value| *value != 0) {
                record.set_item("rotation", value)?;
            }
            if let Some(value) = alignment.indent.filter(|value| *value != 0) {
                record.set_item("indent", value)?;
            }
        }
        if let Some(border) = book.border_for_style_id(style_id) {
            if let Some(bottom) = &border.bottom {
                record.set_item("bottom_border_style", &bottom.style)?;
                record.set_item("has_bottom_border", true)?;
                if bottom.style == "double" {
                    record.set_item("is_double_underline", true)?;
                }
            }
        }
    }
    if let Some(number_format) = book.number_format_for_style_id(style_id) {
        record.set_item("number_format", number_format)?;
    }
    Ok(())
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
            return formula_to_py(py, formula);
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

fn rich_font_to_py(py: Python<'_>, font: &InlineFontProps) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    if let Some(value) = font.bold {
        d.set_item("b", value)?;
    }
    if let Some(value) = font.italic {
        d.set_item("i", value)?;
    }
    if let Some(value) = font.strike {
        d.set_item("strike", value)?;
    }
    if let Some(value) = &font.underline {
        d.set_item("u", value)?;
    }
    if let Some(value) = font.size {
        d.set_item("sz", value)?;
    }
    if let Some(value) = &font.color {
        d.set_item("color", value)?;
    }
    if let Some(value) = &font.name {
        d.set_item("rFont", value)?;
    }
    if let Some(value) = font.family {
        d.set_item("family", value)?;
    }
    if let Some(value) = font.charset {
        d.set_item("charset", value)?;
    }
    if let Some(value) = &font.vert_align {
        d.set_item("vertAlign", value)?;
    }
    if let Some(value) = &font.scheme {
        d.set_item("scheme", value)?;
    }
    Ok(d.into())
}

fn formula_to_py(py: Python<'_>, formula: &str) -> PyResult<PyObject> {
    let formula = ensure_formula_prefix(formula);
    let d = PyDict::new(py);
    d.set_item("type", "formula")?;
    d.set_item("formula", &formula)?;
    d.set_item("value", &formula)?;
    Ok(d.into())
}

fn sheet_protection_to_py(py: Python<'_>, protection: &SheetProtection) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("sheet", protection.sheet)?;
    d.set_item("objects", protection.objects)?;
    d.set_item("scenarios", protection.scenarios)?;
    d.set_item("format_cells", protection.format_cells)?;
    d.set_item("format_columns", protection.format_columns)?;
    d.set_item("format_rows", protection.format_rows)?;
    d.set_item("insert_columns", protection.insert_columns)?;
    d.set_item("insert_rows", protection.insert_rows)?;
    d.set_item("insert_hyperlinks", protection.insert_hyperlinks)?;
    d.set_item("delete_columns", protection.delete_columns)?;
    d.set_item("delete_rows", protection.delete_rows)?;
    d.set_item("select_locked_cells", protection.select_locked_cells)?;
    d.set_item("sort", protection.sort)?;
    d.set_item("auto_filter", protection.auto_filter)?;
    d.set_item("pivot_tables", protection.pivot_tables)?;
    d.set_item("select_unlocked_cells", protection.select_unlocked_cells)?;
    if let Some(password_hash) = &protection.password_hash {
        d.set_item("password_hash", password_hash)?;
    }
    Ok(d.into())
}

fn workbook_security_to_py(py: Python<'_>, security: &WorkbookSecurity) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    match &security.workbook_protection {
        Some(protection) => {
            let p = PyDict::new(py);
            p.set_item("lock_structure", protection.lock_structure)?;
            p.set_item("lock_windows", protection.lock_windows)?;
            p.set_item("lock_revision", protection.lock_revision)?;
            p.set_item(
                "workbook_algorithm_name",
                protection.workbook_algorithm_name.as_deref(),
            )?;
            p.set_item(
                "workbook_hash_value",
                protection.workbook_hash_value.as_deref(),
            )?;
            p.set_item(
                "workbook_salt_value",
                protection.workbook_salt_value.as_deref(),
            )?;
            p.set_item("workbook_spin_count", protection.workbook_spin_count)?;
            p.set_item(
                "revisions_algorithm_name",
                protection.revisions_algorithm_name.as_deref(),
            )?;
            p.set_item(
                "revisions_hash_value",
                protection.revisions_hash_value.as_deref(),
            )?;
            p.set_item(
                "revisions_salt_value",
                protection.revisions_salt_value.as_deref(),
            )?;
            p.set_item("revisions_spin_count", protection.revisions_spin_count)?;
            d.set_item("workbook_protection", p)?;
        }
        None => d.set_item("workbook_protection", py.None())?,
    }
    match &security.file_sharing {
        Some(file_sharing) => {
            let f = PyDict::new(py);
            f.set_item("read_only_recommended", file_sharing.read_only_recommended)?;
            f.set_item("user_name", file_sharing.user_name.as_deref())?;
            f.set_item("algorithm_name", file_sharing.algorithm_name.as_deref())?;
            f.set_item("hash_value", file_sharing.hash_value.as_deref())?;
            f.set_item("salt_value", file_sharing.salt_value.as_deref())?;
            f.set_item("spin_count", file_sharing.spin_count)?;
            d.set_item("file_sharing", f)?;
        }
        None => d.set_item("file_sharing", py.None())?,
    }
    Ok(d.into())
}

fn workbook_properties_to_py(
    py: Python<'_>,
    properties: &WorkbookPropertiesInfo,
) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("date1904", properties.date1904)?;
    d.set_item("date_compatibility", properties.date_compatibility)?;
    d.set_item("show_objects", properties.show_objects.as_deref())?;
    d.set_item(
        "show_border_unselected_tables",
        properties.show_border_unselected_tables,
    )?;
    d.set_item("filter_privacy", properties.filter_privacy)?;
    d.set_item("prompted_solutions", properties.prompted_solutions)?;
    d.set_item("show_ink_annotation", properties.show_ink_annotation)?;
    d.set_item("backup_file", properties.backup_file)?;
    d.set_item(
        "save_external_link_values",
        properties.save_external_link_values,
    )?;
    d.set_item("update_links", properties.update_links.as_deref())?;
    d.set_item("code_name", properties.code_name.as_deref())?;
    d.set_item("hide_pivot_field_list", properties.hide_pivot_field_list)?;
    d.set_item(
        "show_pivot_chart_filter",
        properties.show_pivot_chart_filter,
    )?;
    d.set_item("allow_refresh_query", properties.allow_refresh_query)?;
    d.set_item("publish_items", properties.publish_items)?;
    d.set_item("check_compatibility", properties.check_compatibility)?;
    d.set_item("auto_compress_pictures", properties.auto_compress_pictures)?;
    d.set_item(
        "refresh_all_connections",
        properties.refresh_all_connections,
    )?;
    d.set_item("default_theme_version", properties.default_theme_version)?;
    Ok(d.into())
}

fn calc_properties_to_py(py: Python<'_>, properties: &CalcPropertiesInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("calc_id", properties.calc_id)?;
    d.set_item("calc_mode", properties.calc_mode.as_deref())?;
    d.set_item("full_calc_on_load", properties.full_calc_on_load)?;
    d.set_item("ref_mode", properties.ref_mode.as_deref())?;
    d.set_item("iterate", properties.iterate)?;
    d.set_item("iterate_count", properties.iterate_count)?;
    d.set_item("iterate_delta", properties.iterate_delta)?;
    d.set_item("full_precision", properties.full_precision)?;
    d.set_item("calc_completed", properties.calc_completed)?;
    d.set_item("calc_on_save", properties.calc_on_save)?;
    d.set_item("concurrent_calc", properties.concurrent_calc)?;
    d.set_item(
        "concurrent_manual_count",
        properties.concurrent_manual_count,
    )?;
    d.set_item("force_full_calc", properties.force_full_calc)?;
    Ok(d.into())
}

fn book_view_to_py(py: Python<'_>, view: &BookViewInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("visibility", &view.visibility)?;
    d.set_item("minimized", view.minimized)?;
    d.set_item("show_horizontal_scroll", view.show_horizontal_scroll)?;
    d.set_item("show_vertical_scroll", view.show_vertical_scroll)?;
    d.set_item("show_sheet_tabs", view.show_sheet_tabs)?;
    d.set_item("x_window", view.x_window)?;
    d.set_item("y_window", view.y_window)?;
    d.set_item("window_width", view.window_width)?;
    d.set_item("window_height", view.window_height)?;
    d.set_item("tab_ratio", view.tab_ratio)?;
    d.set_item("first_sheet", view.first_sheet)?;
    d.set_item("active_tab", view.active_tab)?;
    d.set_item("auto_filter_date_grouping", view.auto_filter_date_grouping)?;
    Ok(d.into())
}

fn custom_property_to_py(py: Python<'_>, property: &CustomPropertyInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("name", &property.name)?;
    d.set_item("kind", &property.kind)?;
    d.set_item("value", &property.value)?;
    Ok(d.into())
}

fn image_to_py(py: Python<'_>, image: &ImageInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("data", PyBytes::new(py, &image.data))?;
    d.set_item("ext", &image.ext)?;
    d.set_item("anchor", image_anchor_to_py(py, &image.anchor)?)?;
    Ok(d.into())
}

fn chart_to_py(py: Python<'_>, chart: &ChartInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("kind", &chart.kind)?;
    d.set_item("title", chart.title.as_deref())?;
    d.set_item("x_axis_title", chart.x_axis_title.as_deref())?;
    d.set_item("y_axis_title", chart.y_axis_title.as_deref())?;
    d.set_item("x_axis", chart_axis_to_py(py, chart.x_axis.as_ref())?)?;
    d.set_item("y_axis", chart_axis_to_py(py, chart.y_axis.as_ref())?)?;
    d.set_item(
        "data_labels",
        chart_data_labels_to_py(py, chart.data_labels.as_ref())?,
    )?;
    d.set_item("legend_position", chart.legend_position.as_deref())?;
    d.set_item("bar_dir", chart.bar_dir.as_deref())?;
    d.set_item("grouping", chart.grouping.as_deref())?;
    d.set_item("scatter_style", chart.scatter_style.as_deref())?;
    d.set_item("vary_colors", chart.vary_colors)?;
    d.set_item("style", chart.style)?;
    d.set_item("anchor", image_anchor_to_py(py, &chart.anchor)?)?;
    let series = PyList::empty(py);
    for item in &chart.series {
        series.append(chart_series_to_py(py, item)?)?;
    }
    d.set_item("series", series)?;
    Ok(d.into())
}

fn chart_axis_to_py(py: Python<'_>, axis: Option<&ChartAxisInfo>) -> PyResult<PyObject> {
    let Some(axis) = axis else {
        return Ok(py.None());
    };
    let d = PyDict::new(py);
    d.set_item("axis_type", &axis.axis_type)?;
    d.set_item("axis_position", axis.axis_position.as_deref())?;
    d.set_item("ax_id", axis.ax_id)?;
    d.set_item("cross_ax", axis.cross_ax)?;
    d.set_item("scaling_min", axis.scaling_min)?;
    d.set_item("scaling_max", axis.scaling_max)?;
    d.set_item("scaling_orientation", axis.scaling_orientation.as_deref())?;
    d.set_item("scaling_log_base", axis.scaling_log_base)?;
    d.set_item("num_format_code", axis.num_format_code.as_deref())?;
    d.set_item("num_format_source_linked", axis.num_format_source_linked)?;
    d.set_item("major_unit", axis.major_unit)?;
    d.set_item("minor_unit", axis.minor_unit)?;
    d.set_item("tick_lbl_pos", axis.tick_lbl_pos.as_deref())?;
    d.set_item("major_tick_mark", axis.major_tick_mark.as_deref())?;
    d.set_item("minor_tick_mark", axis.minor_tick_mark.as_deref())?;
    d.set_item("crosses", axis.crosses.as_deref())?;
    d.set_item("crosses_at", axis.crosses_at)?;
    d.set_item("cross_between", axis.cross_between.as_deref())?;
    d.set_item("display_unit", axis.display_unit.as_deref())?;
    Ok(d.into())
}

fn chart_series_to_py(py: Python<'_>, series: &ChartSeriesInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("idx", series.idx)?;
    d.set_item("order", series.order)?;
    d.set_item("title_ref", series.title_ref.as_deref())?;
    d.set_item("title_value", series.title_value.as_deref())?;
    d.set_item(
        "data_labels",
        chart_data_labels_to_py(py, series.data_labels.as_ref())?,
    )?;
    d.set_item(
        "trendline",
        chart_trendline_to_py(py, series.trendline.as_ref())?,
    )?;
    d.set_item(
        "error_bars",
        chart_error_bars_to_py(py, series.error_bars.as_ref())?,
    )?;
    d.set_item("cat_ref", series.cat_ref.as_deref())?;
    d.set_item("val_ref", series.val_ref.as_deref())?;
    d.set_item("x_ref", series.x_ref.as_deref())?;
    d.set_item("y_ref", series.y_ref.as_deref())?;
    d.set_item("bubble_size_ref", series.bubble_size_ref.as_deref())?;
    Ok(d.into())
}

fn chart_data_labels_to_py(
    py: Python<'_>,
    labels: Option<&ChartDataLabelsInfo>,
) -> PyResult<PyObject> {
    let Some(labels) = labels else {
        return Ok(py.None());
    };
    let d = PyDict::new(py);
    d.set_item("position", labels.position.as_deref())?;
    d.set_item("show_legend_key", labels.show_legend_key)?;
    d.set_item("show_val", labels.show_val)?;
    d.set_item("show_cat_name", labels.show_cat_name)?;
    d.set_item("show_ser_name", labels.show_ser_name)?;
    d.set_item("show_percent", labels.show_percent)?;
    d.set_item("show_bubble_size", labels.show_bubble_size)?;
    d.set_item("show_leader_lines", labels.show_leader_lines)?;
    Ok(d.into())
}

fn chart_trendline_to_py(
    py: Python<'_>,
    trendline: Option<&ChartTrendlineInfo>,
) -> PyResult<PyObject> {
    let Some(trendline) = trendline else {
        return Ok(py.None());
    };
    let d = PyDict::new(py);
    d.set_item("trendline_type", trendline.trendline_type.as_deref())?;
    d.set_item("order", trendline.order)?;
    d.set_item("period", trendline.period)?;
    d.set_item("forward", trendline.forward)?;
    d.set_item("backward", trendline.backward)?;
    d.set_item("intercept", trendline.intercept)?;
    d.set_item("display_equation", trendline.display_equation)?;
    d.set_item("display_r_squared", trendline.display_r_squared)?;
    Ok(d.into())
}

fn chart_error_bars_to_py(
    py: Python<'_>,
    error_bars: Option<&ChartErrorBarsInfo>,
) -> PyResult<PyObject> {
    let Some(error_bars) = error_bars else {
        return Ok(py.None());
    };
    let d = PyDict::new(py);
    d.set_item("direction", error_bars.direction.as_deref())?;
    d.set_item("bar_type", error_bars.bar_type.as_deref())?;
    d.set_item("val_type", error_bars.val_type.as_deref())?;
    d.set_item("no_end_cap", error_bars.no_end_cap)?;
    d.set_item("val", error_bars.val)?;
    Ok(d.into())
}

fn page_margins_to_py(py: Python<'_>, margins: &PageMarginsInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("left", margins.left)?;
    d.set_item("right", margins.right)?;
    d.set_item("top", margins.top)?;
    d.set_item("bottom", margins.bottom)?;
    d.set_item("header", margins.header)?;
    d.set_item("footer", margins.footer)?;
    Ok(d.into())
}

fn page_setup_to_py(py: Python<'_>, setup: &PageSetupInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("orientation", setup.orientation.as_deref())?;
    d.set_item("paper_size", setup.paper_size)?;
    d.set_item("fit_to_width", setup.fit_to_width)?;
    d.set_item("fit_to_height", setup.fit_to_height)?;
    d.set_item("scale", setup.scale)?;
    d.set_item("first_page_number", setup.first_page_number)?;
    d.set_item("horizontal_dpi", setup.horizontal_dpi)?;
    d.set_item("vertical_dpi", setup.vertical_dpi)?;
    d.set_item("cell_comments", setup.cell_comments.as_deref())?;
    d.set_item("errors", setup.errors.as_deref())?;
    d.set_item("use_first_page_number", setup.use_first_page_number)?;
    d.set_item("use_printer_defaults", setup.use_printer_defaults)?;
    d.set_item("black_and_white", setup.black_and_white)?;
    d.set_item("draft", setup.draft)?;
    Ok(d.into())
}

fn header_footer_to_py(py: Python<'_>, header_footer: &HeaderFooterInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item(
        "odd_header",
        header_footer_item_to_py(py, &header_footer.odd_header)?,
    )?;
    d.set_item(
        "odd_footer",
        header_footer_item_to_py(py, &header_footer.odd_footer)?,
    )?;
    d.set_item(
        "even_header",
        header_footer_item_to_py(py, &header_footer.even_header)?,
    )?;
    d.set_item(
        "even_footer",
        header_footer_item_to_py(py, &header_footer.even_footer)?,
    )?;
    d.set_item(
        "first_header",
        header_footer_item_to_py(py, &header_footer.first_header)?,
    )?;
    d.set_item(
        "first_footer",
        header_footer_item_to_py(py, &header_footer.first_footer)?,
    )?;
    d.set_item("different_odd_even", header_footer.different_odd_even)?;
    d.set_item("different_first", header_footer.different_first)?;
    d.set_item("scale_with_doc", header_footer.scale_with_doc)?;
    d.set_item("align_with_margins", header_footer.align_with_margins)?;
    Ok(d.into())
}

fn header_footer_item_to_py(py: Python<'_>, item: &HeaderFooterItemInfo) -> PyResult<PyObject> {
    if item.left.is_none() && item.center.is_none() && item.right.is_none() {
        return Ok(py.None());
    }
    let d = PyDict::new(py);
    d.set_item("left", item.left.as_deref())?;
    d.set_item("center", item.center.as_deref())?;
    d.set_item("right", item.right.as_deref())?;
    Ok(d.into())
}

fn page_breaks_to_py(py: Python<'_>, breaks: &PageBreakListInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("count", breaks.count)?;
    d.set_item("manual_break_count", breaks.manual_break_count)?;
    let items = PyList::empty(py);
    for item in &breaks.breaks {
        items.append(break_to_py(py, item)?)?;
    }
    d.set_item("breaks", items)?;
    Ok(d.into())
}

fn break_to_py(py: Python<'_>, item: &BreakInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("id", item.id)?;
    d.set_item("min", item.min)?;
    d.set_item("max", item.max)?;
    d.set_item("man", item.man)?;
    d.set_item("pt", item.pt)?;
    Ok(d.into())
}

fn sheet_format_to_py(py: Python<'_>, format: &SheetFormatInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("base_col_width", format.base_col_width)?;
    d.set_item("default_col_width", format.default_col_width)?;
    d.set_item("default_row_height", format.default_row_height)?;
    d.set_item("custom_height", format.custom_height)?;
    d.set_item("zero_height", format.zero_height)?;
    d.set_item("thick_top", format.thick_top)?;
    d.set_item("thick_bottom", format.thick_bottom)?;
    d.set_item("outline_level_row", format.outline_level_row)?;
    d.set_item("outline_level_col", format.outline_level_col)?;
    Ok(d.into())
}

fn sheet_properties_to_py(py: Python<'_>, properties: &SheetPropertiesInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("code_name", properties.code_name.as_deref())?;
    d.set_item(
        "enable_format_conditions_calculation",
        properties.enable_format_conditions_calculation,
    )?;
    d.set_item("filter_mode", properties.filter_mode)?;
    d.set_item("published", properties.published)?;
    d.set_item("sync_horizontal", properties.sync_horizontal)?;
    d.set_item("sync_ref", properties.sync_ref.as_deref())?;
    d.set_item("sync_vertical", properties.sync_vertical)?;
    d.set_item("transition_evaluation", properties.transition_evaluation)?;
    d.set_item("transition_entry", properties.transition_entry)?;
    d.set_item("tab_color", properties.tab_color.as_deref())?;

    let outline = PyDict::new(py);
    outline.set_item("summary_below", properties.outline.summary_below)?;
    outline.set_item("summary_right", properties.outline.summary_right)?;
    outline.set_item("apply_styles", properties.outline.apply_styles)?;
    outline.set_item(
        "show_outline_symbols",
        properties.outline.show_outline_symbols,
    )?;
    d.set_item("outline", outline)?;

    let page_setup = PyDict::new(py);
    page_setup.set_item("auto_page_breaks", properties.page_setup.auto_page_breaks)?;
    page_setup.set_item("fit_to_page", properties.page_setup.fit_to_page)?;
    d.set_item("page_setup", page_setup)?;
    Ok(d.into())
}

fn sheet_view_to_py(py: Python<'_>, sheet_view: &SheetViewInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("zoom_scale", sheet_view.zoom_scale)?;
    d.set_item("zoom_scale_normal", sheet_view.zoom_scale_normal)?;
    d.set_item("view", &sheet_view.view)?;
    d.set_item("show_grid_lines", sheet_view.show_grid_lines)?;
    d.set_item("show_row_col_headers", sheet_view.show_row_col_headers)?;
    d.set_item("show_outline_symbols", sheet_view.show_outline_symbols)?;
    d.set_item("show_zeros", sheet_view.show_zeros)?;
    d.set_item("right_to_left", sheet_view.right_to_left)?;
    d.set_item("tab_selected", sheet_view.tab_selected)?;
    d.set_item("top_left_cell", sheet_view.top_left_cell.as_deref())?;
    d.set_item("workbook_view_id", sheet_view.workbook_view_id)?;
    match &sheet_view.pane {
        Some(pane) => d.set_item("pane", pane_to_py(py, pane)?)?,
        None => d.set_item("pane", py.None())?,
    }
    let selections = PyList::empty(py);
    for selection in &sheet_view.selections {
        selections.append(selection_to_py(py, selection)?)?;
    }
    d.set_item("selection", selections)?;
    Ok(d.into())
}

fn pane_to_py(py: Python<'_>, pane: &wolfxl_reader::FreezePane) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("x_split", pane.x_split.unwrap_or_default())?;
    d.set_item("y_split", pane.y_split.unwrap_or_default())?;
    d.set_item(
        "top_left_cell",
        pane.top_left_cell.as_deref().unwrap_or("A1"),
    )?;
    d.set_item(
        "active_pane",
        pane.active_pane.as_deref().unwrap_or("topLeft"),
    )?;
    d.set_item(
        "state",
        match pane.mode {
            PaneMode::Freeze => "frozen",
            PaneMode::Split => "split",
        },
    )?;
    Ok(d.into())
}

fn selection_to_py(py: Python<'_>, selection: &SelectionInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("active_cell", selection.active_cell.as_deref())?;
    d.set_item("sqref", selection.sqref.as_deref())?;
    d.set_item("pane", selection.pane.as_deref())?;
    d.set_item("active_cell_id", selection.active_cell_id)?;
    Ok(d.into())
}

fn image_anchor_to_py(py: Python<'_>, anchor: &ImageAnchorInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    match anchor {
        ImageAnchorInfo::OneCell { from, ext } => {
            d.set_item("type", "one_cell")?;
            populate_marker(&d, "from", from)?;
            match ext {
                Some(ext) => populate_extent(&d, ext)?,
                None => {
                    d.set_item("cx_emu", py.None())?;
                    d.set_item("cy_emu", py.None())?;
                }
            }
        }
        ImageAnchorInfo::TwoCell { from, to, edit_as } => {
            d.set_item("type", "two_cell")?;
            populate_marker(&d, "from", from)?;
            populate_marker(&d, "to", to)?;
            d.set_item("edit_as", edit_as)?;
        }
        ImageAnchorInfo::Absolute { pos, ext } => {
            d.set_item("type", "absolute")?;
            populate_position(&d, pos)?;
            populate_extent(&d, ext)?;
        }
    }
    Ok(d.into())
}

fn populate_marker(d: &Bound<'_, PyDict>, prefix: &str, marker: &AnchorMarkerInfo) -> PyResult<()> {
    d.set_item(format!("{prefix}_col"), marker.col)?;
    d.set_item(format!("{prefix}_row"), marker.row)?;
    d.set_item(format!("{prefix}_col_off"), marker.col_off)?;
    d.set_item(format!("{prefix}_row_off"), marker.row_off)?;
    Ok(())
}

fn populate_position(d: &Bound<'_, PyDict>, pos: &AnchorPositionInfo) -> PyResult<()> {
    d.set_item("x_emu", pos.x)?;
    d.set_item("y_emu", pos.y)?;
    Ok(())
}

fn populate_extent(d: &Bound<'_, PyDict>, ext: &AnchorExtentInfo) -> PyResult<()> {
    d.set_item("cx_emu", ext.cx)?;
    d.set_item("cy_emu", ext.cy)?;
    Ok(())
}

fn filter_column_to_py(py: Python<'_>, column: &FilterColumnInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("col_id", column.col_id)?;
    d.set_item("hidden_button", column.hidden_button)?;
    d.set_item("show_button", column.show_button)?;
    match &column.filter {
        Some(filter) => d.set_item("filter", filter_info_to_py(py, filter)?)?,
        None => d.set_item("filter", py.None())?,
    }
    let date_items = PyList::empty(py);
    for item in &column.date_group_items {
        date_items.append(date_group_item_to_py(py, item)?)?;
    }
    d.set_item("date_group_items", date_items)?;
    Ok(d.into())
}

fn filter_info_to_py(py: Python<'_>, filter: &FilterInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    match filter {
        FilterInfo::Blank => {
            d.set_item("kind", "blank")?;
        }
        FilterInfo::Color { dxf_id, cell_color } => {
            d.set_item("kind", "color")?;
            d.set_item("dxf_id", dxf_id)?;
            d.set_item("cell_color", cell_color)?;
        }
        FilterInfo::Custom { and_, filters } => {
            d.set_item("kind", "custom")?;
            d.set_item("and_", and_)?;
            let out = PyList::empty(py);
            for filter in filters {
                let item = PyDict::new(py);
                item.set_item("operator", &filter.operator)?;
                item.set_item("val", &filter.val)?;
                out.append(item)?;
            }
            d.set_item("filters", out)?;
        }
        FilterInfo::Dynamic {
            filter_type,
            val,
            val_iso,
            max_val_iso,
        } => {
            d.set_item("kind", "dynamic")?;
            d.set_item("type", filter_type)?;
            d.set_item("val", val)?;
            d.set_item("val_iso", val_iso)?;
            d.set_item("max_val_iso", max_val_iso)?;
        }
        FilterInfo::Icon { icon_set, icon_id } => {
            d.set_item("kind", "icon")?;
            d.set_item("icon_set", icon_set)?;
            d.set_item("icon_id", icon_id)?;
        }
        FilterInfo::String { values } => {
            d.set_item("kind", "string")?;
            d.set_item("values", values)?;
        }
        FilterInfo::Top10 {
            top,
            percent,
            val,
            filter_val,
        } => {
            d.set_item("kind", "top10")?;
            d.set_item("top", top)?;
            d.set_item("percent", percent)?;
            d.set_item("val", val)?;
            d.set_item("filter_val", filter_val)?;
        }
    }
    Ok(d.into())
}

fn date_group_item_to_py(py: Python<'_>, item: &DateGroupItemInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("year", item.year)?;
    d.set_item("month", item.month)?;
    d.set_item("day", item.day)?;
    d.set_item("hour", item.hour)?;
    d.set_item("minute", item.minute)?;
    d.set_item("second", item.second)?;
    d.set_item("date_time_grouping", &item.date_time_grouping)?;
    Ok(d.into())
}

fn sort_state_to_py(py: Python<'_>, state: &SortStateInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    let conditions = PyList::empty(py);
    for condition in &state.sort_conditions {
        conditions.append(sort_condition_to_py(py, condition)?)?;
    }
    d.set_item("sort_conditions", conditions)?;
    d.set_item("column_sort", state.column_sort)?;
    d.set_item("case_sensitive", state.case_sensitive)?;
    d.set_item("ref", state.ref_range.as_deref())?;
    Ok(d.into())
}

fn sort_condition_to_py(py: Python<'_>, condition: &SortConditionInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("ref", &condition.ref_range)?;
    d.set_item("descending", condition.descending)?;
    d.set_item("sort_by", &condition.sort_by)?;
    d.set_item("custom_list", condition.custom_list.as_deref())?;
    d.set_item("dxf_id", condition.dxf_id)?;
    d.set_item("icon_set", condition.icon_set.as_deref())?;
    d.set_item("icon_id", condition.icon_id)?;
    Ok(d.into())
}

fn populate_font(d: &Bound<'_, PyDict>, font: &FontInfo) -> PyResult<()> {
    if font.bold {
        d.set_item("bold", true)?;
    }
    if font.italic {
        d.set_item("italic", true)?;
    }
    if let Some(value) = &font.underline {
        d.set_item("underline", value)?;
    }
    if font.strikethrough {
        d.set_item("strikethrough", true)?;
    }
    if let Some(value) = &font.name {
        d.set_item("font_name", value)?;
    }
    if let Some(value) = font.size {
        d.set_item("font_size", value)?;
    }
    if let Some(value) = &font.color {
        d.set_item("font_color", value)?;
    }
    Ok(())
}

fn populate_fill(d: &Bound<'_, PyDict>, fill: &FillInfo) -> PyResult<()> {
    if let Some(value) = &fill.bg_color {
        d.set_item("bg_color", value)?;
    }
    Ok(())
}

fn populate_alignment(d: &Bound<'_, PyDict>, alignment: &AlignmentInfo) -> PyResult<()> {
    if let Some(value) = &alignment.horizontal {
        d.set_item("h_align", value)?;
    }
    if let Some(value) = &alignment.vertical {
        d.set_item("v_align", value)?;
    }
    if alignment.wrap_text {
        d.set_item("wrap", true)?;
    }
    if let Some(value) = alignment.text_rotation.filter(|value| *value != 0) {
        d.set_item("rotation", value)?;
    }
    if let Some(value) = alignment.indent.filter(|value| *value != 0) {
        d.set_item("indent", value)?;
    }
    Ok(())
}

fn populate_border(py: Python<'_>, d: &Bound<'_, PyDict>, border: &BorderInfo) -> PyResult<()> {
    if let Some(side) = &border.left {
        set_border_side(py, d, "left", side)?;
    }
    if let Some(side) = &border.right {
        set_border_side(py, d, "right", side)?;
    }
    if let Some(side) = &border.top {
        set_border_side(py, d, "top", side)?;
    }
    if let Some(side) = &border.bottom {
        set_border_side(py, d, "bottom", side)?;
    }
    if let Some(side) = &border.diagonal_up {
        set_border_side(py, d, "diagonal_up", side)?;
    }
    if let Some(side) = &border.diagonal_down {
        set_border_side(py, d, "diagonal_down", side)?;
    }
    Ok(())
}

fn set_border_side(
    py: Python<'_>,
    d: &Bound<'_, PyDict>,
    key: &str,
    side: &wolfxl_reader::BorderSide,
) -> PyResult<()> {
    let edge = PyDict::new(py);
    edge.set_item("style", &side.style)?;
    edge.set_item("color", &side.color)?;
    d.set_item(key, edge)?;
    Ok(())
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

fn is_merged_subordinate(merged_ranges: &[String], row: u32, col: u32) -> bool {
    merged_ranges.iter().any(|range| {
        parse_range_1based(range).is_some_and(|(min_row, min_col, max_row, max_col)| {
            row >= min_row
                && row <= max_row
                && col >= min_col
                && col <= max_col
                && !(row == min_row && col == min_col)
        })
    })
}

fn row_col_to_a1_1based(row: u32, col: u32) -> String {
    let mut n = col;
    let mut letters = String::new();
    while n > 0 {
        n -= 1;
        letters.insert(0, (b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    format!("{letters}{row}")
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
