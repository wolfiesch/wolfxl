//! Native workbook readers for WolfXL.
//!
//! This crate is the dependency-free-from-calamine reader foundation. The
//! first production target is XLSX/XLSM because those files already use ZIP +
//! OOXML helpers elsewhere in WolfXL. XLSB and XLS readers will grow beside
//! this API while preserving the same value-only public contract they have
//! today.

use std::collections::HashMap;
use std::fs;
use std::io::{Cursor, Read};
use std::path::{Path, PathBuf};

use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use wolfxl_formula::{translate as translate_formula, RefDelta};
use wolfxl_rels::{RelId, RelsGraph};
use zip::ZipArchive;

/// Native reader result type.
pub type Result<T> = std::result::Result<T, ReaderError>;

/// Errors surfaced by native readers.
#[derive(Debug)]
pub enum ReaderError {
    Io(std::io::Error),
    Zip(zip::result::ZipError),
    Xml(String),
    MissingPart(String),
    SheetNotFound(String),
    Unsupported(String),
}

impl std::fmt::Display for ReaderError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ReaderError::Io(e) => write!(f, "{e}"),
            ReaderError::Zip(e) => write!(f, "{e}"),
            ReaderError::Xml(e) => f.write_str(e),
            ReaderError::MissingPart(part) => write!(f, "missing workbook part: {part}"),
            ReaderError::SheetNotFound(sheet) => write!(f, "sheet not found: {sheet}"),
            ReaderError::Unsupported(msg) => f.write_str(msg),
        }
    }
}

impl std::error::Error for ReaderError {}

impl From<std::io::Error> for ReaderError {
    fn from(value: std::io::Error) -> Self {
        ReaderError::Io(value)
    }
}

impl From<zip::result::ZipError> for ReaderError {
    fn from(value: zip::result::ZipError) -> Self {
        ReaderError::Zip(value)
    }
}

/// Sheet metadata from `xl/workbook.xml` and `xl/_rels/workbook.xml.rels`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetInfo {
    pub name: String,
    pub sheet_id: Option<String>,
    pub state: SheetState,
    pub path: String,
}

/// Excel sheet visibility state.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetState {
    Visible,
    Hidden,
    VeryHidden,
}

impl Default for SheetState {
    fn default() -> Self {
        Self::Visible
    }
}

/// A decoded worksheet cell.
#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    /// 1-based row index.
    pub row: u32,
    /// 1-based column index.
    pub col: u32,
    /// A1 coordinate as written or inferred from row/column.
    pub coordinate: String,
    /// Raw `s` style id from the worksheet XML.
    pub style_id: Option<u32>,
    /// Cell type label from OOXML (`s`, `inlineStr`, `b`, `e`, `str`, or `n`).
    pub data_type: CellDataType,
    /// Cached/display value. Formula text lives in `formula`.
    pub value: CellValue,
    /// Formula text without a leading equals sign when present.
    pub formula: Option<String>,
    /// Raw OOXML formula kind (`array`, `dataTable`, `shared`, etc.).
    pub formula_kind: Option<String>,
    /// Raw OOXML shared-formula index (`si`) when present.
    pub formula_shared_index: Option<String>,
    /// Array/data-table formula metadata when this cell is the master.
    pub array_formula: Option<ArrayFormulaInfo>,
    /// Structured rich-text runs for shared-string or inline-string cells.
    pub rich_text: Option<Vec<RichTextRun>>,
}

/// Native cell value model shared by future readers.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Empty,
    String(String),
    Number(f64),
    Bool(bool),
    Error(String),
}

/// OOXML cell type classification.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellDataType {
    Number,
    SharedString,
    InlineString,
    FormulaString,
    Bool,
    Error,
}

/// Parsed worksheet data returned by the native XLSX reader.
#[derive(Debug, Clone, PartialEq)]
pub struct WorksheetData {
    pub dimension: Option<String>,
    pub merged_ranges: Vec<String>,
    pub hyperlinks: Vec<Hyperlink>,
    pub freeze_panes: Option<FreezePane>,
    pub sheet_properties: Option<SheetPropertiesInfo>,
    pub sheet_view: Option<SheetViewInfo>,
    pub comments: Vec<Comment>,
    pub row_heights: HashMap<u32, RowHeight>,
    pub column_widths: Vec<ColumnWidth>,
    pub data_validations: Vec<DataValidation>,
    pub sheet_protection: Option<SheetProtection>,
    pub auto_filter: Option<AutoFilterInfo>,
    pub page_margins: Option<PageMarginsInfo>,
    pub page_setup: Option<PageSetupInfo>,
    pub header_footer: Option<HeaderFooterInfo>,
    pub row_breaks: Option<PageBreakListInfo>,
    pub column_breaks: Option<PageBreakListInfo>,
    pub sheet_format: Option<SheetFormatInfo>,
    pub images: Vec<ImageInfo>,
    pub charts: Vec<ChartInfo>,
    pub tables: Vec<Table>,
    pub conditional_formats: Vec<ConditionalFormatRule>,
    pub hidden_rows: Vec<u32>,
    pub hidden_columns: Vec<u32>,
    pub row_outline_levels: Vec<(u32, u8)>,
    pub column_outline_levels: Vec<(u32, u8)>,
    pub array_formulas: HashMap<(u32, u32), ArrayFormulaInfo>,
    pub cells: Vec<Cell>,
}

/// Parsed worksheet hyperlink metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hyperlink {
    pub cell: String,
    pub target: String,
    pub display: String,
    pub tooltip: Option<String>,
    pub internal: bool,
}

/// Parsed worksheet pane metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FreezePane {
    pub mode: PaneMode,
    pub top_left_cell: Option<String>,
    pub x_split: Option<i64>,
    pub y_split: Option<i64>,
    pub active_pane: Option<String>,
}

/// Parsed worksheet view metadata from the first `<sheetView>`.
#[derive(Debug, Clone, PartialEq)]
pub struct SheetViewInfo {
    pub zoom_scale: u32,
    pub zoom_scale_normal: u32,
    pub view: String,
    pub show_grid_lines: bool,
    pub show_row_col_headers: bool,
    pub show_outline_symbols: bool,
    pub show_zeros: bool,
    pub right_to_left: bool,
    pub tab_selected: bool,
    pub top_left_cell: Option<String>,
    pub workbook_view_id: u32,
    pub pane: Option<FreezePane>,
    pub selections: Vec<SelectionInfo>,
}

/// Parsed worksheet selection metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SelectionInfo {
    pub active_cell: Option<String>,
    pub sqref: Option<String>,
    pub pane: Option<String>,
    pub active_cell_id: Option<u32>,
}

/// OOXML pane mode relevant to read compatibility.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PaneMode {
    Freeze,
    Split,
}

/// Parsed worksheet properties from `<sheetPr>`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SheetPropertiesInfo {
    pub code_name: Option<String>,
    pub enable_format_conditions_calculation: Option<bool>,
    pub filter_mode: Option<bool>,
    pub published: Option<bool>,
    pub sync_horizontal: Option<bool>,
    pub sync_ref: Option<String>,
    pub sync_vertical: Option<bool>,
    pub transition_evaluation: Option<bool>,
    pub transition_entry: Option<bool>,
    pub tab_color: Option<String>,
    pub outline: OutlineInfo,
    pub page_setup: PageSetupPropertiesInfo,
}

/// Parsed worksheet outline display properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OutlineInfo {
    pub summary_below: bool,
    pub summary_right: bool,
    pub apply_styles: bool,
    pub show_outline_symbols: bool,
}

/// Parsed worksheet page-setup flags from `<sheetPr><pageSetUpPr>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PageSetupPropertiesInfo {
    pub auto_page_breaks: bool,
    pub fit_to_page: bool,
}

impl Default for OutlineInfo {
    fn default() -> Self {
        Self {
            summary_below: true,
            summary_right: true,
            apply_styles: false,
            show_outline_symbols: true,
        }
    }
}

impl Default for PageSetupPropertiesInfo {
    fn default() -> Self {
        Self {
            auto_page_breaks: true,
            fit_to_page: false,
        }
    }
}

impl SheetPropertiesInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            code_name: attr_value(e, b"codeName"),
            enable_format_conditions_calculation: attr_bool(
                e,
                b"enableFormatConditionsCalculation",
            ),
            filter_mode: attr_bool(e, b"filterMode"),
            published: attr_bool(e, b"published"),
            sync_horizontal: attr_bool(e, b"syncHorizontal"),
            sync_ref: attr_value(e, b"syncRef"),
            sync_vertical: attr_bool(e, b"syncVertical"),
            transition_evaluation: attr_bool(e, b"transitionEvaluation"),
            transition_entry: attr_bool(e, b"transitionEntry"),
            ..Self::default()
        }
    }

    fn apply_tab_color(&mut self, e: &BytesStart<'_>) {
        self.tab_color = parse_ooxml_color(e);
    }

    fn apply_outline(&mut self, e: &BytesStart<'_>) {
        self.outline = OutlineInfo {
            summary_below: attr_bool_default(e, b"summaryBelow", true),
            summary_right: attr_bool_default(e, b"summaryRight", true),
            apply_styles: attr_bool_default(e, b"applyStyles", false),
            show_outline_symbols: attr_bool_default(e, b"showOutlineSymbols", true),
        };
    }

    fn apply_page_setup(&mut self, e: &BytesStart<'_>) {
        self.page_setup = PageSetupPropertiesInfo {
            auto_page_breaks: attr_bool_default(e, b"autoPageBreaks", true),
            fit_to_page: attr_bool_default(e, b"fitToPage", false),
        };
    }
}

impl SheetViewInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            zoom_scale: attr_u32(e, b"zoomScale").unwrap_or(100),
            zoom_scale_normal: attr_u32(e, b"zoomScaleNormal").unwrap_or(100),
            view: attr_value(e, b"view").unwrap_or_else(|| "normal".to_string()),
            show_grid_lines: attr_bool_default(e, b"showGridLines", true),
            show_row_col_headers: attr_bool_default(e, b"showRowColHeaders", true),
            show_outline_symbols: attr_bool_default(e, b"showOutlineSymbols", true),
            show_zeros: attr_bool_default(e, b"showZeros", true),
            right_to_left: attr_bool_default(e, b"rightToLeft", false),
            tab_selected: attr_bool_default(e, b"tabSelected", false),
            top_left_cell: attr_value(e, b"topLeftCell"),
            workbook_view_id: attr_u32(e, b"workbookViewId").unwrap_or_default(),
            pane: None,
            selections: Vec::new(),
        }
    }
}

impl SelectionInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            active_cell: attr_value(e, b"activeCell"),
            sqref: attr_value(e, b"sqref"),
            pane: attr_value(e, b"pane"),
            active_cell_id: attr_u32(e, b"activeCellId"),
        }
    }
}

/// Parsed worksheet comment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment {
    pub cell: String,
    pub text: String,
    pub author: String,
    pub threaded: bool,
}

/// Row dimension metadata from worksheet XML.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct RowHeight {
    pub height: f64,
    pub custom_height: bool,
}

/// Column dimension metadata from worksheet XML.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ColumnWidth {
    pub min: u32,
    pub max: u32,
    pub width: f64,
    pub custom_width: bool,
}

/// Parsed worksheet data-validation rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataValidation {
    pub range: String,
    pub validation_type: String,
    pub operator: Option<String>,
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    pub allow_blank: bool,
    pub error_title: Option<String>,
    pub error: Option<String>,
}

/// Parsed worksheet protection flags.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetProtection {
    pub sheet: bool,
    pub objects: bool,
    pub scenarios: bool,
    pub format_cells: bool,
    pub format_columns: bool,
    pub format_rows: bool,
    pub insert_columns: bool,
    pub insert_rows: bool,
    pub insert_hyperlinks: bool,
    pub delete_columns: bool,
    pub delete_rows: bool,
    pub select_locked_cells: bool,
    pub sort: bool,
    pub auto_filter: bool,
    pub pivot_tables: bool,
    pub select_unlocked_cells: bool,
    pub password_hash: Option<String>,
}

/// Parsed worksheet page margins.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PageMarginsInfo {
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub header: f64,
    pub footer: f64,
}

/// Parsed worksheet page setup.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PageSetupInfo {
    pub orientation: Option<String>,
    pub paper_size: Option<u32>,
    pub fit_to_width: Option<u32>,
    pub fit_to_height: Option<u32>,
    pub scale: Option<u32>,
    pub first_page_number: Option<u32>,
    pub horizontal_dpi: Option<u32>,
    pub vertical_dpi: Option<u32>,
    pub cell_comments: Option<String>,
    pub errors: Option<String>,
    pub use_first_page_number: Option<bool>,
    pub use_printer_defaults: Option<bool>,
    pub black_and_white: Option<bool>,
    pub draft: Option<bool>,
}

/// Parsed worksheet header/footer settings.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HeaderFooterInfo {
    pub odd_header: HeaderFooterItemInfo,
    pub odd_footer: HeaderFooterItemInfo,
    pub even_header: HeaderFooterItemInfo,
    pub even_footer: HeaderFooterItemInfo,
    pub first_header: HeaderFooterItemInfo,
    pub first_footer: HeaderFooterItemInfo,
    pub different_odd_even: bool,
    pub different_first: bool,
    pub scale_with_doc: bool,
    pub align_with_margins: bool,
}

/// Parsed left/center/right header/footer text.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HeaderFooterItemInfo {
    pub left: Option<String>,
    pub center: Option<String>,
    pub right: Option<String>,
}

/// Parsed row or column page-break list.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PageBreakListInfo {
    pub count: u32,
    pub manual_break_count: u32,
    pub breaks: Vec<BreakInfo>,
}

/// Parsed single row/column page break.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BreakInfo {
    pub id: u32,
    pub min: Option<u32>,
    pub max: Option<u32>,
    pub man: bool,
    pub pt: bool,
}

/// Parsed worksheet sheet-format defaults.
#[derive(Debug, Clone, PartialEq)]
pub struct SheetFormatInfo {
    pub base_col_width: u32,
    pub default_col_width: Option<f64>,
    pub default_row_height: f64,
    pub custom_height: bool,
    pub zero_height: bool,
    pub thick_top: bool,
    pub thick_bottom: bool,
    pub outline_level_row: u32,
    pub outline_level_col: u32,
}

impl PageSetupInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            orientation: attr_value(e, b"orientation"),
            paper_size: attr_u32(e, b"paperSize"),
            fit_to_width: attr_u32(e, b"fitToWidth"),
            fit_to_height: attr_u32(e, b"fitToHeight"),
            scale: attr_u32(e, b"scale"),
            first_page_number: attr_u32(e, b"firstPageNumber"),
            horizontal_dpi: attr_u32(e, b"horizontalDpi"),
            vertical_dpi: attr_u32(e, b"verticalDpi"),
            cell_comments: attr_value(e, b"cellComments"),
            errors: attr_value(e, b"errors"),
            use_first_page_number: attr_bool(e, b"useFirstPageNumber"),
            use_printer_defaults: attr_bool(e, b"usePrinterDefaults"),
            black_and_white: attr_bool(e, b"blackAndWhite"),
            draft: attr_bool(e, b"draft"),
        }
    }
}

impl HeaderFooterInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            different_odd_even: attr_bool_default(e, b"differentOddEven", false),
            different_first: attr_bool_default(e, b"differentFirst", false),
            scale_with_doc: attr_bool_default(e, b"scaleWithDoc", true),
            align_with_margins: attr_bool_default(e, b"alignWithMargins", true),
            ..Self::default()
        }
    }

    fn set_part(&mut self, part: HeaderFooterPart, item: HeaderFooterItemInfo) {
        match part {
            HeaderFooterPart::OddHeader => self.odd_header = item,
            HeaderFooterPart::OddFooter => self.odd_footer = item,
            HeaderFooterPart::EvenHeader => self.even_header = item,
            HeaderFooterPart::EvenFooter => self.even_footer = item,
            HeaderFooterPart::FirstHeader => self.first_header = item,
            HeaderFooterPart::FirstFooter => self.first_footer = item,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum HeaderFooterPart {
    OddHeader,
    OddFooter,
    EvenHeader,
    EvenFooter,
    FirstHeader,
    FirstFooter,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BreakListKind {
    Row,
    Column,
}

impl PageBreakListInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            count: attr_u32(e, b"count").unwrap_or_default(),
            manual_break_count: attr_u32(e, b"manualBreakCount").unwrap_or_default(),
            breaks: Vec::new(),
        }
    }
}

impl SheetFormatInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            base_col_width: attr_u32(e, b"baseColWidth").unwrap_or(8),
            default_col_width: attr_f64(e, b"defaultColWidth"),
            default_row_height: attr_f64(e, b"defaultRowHeight").unwrap_or(15.0),
            custom_height: attr_bool_default(e, b"customHeight", false),
            zero_height: attr_bool_default(e, b"zeroHeight", false),
            thick_top: attr_bool_default(e, b"thickTop", false),
            thick_bottom: attr_bool_default(e, b"thickBottom", false),
            outline_level_row: attr_u32(e, b"outlineLevelRow").unwrap_or_default(),
            outline_level_col: attr_u32(e, b"outlineLevelCol").unwrap_or_default(),
        }
    }
}

/// Parsed worksheet-level auto-filter metadata.
#[derive(Debug, Clone, PartialEq)]
pub struct AutoFilterInfo {
    pub ref_range: String,
    pub filter_columns: Vec<FilterColumnInfo>,
    pub sort_state: Option<SortStateInfo>,
}

/// Parsed worksheet-level auto-filter column.
#[derive(Debug, Clone, PartialEq)]
pub struct FilterColumnInfo {
    pub col_id: u32,
    pub hidden_button: bool,
    pub show_button: bool,
    pub filter: Option<FilterInfo>,
    pub date_group_items: Vec<DateGroupItemInfo>,
}

/// Parsed worksheet-level auto-filter predicate.
#[derive(Debug, Clone, PartialEq)]
pub enum FilterInfo {
    Blank,
    Color {
        dxf_id: u32,
        cell_color: bool,
    },
    Custom {
        and_: bool,
        filters: Vec<CustomFilterInfo>,
    },
    Dynamic {
        filter_type: String,
        val: Option<f64>,
        val_iso: Option<String>,
        max_val_iso: Option<String>,
    },
    Icon {
        icon_set: String,
        icon_id: u32,
    },
    String {
        values: Vec<String>,
    },
    Top10 {
        top: bool,
        percent: bool,
        val: f64,
        filter_val: Option<f64>,
    },
}

/// Parsed custom auto-filter predicate.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomFilterInfo {
    pub operator: String,
    pub val: String,
}

/// Parsed date-group auto-filter predicate.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DateGroupItemInfo {
    pub year: u32,
    pub month: Option<u32>,
    pub day: Option<u32>,
    pub hour: Option<u32>,
    pub minute: Option<u32>,
    pub second: Option<u32>,
    pub date_time_grouping: String,
}

/// Parsed worksheet-level auto-filter sort state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortStateInfo {
    pub sort_conditions: Vec<SortConditionInfo>,
    pub column_sort: bool,
    pub case_sensitive: bool,
    pub ref_range: Option<String>,
}

/// Parsed worksheet-level auto-filter sort condition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortConditionInfo {
    pub ref_range: String,
    pub descending: bool,
    pub sort_by: String,
    pub custom_list: Option<String>,
    pub dxf_id: Option<u32>,
    pub icon_set: Option<String>,
    pub icon_id: Option<u32>,
}

/// Parsed worksheet image payload and anchor metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageInfo {
    pub data: Vec<u8>,
    pub ext: String,
    pub anchor: ImageAnchorInfo,
}

/// Parsed worksheet chart payload and anchor metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartInfo {
    pub kind: String,
    pub title: Option<String>,
    pub style: Option<u32>,
    pub anchor: ImageAnchorInfo,
    pub series: Vec<ChartSeriesInfo>,
}

/// Parsed chart series references.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartSeriesInfo {
    pub idx: Option<u32>,
    pub order: Option<u32>,
    pub title_ref: Option<String>,
    pub title_value: Option<String>,
    pub cat_ref: Option<String>,
    pub val_ref: Option<String>,
    pub x_ref: Option<String>,
    pub y_ref: Option<String>,
    pub bubble_size_ref: Option<String>,
}

/// Parsed drawing anchor for an embedded image.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ImageAnchorInfo {
    OneCell {
        from: AnchorMarkerInfo,
        ext: Option<AnchorExtentInfo>,
    },
    TwoCell {
        from: AnchorMarkerInfo,
        to: AnchorMarkerInfo,
        edit_as: String,
    },
    Absolute {
        pos: AnchorPositionInfo,
        ext: AnchorExtentInfo,
    },
}

/// Parsed cell-relative drawing marker.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct AnchorMarkerInfo {
    pub col: i64,
    pub row: i64,
    pub col_off: i64,
    pub row_off: i64,
}

/// Parsed EMU drawing position.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct AnchorPositionInfo {
    pub x: i64,
    pub y: i64,
}

/// Parsed EMU drawing extent.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct AnchorExtentInfo {
    pub cx: i64,
    pub cy: i64,
}

/// Parsed worksheet table metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Table {
    pub name: String,
    pub ref_range: String,
    pub header_row: bool,
    pub totals_row: bool,
    pub style: Option<String>,
    pub columns: Vec<String>,
    pub autofilter: bool,
}

/// Parsed workbook defined name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedRange {
    pub name: String,
    pub scope: String,
    pub refers_to: String,
}

/// Parsed worksheet print-title ranges.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PrintTitlesInfo {
    pub rows: Option<String>,
    pub cols: Option<String>,
}

/// Parsed workbook-level protection and file-sharing metadata.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WorkbookSecurity {
    pub workbook_protection: Option<WorkbookProtection>,
    pub file_sharing: Option<FileSharing>,
}

/// Parsed `<workbookPr>` metadata.
#[derive(Debug, Clone, PartialEq)]
pub struct WorkbookPropertiesInfo {
    pub date1904: bool,
    pub date_compatibility: Option<bool>,
    pub show_objects: Option<String>,
    pub show_border_unselected_tables: Option<bool>,
    pub filter_privacy: Option<bool>,
    pub prompted_solutions: Option<bool>,
    pub show_ink_annotation: Option<bool>,
    pub backup_file: Option<bool>,
    pub save_external_link_values: Option<bool>,
    pub update_links: Option<String>,
    pub code_name: Option<String>,
    pub hide_pivot_field_list: Option<bool>,
    pub show_pivot_chart_filter: Option<bool>,
    pub allow_refresh_query: Option<bool>,
    pub publish_items: Option<bool>,
    pub check_compatibility: Option<bool>,
    pub auto_compress_pictures: Option<bool>,
    pub refresh_all_connections: Option<bool>,
    pub default_theme_version: Option<u32>,
}

/// Parsed `<calcPr>` metadata.
#[derive(Debug, Clone, PartialEq)]
pub struct CalcPropertiesInfo {
    pub calc_id: Option<u32>,
    pub calc_mode: Option<String>,
    pub full_calc_on_load: Option<bool>,
    pub ref_mode: Option<String>,
    pub iterate: Option<bool>,
    pub iterate_count: Option<u32>,
    pub iterate_delta: Option<f64>,
    pub full_precision: Option<bool>,
    pub calc_completed: Option<bool>,
    pub calc_on_save: Option<bool>,
    pub concurrent_calc: Option<bool>,
    pub concurrent_manual_count: Option<u32>,
    pub force_full_calc: Option<bool>,
}

/// Parsed `<workbookProtection>` metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorkbookProtection {
    pub lock_structure: bool,
    pub lock_windows: bool,
    pub lock_revision: bool,
    pub workbook_algorithm_name: Option<String>,
    pub workbook_hash_value: Option<String>,
    pub workbook_salt_value: Option<String>,
    pub workbook_spin_count: Option<u32>,
    pub revisions_algorithm_name: Option<String>,
    pub revisions_hash_value: Option<String>,
    pub revisions_salt_value: Option<String>,
    pub revisions_spin_count: Option<u32>,
}

/// Parsed `<fileSharing>` metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FileSharing {
    pub read_only_recommended: bool,
    pub user_name: Option<String>,
    pub algorithm_name: Option<String>,
    pub hash_value: Option<String>,
    pub salt_value: Option<String>,
    pub spin_count: Option<u32>,
}

/// Parsed worksheet conditional-formatting rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ConditionalFormatRule {
    pub range: String,
    pub rule_type: String,
    pub operator: Option<String>,
    pub formula: Option<String>,
    pub priority: Option<i64>,
    pub stop_if_true: Option<bool>,
}

/// Array/data-table formula metadata for a cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ArrayFormulaInfo {
    Array {
        ref_range: String,
        text: String,
    },
    DataTable {
        ref_range: String,
        ca: bool,
        dt2_d: bool,
        dtr: bool,
        r1: Option<String>,
        r2: Option<String>,
    },
    SpillChild,
}

/// Inline font properties attached to a rich-text run.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct InlineFontProps {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub strike: Option<bool>,
    pub underline: Option<String>,
    pub size: Option<f64>,
    pub color: Option<String>,
    pub name: Option<String>,
    pub family: Option<i32>,
    pub charset: Option<i32>,
    pub vert_align: Option<String>,
    pub scheme: Option<String>,
}

/// One structured rich-text run from `<si>` or inline `<is>` content.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct RichTextRun {
    pub text: String,
    pub font: Option<InlineFontProps>,
}

/// Native XLSX/XLSM workbook reader.
#[derive(Debug, Clone)]
pub struct NativeXlsxBook {
    bytes: Vec<u8>,
    sheets: Vec<SheetInfo>,
    named_ranges: Vec<NamedRange>,
    print_areas: HashMap<String, String>,
    print_titles: HashMap<String, PrintTitlesInfo>,
    workbook_security: WorkbookSecurity,
    workbook_properties: Option<WorkbookPropertiesInfo>,
    calc_properties: Option<CalcPropertiesInfo>,
    doc_properties: HashMap<String, String>,
    shared_strings: SharedStrings,
    styles: StyleTables,
    date1904: bool,
}

impl NativeXlsxBook {
    /// Open an OOXML workbook from disk.
    pub fn open_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_bytes(fs::read(path)?)
    }

    /// Open an OOXML workbook from bytes.
    pub fn open_bytes(bytes: impl Into<Vec<u8>>) -> Result<Self> {
        let bytes = bytes.into();
        let mut zip = zip_from_bytes(&bytes)?;
        let workbook_xml = read_part_required(&mut zip, "xl/workbook.xml")?;
        let workbook_rels = read_part_required(&mut zip, "xl/_rels/workbook.xml.rels")?;
        let rels = RelsGraph::parse(workbook_rels.as_bytes())
            .map_err(|e| ReaderError::Xml(format!("failed to parse workbook rels: {e}")))?;
        let (
            sheet_refs,
            date1904,
            named_ranges,
            print_areas,
            print_titles,
            workbook_security,
            workbook_properties,
            calc_properties,
        ) = parse_workbook(&workbook_xml)?;
        let sheets = resolve_sheet_paths(sheet_refs, &rels)?;
        let shared_strings = match read_part_optional(&mut zip, "xl/sharedStrings.xml")? {
            Some(xml) => parse_shared_strings(&xml)?,
            None => SharedStrings::default(),
        };
        let styles = match read_part_optional(&mut zip, "xl/styles.xml")? {
            Some(xml) => parse_style_tables(&xml)?,
            None => StyleTables::default(),
        };
        let mut doc_properties = HashMap::new();
        if let Some(xml) = read_part_optional(&mut zip, "docProps/core.xml")? {
            parse_doc_properties_into(&xml, &mut doc_properties, doc_property_core_key);
        }
        if let Some(xml) = read_part_optional(&mut zip, "docProps/app.xml")? {
            parse_doc_properties_into(&xml, &mut doc_properties, doc_property_app_key);
        }

        Ok(Self {
            bytes,
            sheets,
            named_ranges,
            print_areas,
            print_titles,
            workbook_security,
            workbook_properties,
            calc_properties,
            doc_properties,
            shared_strings,
            styles,
            date1904,
        })
    }

    /// Workbook sheet names in document order.
    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name.as_str()).collect()
    }

    /// Workbook sheet metadata in document order.
    pub fn sheets(&self) -> &[SheetInfo] {
        &self.sheets
    }

    /// Resolve a worksheet's workbook-tab visibility state.
    pub fn sheet_state(&self, sheet_name: &str) -> Result<SheetState> {
        self.sheets
            .iter()
            .find(|s| s.name == sheet_name)
            .map(|sheet| sheet.state)
            .ok_or_else(|| ReaderError::SheetNotFound(sheet_name.to_string()))
    }

    /// Workbook defined names.
    pub fn named_ranges(&self) -> &[NamedRange] {
        &self.named_ranges
    }

    /// Worksheet print area parsed from `_xlnm.Print_Area`, if present.
    pub fn print_area(&self, sheet_name: &str) -> Option<&str> {
        self.print_areas.get(sheet_name).map(String::as_str)
    }

    /// Worksheet print titles parsed from `_xlnm.Print_Titles`, if present.
    pub fn print_titles(&self, sheet_name: &str) -> Option<&PrintTitlesInfo> {
        self.print_titles.get(sheet_name)
    }

    /// Workbook protection and file-sharing blocks.
    pub fn workbook_security(&self) -> &WorkbookSecurity {
        &self.workbook_security
    }

    /// Workbook-level properties parsed from `<workbookPr>`.
    pub fn workbook_properties(&self) -> Option<&WorkbookPropertiesInfo> {
        self.workbook_properties.as_ref()
    }

    /// Workbook calculation properties parsed from `<calcPr>`.
    pub fn calc_properties(&self) -> Option<&CalcPropertiesInfo> {
        self.calc_properties.as_ref()
    }

    /// Workbook document properties parsed from `docProps/core.xml` and app.xml.
    pub fn doc_properties(&self) -> &HashMap<String, String> {
        &self.doc_properties
    }

    /// Whether the workbook uses the 1904 date system.
    pub fn date1904(&self) -> bool {
        self.date1904
    }

    /// Shared-string table as plain strings.
    pub fn shared_strings(&self) -> &[String] {
        &self.shared_strings.values
    }

    /// Resolve a style id to an Excel number format code.
    pub fn number_format_for_style_id(&self, style_id: u32) -> Option<&str> {
        self.styles.number_format_for_style_id(style_id)
    }

    /// Resolve a style id to cell border metadata.
    pub fn border_for_style_id(&self, style_id: u32) -> Option<&BorderInfo> {
        self.styles.border_for_style_id(style_id)
    }

    /// Resolve a style id to font metadata.
    pub fn font_for_style_id(&self, style_id: u32) -> Option<&FontInfo> {
        self.styles.font_for_style_id(style_id)
    }

    /// Resolve a style id to fill metadata.
    pub fn fill_for_style_id(&self, style_id: u32) -> Option<&FillInfo> {
        self.styles.fill_for_style_id(style_id)
    }

    /// Resolve a style id to alignment metadata.
    pub fn alignment_for_style_id(&self, style_id: u32) -> Option<&AlignmentInfo> {
        self.styles.alignment_for_style_id(style_id)
    }

    /// Parse a worksheet into sparse decoded cells.
    pub fn worksheet(&self, sheet_name: &str) -> Result<WorksheetData> {
        let Some(info) = self.sheets.iter().find(|s| s.name == sheet_name) else {
            return Err(ReaderError::SheetNotFound(sheet_name.to_string()));
        };
        let mut zip = zip_from_bytes(&self.bytes)?;
        let xml = read_part_required(&mut zip, &info.path)?;
        let rels = read_part_optional(&mut zip, &sheet_rels_path(&info.path))?
            .map(|xml| {
                RelsGraph::parse(xml.as_bytes())
                    .map_err(|e| ReaderError::Xml(format!("failed to parse sheet rels: {e}")))
            })
            .transpose()?;
        let comments = match rels.as_ref().and_then(comments_target) {
            Some(target) => read_part_optional(
                &mut zip,
                &join_and_normalize(&part_dir(&info.path), &target),
            )?
            .map(|xml| parse_comments(&xml))
            .transpose()?
            .unwrap_or_default(),
            None => Vec::new(),
        };
        let tables = read_tables(&mut zip, &info.path, &xml, rels.as_ref())?;
        let images = read_images(&mut zip, &info.path, rels.as_ref())?;
        let charts = read_charts(&mut zip, &info.path, rels.as_ref())?;
        let mut data =
            parse_worksheet(&xml, &self.shared_strings, rels.as_ref(), comments, tables)?;
        data.images = images;
        data.charts = charts;
        Ok(data)
    }
}

#[derive(Debug)]
struct SheetRef {
    name: String,
    sheet_id: Option<String>,
    state: SheetState,
    rid: String,
}

#[derive(Debug, Clone, Default, PartialEq)]
struct StyleTables {
    custom_num_fmts: HashMap<u32, String>,
    cell_xfs: Vec<XfEntry>,
    fonts: Vec<FontInfo>,
    fills: Vec<FillInfo>,
    borders: Vec<BorderInfo>,
}

impl StyleTables {
    fn number_format_for_style_id(&self, style_id: u32) -> Option<&str> {
        if style_id == 0 {
            return None;
        }
        let xf = self.cell_xfs.get(style_id as usize)?;
        if xf.num_fmt_id == 0 {
            return None;
        }
        if let Some(custom) = self.custom_num_fmts.get(&xf.num_fmt_id) {
            return Some(custom.as_str());
        }
        match builtin_num_fmt(xf.num_fmt_id) {
            Some("General") => None,
            other => other,
        }
    }

    fn border_for_style_id(&self, style_id: u32) -> Option<&BorderInfo> {
        let xf = self.cell_xfs.get(style_id as usize)?;
        self.borders.get(xf.border_id as usize)
    }

    fn font_for_style_id(&self, style_id: u32) -> Option<&FontInfo> {
        let xf = self.cell_xfs.get(style_id as usize)?;
        self.fonts.get(xf.font_id as usize)
    }

    fn fill_for_style_id(&self, style_id: u32) -> Option<&FillInfo> {
        let xf = self.cell_xfs.get(style_id as usize)?;
        self.fills.get(xf.fill_id as usize)
    }

    fn alignment_for_style_id(&self, style_id: u32) -> Option<&AlignmentInfo> {
        self.cell_xfs.get(style_id as usize)?.alignment.as_ref()
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct XfEntry {
    num_fmt_id: u32,
    font_id: u32,
    fill_id: u32,
    border_id: u32,
    alignment: Option<AlignmentInfo>,
}

/// Parsed cell font.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct FontInfo {
    pub bold: bool,
    pub italic: bool,
    pub underline: Option<String>,
    pub strikethrough: bool,
    pub name: Option<String>,
    pub size: Option<f64>,
    pub color: Option<String>,
}

/// Parsed cell fill.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FillInfo {
    pub bg_color: Option<String>,
}

/// Parsed cell alignment.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct AlignmentInfo {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: bool,
    pub text_rotation: Option<u32>,
    pub indent: Option<u32>,
}

/// Parsed cell border.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct BorderInfo {
    pub left: Option<BorderSide>,
    pub right: Option<BorderSide>,
    pub top: Option<BorderSide>,
    pub bottom: Option<BorderSide>,
    pub diagonal_up: Option<BorderSide>,
    pub diagonal_down: Option<BorderSide>,
}

/// Parsed border side.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BorderSide {
    pub style: String,
    pub color: String,
}

fn parse_workbook(
    xml: &str,
) -> Result<(
    Vec<SheetRef>,
    bool,
    Vec<NamedRange>,
    HashMap<String, String>,
    HashMap<String, PrintTitlesInfo>,
    WorkbookSecurity,
    Option<WorkbookPropertiesInfo>,
    Option<CalcPropertiesInfo>,
)> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut sheets = Vec::new();
    let mut raw_names = Vec::new();
    let mut raw_print_areas = Vec::new();
    let mut raw_print_titles = Vec::new();
    let mut date1904 = false;
    let mut workbook_protection = None;
    let mut file_sharing = None;
    let mut workbook_properties = None;
    let mut calc_properties = None;
    let mut in_defined_name = false;
    let mut current_name: Option<String> = None;
    let mut current_local_id: Option<usize> = None;
    let mut current_name_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"workbookPr" => {
                    date1904 = attr_truthy(attr_value(&e, b"date1904").as_deref());
                    workbook_properties = Some(WorkbookPropertiesInfo::from_start(&e));
                }
                b"calcPr" => {
                    calc_properties = Some(CalcPropertiesInfo::from_start(&e));
                }
                b"workbookProtection" => {
                    workbook_protection = Some(WorkbookProtection::from_start(&e));
                }
                b"fileSharing" => {
                    file_sharing = Some(FileSharing::from_start(&e));
                }
                b"sheet" => {
                    let name = attr_value(&e, b"name");
                    let rid = attr_value(&e, b"r:id");
                    if let (Some(name), Some(rid)) = (name, rid) {
                        sheets.push(SheetRef {
                            name,
                            sheet_id: attr_value(&e, b"sheetId"),
                            state: parse_sheet_state(attr_value(&e, b"state").as_deref()),
                            rid,
                        });
                    }
                }
                b"definedName" => {
                    in_defined_name = true;
                    current_name = attr_value(&e, b"name");
                    current_local_id =
                        attr_value(&e, b"localSheetId").and_then(|v| v.parse::<usize>().ok());
                    current_name_text.clear();
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_defined_name {
                    current_name_text
                        .push_str(&e.unescape().map_err(|err| {
                            ReaderError::Xml(format!("defined name text: {err}"))
                        })?);
                }
            }
            Ok(Event::End(e)) => {
                if e.local_name().as_ref() == b"definedName" {
                    in_defined_name = false;
                    if let Some(name) = current_name.take() {
                        let refers_to = current_name_text.trim().to_string();
                        if name == "_xlnm.Print_Area" && !refers_to.is_empty() {
                            raw_print_areas.push(RawNamedRange {
                                name,
                                local_id: current_local_id.take(),
                                refers_to,
                            });
                        } else if name == "_xlnm.Print_Titles" && !refers_to.is_empty() {
                            raw_print_titles.push(RawNamedRange {
                                name,
                                local_id: current_local_id.take(),
                                refers_to,
                            });
                        } else if !name.starts_with("_xlnm.") && !refers_to.is_empty() {
                            raw_names.push(RawNamedRange {
                                name,
                                local_id: current_local_id.take(),
                                refers_to,
                            });
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ReaderError::Xml(format!(
                    "failed to parse xl/workbook.xml: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    let named_ranges = resolve_named_ranges(&sheets, raw_names);
    let print_areas = resolve_print_areas(&sheets, raw_print_areas);
    let print_titles = resolve_print_titles(&sheets, raw_print_titles);
    Ok((
        sheets,
        date1904,
        named_ranges,
        print_areas,
        print_titles,
        WorkbookSecurity {
            workbook_protection,
            file_sharing,
        },
        workbook_properties,
        calc_properties,
    ))
}

#[derive(Debug)]
struct RawNamedRange {
    name: String,
    local_id: Option<usize>,
    refers_to: String,
}

impl WorkbookProtection {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            lock_structure: attr_bool_default(e, b"lockStructure", false),
            lock_windows: attr_bool_default(e, b"lockWindows", false),
            lock_revision: attr_bool_default(e, b"lockRevision", false),
            workbook_algorithm_name: attr_value(e, b"workbookAlgorithmName"),
            workbook_hash_value: attr_value(e, b"workbookHashValue"),
            workbook_salt_value: attr_value(e, b"workbookSaltValue"),
            workbook_spin_count: attr_value(e, b"workbookSpinCount")
                .and_then(|value| value.parse::<u32>().ok()),
            revisions_algorithm_name: attr_value(e, b"revisionsAlgorithmName"),
            revisions_hash_value: attr_value(e, b"revisionsHashValue"),
            revisions_salt_value: attr_value(e, b"revisionsSaltValue"),
            revisions_spin_count: attr_value(e, b"revisionsSpinCount")
                .and_then(|value| value.parse::<u32>().ok()),
        }
    }
}

impl WorkbookPropertiesInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            date1904: attr_bool_default(e, b"date1904", false),
            date_compatibility: attr_bool(e, b"dateCompatibility"),
            show_objects: attr_value(e, b"showObjects"),
            show_border_unselected_tables: attr_bool(e, b"showBorderUnselectedTables"),
            filter_privacy: attr_bool(e, b"filterPrivacy"),
            prompted_solutions: attr_bool(e, b"promptedSolutions"),
            show_ink_annotation: attr_bool(e, b"showInkAnnotation"),
            backup_file: attr_bool(e, b"backupFile"),
            save_external_link_values: attr_bool(e, b"saveExternalLinkValues"),
            update_links: attr_value(e, b"updateLinks"),
            code_name: attr_value(e, b"codeName"),
            hide_pivot_field_list: attr_bool(e, b"hidePivotFieldList"),
            show_pivot_chart_filter: attr_bool(e, b"showPivotChartFilter"),
            allow_refresh_query: attr_bool(e, b"allowRefreshQuery"),
            publish_items: attr_bool(e, b"publishItems"),
            check_compatibility: attr_bool(e, b"checkCompatibility"),
            auto_compress_pictures: attr_bool(e, b"autoCompressPictures"),
            refresh_all_connections: attr_bool(e, b"refreshAllConnections"),
            default_theme_version: attr_u32(e, b"defaultThemeVersion"),
        }
    }
}

impl CalcPropertiesInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            calc_id: attr_u32(e, b"calcId"),
            calc_mode: attr_value(e, b"calcMode"),
            full_calc_on_load: attr_bool(e, b"fullCalcOnLoad"),
            ref_mode: attr_value(e, b"refMode"),
            iterate: attr_bool(e, b"iterate"),
            iterate_count: attr_u32(e, b"iterateCount"),
            iterate_delta: attr_f64(e, b"iterateDelta"),
            full_precision: attr_bool(e, b"fullPrecision"),
            calc_completed: attr_bool(e, b"calcCompleted"),
            calc_on_save: attr_bool(e, b"calcOnSave"),
            concurrent_calc: attr_bool(e, b"concurrentCalc"),
            concurrent_manual_count: attr_u32(e, b"concurrentManualCount"),
            force_full_calc: attr_bool(e, b"forceFullCalc"),
        }
    }
}

impl FileSharing {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            read_only_recommended: attr_bool_default(e, b"readOnlyRecommended", false),
            user_name: attr_value(e, b"userName"),
            algorithm_name: attr_value(e, b"algorithmName"),
            hash_value: attr_value(e, b"hashValue"),
            salt_value: attr_value(e, b"saltValue"),
            spin_count: attr_value(e, b"spinCount").and_then(|value| value.parse::<u32>().ok()),
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq)]
struct SharedStrings {
    values: Vec<String>,
    rich_text: Vec<Option<Vec<RichTextRun>>>,
}

fn resolve_named_ranges(sheet_refs: &[SheetRef], raw_names: Vec<RawNamedRange>) -> Vec<NamedRange> {
    raw_names
        .into_iter()
        .map(|raw| {
            let (scope, sheet_name) = match raw.local_id {
                Some(index) => (
                    "sheet".to_string(),
                    sheet_refs.get(index).map(|sheet| sheet.name.clone()),
                ),
                None => ("workbook".to_string(), None),
            };
            let refers_to = if scope == "sheet" && !raw.refers_to.contains('!') {
                if let Some(sheet_name) = sheet_name {
                    format!("{sheet_name}!{}", raw.refers_to)
                } else {
                    raw.refers_to
                }
            } else {
                raw.refers_to
            };
            NamedRange {
                name: raw.name,
                scope,
                refers_to,
            }
        })
        .collect()
}

fn resolve_print_areas(
    sheet_refs: &[SheetRef],
    raw_print_areas: Vec<RawNamedRange>,
) -> HashMap<String, String> {
    let mut out = HashMap::new();
    for raw in raw_print_areas {
        let sheet_name = raw
            .local_id
            .and_then(|index| sheet_refs.get(index).map(|sheet| sheet.name.clone()))
            .or_else(|| sheet_name_from_ref(&raw.refers_to));
        let Some(sheet_name) = sheet_name else {
            continue;
        };
        let refers_to = if raw.refers_to.contains('!') {
            raw.refers_to
        } else {
            format!("{sheet_name}!{}", raw.refers_to)
        };
        out.insert(sheet_name, refers_to);
    }
    out
}

fn resolve_print_titles(
    sheet_refs: &[SheetRef],
    raw_print_titles: Vec<RawNamedRange>,
) -> HashMap<String, PrintTitlesInfo> {
    let mut out = HashMap::new();
    for raw in raw_print_titles {
        let sheet_name = raw
            .local_id
            .and_then(|index| sheet_refs.get(index).map(|sheet| sheet.name.clone()))
            .or_else(|| sheet_name_from_ref(&raw.refers_to));
        let Some(sheet_name) = sheet_name else {
            continue;
        };
        let titles = parse_print_titles_ref(&raw.refers_to);
        if titles.rows.is_some() || titles.cols.is_some() {
            out.insert(sheet_name, titles);
        }
    }
    out
}

fn parse_print_titles_ref(refers_to: &str) -> PrintTitlesInfo {
    let mut info = PrintTitlesInfo::default();
    for part in refers_to.split(',') {
        let range = part
            .split_once('!')
            .map(|(_, range)| range)
            .unwrap_or(part)
            .replace('$', "");
        let Some((start, end)) = range.split_once(':') else {
            continue;
        };
        if start.chars().all(|ch| ch.is_ascii_digit()) && end.chars().all(|ch| ch.is_ascii_digit())
        {
            info.rows = Some(format!("{start}:{end}"));
        } else if start.chars().all(|ch| ch.is_ascii_alphabetic())
            && end.chars().all(|ch| ch.is_ascii_alphabetic())
        {
            info.cols = Some(format!(
                "{}:{}",
                start.to_ascii_uppercase(),
                end.to_ascii_uppercase()
            ));
        }
    }
    info
}

fn sheet_name_from_ref(refers_to: &str) -> Option<String> {
    let (sheet, _) = refers_to.split_once('!')?;
    let unquoted = sheet
        .strip_prefix('\'')
        .and_then(|value| value.strip_suffix('\''))
        .unwrap_or(sheet);
    Some(unquoted.replace("''", "'"))
}

fn resolve_sheet_paths(sheet_refs: Vec<SheetRef>, rels: &RelsGraph) -> Result<Vec<SheetInfo>> {
    let mut sheets = Vec::with_capacity(sheet_refs.len());
    for sheet in sheet_refs {
        let rel = rels.get(&RelId(sheet.rid.clone())).ok_or_else(|| {
            ReaderError::MissingPart(format!("workbook relationship {}", sheet.rid))
        })?;
        sheets.push(SheetInfo {
            name: sheet.name,
            sheet_id: sheet.sheet_id,
            state: sheet.state,
            path: join_and_normalize("xl/", &rel.target),
        });
    }
    Ok(sheets)
}

fn parse_shared_strings(xml: &str) -> Result<SharedStrings> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = SharedStrings::default();
    let mut current = String::new();
    let mut runs: Vec<RichTextRun> = Vec::new();
    let mut current_run: Option<RichTextRun> = None;
    let mut current_props: Option<InlineFontProps> = None;
    let mut in_si = false;
    let mut in_t = false;
    let mut saw_r = false;
    let mut text_buf = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"si" => {
                    in_si = true;
                    current.clear();
                    runs.clear();
                    current_run = None;
                    current_props = None;
                    saw_r = false;
                    text_buf.clear();
                }
                b"r" if in_si => {
                    saw_r = true;
                    current_run = Some(RichTextRun::default());
                    current_props = None;
                }
                b"rPr" if in_si && current_run.is_some() => {
                    current_props = Some(InlineFontProps::default());
                }
                b"t" => {
                    if in_si {
                        in_t = true;
                        text_buf.clear();
                    }
                }
                other if current_props.is_some() => {
                    apply_rich_font_attr(current_props.as_mut().unwrap(), other, e.attributes());
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"si" => {
                    out.values.push(String::new());
                    out.rich_text.push(None);
                }
                other if current_props.is_some() => {
                    apply_rich_font_attr(current_props.as_mut().unwrap(), other, e.attributes());
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"si" => {
                    out.values.push(std::mem::take(&mut current));
                    out.rich_text.push(saw_r.then(|| std::mem::take(&mut runs)));
                    in_si = false;
                }
                b"t" => {
                    current.push_str(&text_buf);
                    if let Some(run) = current_run.as_mut() {
                        run.text.push_str(&text_buf);
                    }
                    text_buf.clear();
                    in_t = false;
                }
                b"rPr" => {
                    if let (Some(run), Some(props)) = (current_run.as_mut(), current_props.take()) {
                        run.font = Some(props);
                    }
                }
                b"r" => {
                    if let Some(run) = current_run.take() {
                        runs.push(run);
                    }
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_si && in_t {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("shared string text: {err}")))?;
                    text_buf.push_str(&normalize_ooxml_text(&text));
                }
            }
            Ok(Event::CData(e)) => {
                if in_si && in_t {
                    text_buf.push_str(&normalize_ooxml_text(&String::from_utf8_lossy(e.as_ref())));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ReaderError::Xml(format!(
                    "failed to parse xl/sharedStrings.xml: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn apply_rich_font_attr(
    props: &mut InlineFontProps,
    tag: &[u8],
    mut attrs: quick_xml::events::attributes::Attributes<'_>,
) {
    let mut all_attrs: Vec<(Vec<u8>, String)> = Vec::with_capacity(2);
    for a in attrs.with_checks(false) {
        if let Ok(Attribute { key, value }) = a {
            all_attrs.push((
                key.local_name().as_ref().to_vec(),
                String::from_utf8_lossy(value.as_ref()).into_owned(),
            ));
        }
    }
    let val = || -> Option<String> {
        all_attrs
            .iter()
            .rev()
            .find(|(k, _)| k.as_slice() == b"val")
            .map(|(_, v)| v.clone())
    };
    match tag {
        b"b" => props.bold = Some(parse_rich_bool(val())),
        b"i" => props.italic = Some(parse_rich_bool(val())),
        b"strike" => props.strike = Some(parse_rich_bool(val())),
        b"u" => props.underline = Some(val().unwrap_or_else(|| "single".to_string())),
        b"sz" => props.size = val().and_then(|v| v.parse::<f64>().ok()),
        b"rFont" => props.name = val(),
        b"family" => props.family = val().and_then(|v| v.parse::<i32>().ok()),
        b"charset" => props.charset = val().and_then(|v| v.parse::<i32>().ok()),
        b"vertAlign" => props.vert_align = val(),
        b"scheme" => props.scheme = val(),
        b"color" => {
            let mut rgb = None;
            let mut theme = None;
            let mut indexed = None;
            for (k, v) in &all_attrs {
                match k.as_slice() {
                    b"rgb" => rgb = Some(v.clone()),
                    b"theme" => theme = Some(v.clone()),
                    b"indexed" => indexed = Some(v.clone()),
                    _ => {}
                }
            }
            props.color = rgb
                .or_else(|| theme.map(|v| format!("theme:{v}")))
                .or_else(|| indexed.map(|v| format!("indexed:{v}")));
        }
        _ => {}
    }
}

fn parse_rich_bool(value: Option<String>) -> bool {
    match value.as_deref() {
        None => true,
        Some(raw) => {
            let trimmed = raw.trim();
            !(trimmed == "0" || trimmed.eq_ignore_ascii_case("false"))
        }
    }
}

fn parse_style_tables(xml: &str) -> Result<StyleTables> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut styles = StyleTables::default();
    let mut in_num_fmts = false;
    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_cell_xfs = false;
    let mut in_borders = false;
    let mut current_font: Option<FontBuilder> = None;
    let mut current_fill: Option<FillBuilder> = None;
    let mut in_pattern_fill = false;
    let mut current_border: Option<BorderBuilder> = None;
    let mut current_border_edge: Option<BorderEdgeKind> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"numFmts" => in_num_fmts = true,
                b"fonts" => in_fonts = true,
                b"fills" => in_fills = true,
                b"cellXfs" => in_cell_xfs = true,
                b"borders" => in_borders = true,
                b"font" if in_fonts => {
                    current_font = Some(FontBuilder::default());
                }
                tag if in_fonts && current_font.is_some() => {
                    if let Some(font) = current_font.as_mut() {
                        font.apply_tag(tag, &e);
                    }
                }
                b"fill" if in_fills => {
                    current_fill = Some(FillBuilder::default());
                }
                b"patternFill" if in_fills && current_fill.is_some() => {
                    in_pattern_fill = true;
                    if let Some(fill) = current_fill.as_mut() {
                        fill.pattern_type = attr_value(&e, b"patternType");
                    }
                }
                tag if in_pattern_fill && current_fill.is_some() => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.apply_color_tag(tag, &e);
                    }
                }
                b"border" if in_borders => {
                    current_border = Some(BorderBuilder::from_start(&e));
                }
                tag if in_borders && current_border.is_some() => {
                    if let Some(edge) = BorderEdgeKind::from_tag(tag) {
                        current_border_edge = Some(edge);
                        if let Some(border) = current_border.as_mut() {
                            border.set_edge(edge, parse_border_side(&e));
                        }
                    } else if tag == b"color" {
                        if let (Some(border), Some(edge)) =
                            (current_border.as_mut(), current_border_edge)
                        {
                            border.update_edge_color(edge, parse_border_color(&e));
                        }
                    }
                }
                b"numFmt" if in_num_fmts => {
                    push_num_fmt(&mut styles, &e);
                }
                b"xf" if in_cell_xfs => {
                    styles.cell_xfs.push(parse_xf_entry(&e));
                }
                b"alignment" if in_cell_xfs => {
                    if let Some(xf) = styles.cell_xfs.last_mut() {
                        xf.alignment = parse_alignment(&e);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"numFmt" if in_num_fmts => push_num_fmt(&mut styles, &e),
                b"xf" if in_cell_xfs => styles.cell_xfs.push(parse_xf_entry(&e)),
                b"font" if in_fonts => styles.fonts.push(FontInfo::default()),
                tag if in_fonts && current_font.is_some() => {
                    if let Some(font) = current_font.as_mut() {
                        font.apply_tag(tag, &e);
                    }
                }
                b"fill" if in_fills => styles.fills.push(FillInfo::default()),
                b"patternFill" if in_fills && current_fill.is_some() => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.pattern_type = attr_value(&e, b"patternType");
                    }
                }
                tag if in_pattern_fill && current_fill.is_some() => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.apply_color_tag(tag, &e);
                    }
                }
                b"alignment" if in_cell_xfs => {
                    if let Some(xf) = styles.cell_xfs.last_mut() {
                        xf.alignment = parse_alignment(&e);
                    }
                }
                b"border" if in_borders => {
                    styles.borders.push(BorderBuilder::from_start(&e).finish());
                }
                tag if in_borders && current_border.is_some() => {
                    if let Some(edge) = BorderEdgeKind::from_tag(tag) {
                        if let Some(border) = current_border.as_mut() {
                            border.set_edge(edge, parse_border_side(&e));
                        }
                    } else if tag == b"color" {
                        if let (Some(border), Some(edge)) =
                            (current_border.as_mut(), current_border_edge)
                        {
                            border.update_edge_color(edge, parse_border_color(&e));
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"numFmts" => in_num_fmts = false,
                b"fonts" => in_fonts = false,
                b"fills" => in_fills = false,
                b"cellXfs" => in_cell_xfs = false,
                b"borders" => in_borders = false,
                b"font" => {
                    if let Some(font) = current_font.take() {
                        styles.fonts.push(font.finish());
                    }
                }
                b"fill" => {
                    if let Some(fill) = current_fill.take() {
                        styles.fills.push(fill.finish());
                    }
                    in_pattern_fill = false;
                }
                b"patternFill" => in_pattern_fill = false,
                b"border" => {
                    if let Some(border) = current_border.take() {
                        styles.borders.push(border.finish());
                    }
                    current_border_edge = None;
                }
                tag if BorderEdgeKind::from_tag(tag).is_some() => {
                    current_border_edge = None;
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ReaderError::Xml(format!(
                    "failed to parse xl/styles.xml: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(styles)
}

fn push_num_fmt(styles: &mut StyleTables, e: &BytesStart<'_>) {
    let id = attr_value(e, b"numFmtId").and_then(|value| value.parse::<u32>().ok());
    let code = attr_value(e, b"formatCode");
    if let (Some(id), Some(code)) = (id, code) {
        styles.custom_num_fmts.insert(id, code);
    }
}

fn parse_xf_entry(e: &BytesStart<'_>) -> XfEntry {
    XfEntry {
        num_fmt_id: attr_value(e, b"numFmtId")
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(0),
        font_id: attr_value(e, b"fontId")
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(0),
        fill_id: attr_value(e, b"fillId")
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(0),
        border_id: attr_value(e, b"borderId")
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(0),
        alignment: None,
    }
}

#[derive(Debug, Default)]
struct FontBuilder {
    font: FontInfo,
}

impl FontBuilder {
    fn apply_tag(&mut self, tag: &[u8], e: &BytesStart<'_>) {
        match tag {
            b"b" => self.font.bold = parse_bool_attr_or_true(e, b"val"),
            b"i" => self.font.italic = parse_bool_attr_or_true(e, b"val"),
            b"strike" => self.font.strikethrough = parse_bool_attr_or_true(e, b"val"),
            b"u" => {
                self.font.underline =
                    Some(attr_value(e, b"val").unwrap_or_else(|| "single".to_string()));
            }
            b"name" => self.font.name = attr_value(e, b"val"),
            b"sz" => {
                self.font.size = attr_value(e, b"val").and_then(|value| value.parse().ok());
            }
            b"color" => self.font.color = parse_ooxml_color(e),
            _ => {}
        }
    }

    fn finish(self) -> FontInfo {
        self.font
    }
}

#[derive(Debug, Default)]
struct FillBuilder {
    pattern_type: Option<String>,
    fg_color: Option<String>,
    bg_color: Option<String>,
}

impl FillBuilder {
    fn apply_color_tag(&mut self, tag: &[u8], e: &BytesStart<'_>) {
        match tag {
            b"fgColor" => self.fg_color = parse_ooxml_color(e),
            b"bgColor" => self.bg_color = parse_ooxml_color(e),
            _ => {}
        }
    }

    fn finish(self) -> FillInfo {
        let is_none = self
            .pattern_type
            .as_deref()
            .is_none_or(|pattern| pattern.eq_ignore_ascii_case("none"));
        FillInfo {
            bg_color: (!is_none)
                .then(|| self.fg_color.or(self.bg_color))
                .flatten(),
        }
    }
}

fn parse_alignment(e: &BytesStart<'_>) -> Option<AlignmentInfo> {
    let horizontal = attr_value(e, b"horizontal").filter(|value| value != "general");
    let vertical = attr_value(e, b"vertical");
    let wrap_text = attr_truthy(attr_value(e, b"wrapText").as_deref());
    let text_rotation = attr_value(e, b"textRotation").and_then(|value| value.parse().ok());
    let indent = attr_value(e, b"indent").and_then(|value| value.parse().ok());
    let alignment = AlignmentInfo {
        horizontal,
        vertical,
        wrap_text,
        text_rotation,
        indent,
    };
    (alignment.horizontal.is_some()
        || alignment.vertical.is_some()
        || alignment.wrap_text
        || alignment
            .text_rotation
            .is_some_and(|rotation| rotation != 0)
        || alignment.indent.is_some_and(|indent| indent != 0))
    .then_some(alignment)
}

fn parse_bool_attr_or_true(e: &BytesStart<'_>, key: &[u8]) -> bool {
    attr_value(e, key)
        .map(|value| attr_truthy(Some(&value)))
        .unwrap_or(true)
}

fn parse_ooxml_color(e: &BytesStart<'_>) -> Option<String> {
    if let Some(rgb) = attr_value(e, b"rgb") {
        return Some(normalize_ooxml_rgb(&rgb));
    }
    if let Some(indexed) = attr_value(e, b"indexed").and_then(|value| value.parse::<usize>().ok()) {
        return Some(indexed_color_hex(indexed));
    }
    attr_value(e, b"theme").map(|_| "#000000".to_string())
}

#[derive(Debug, Clone, Copy)]
enum BorderEdgeKind {
    Left,
    Right,
    Top,
    Bottom,
    Diagonal,
}

impl BorderEdgeKind {
    fn from_tag(tag: &[u8]) -> Option<Self> {
        match tag {
            b"left" => Some(Self::Left),
            b"right" => Some(Self::Right),
            b"top" => Some(Self::Top),
            b"bottom" => Some(Self::Bottom),
            b"diagonal" => Some(Self::Diagonal),
            _ => None,
        }
    }
}

#[derive(Debug, Default)]
struct BorderBuilder {
    border: BorderInfo,
    diagonal_up: bool,
    diagonal_down: bool,
}

impl BorderBuilder {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            border: BorderInfo::default(),
            diagonal_up: attr_truthy(attr_value(e, b"diagonalUp").as_deref()),
            diagonal_down: attr_truthy(attr_value(e, b"diagonalDown").as_deref()),
        }
    }

    fn set_edge(&mut self, edge: BorderEdgeKind, side: Option<BorderSide>) {
        match edge {
            BorderEdgeKind::Left => self.border.left = side,
            BorderEdgeKind::Right => self.border.right = side,
            BorderEdgeKind::Top => self.border.top = side,
            BorderEdgeKind::Bottom => self.border.bottom = side,
            BorderEdgeKind::Diagonal => {
                if self.diagonal_up {
                    self.border.diagonal_up = side.clone();
                }
                if self.diagonal_down {
                    self.border.diagonal_down = side;
                }
            }
        }
    }

    fn update_edge_color(&mut self, edge: BorderEdgeKind, color: Option<String>) {
        let Some(color) = color else {
            return;
        };
        match edge {
            BorderEdgeKind::Left => {
                if let Some(side) = self.border.left.as_mut() {
                    side.color = color;
                }
            }
            BorderEdgeKind::Right => {
                if let Some(side) = self.border.right.as_mut() {
                    side.color = color;
                }
            }
            BorderEdgeKind::Top => {
                if let Some(side) = self.border.top.as_mut() {
                    side.color = color;
                }
            }
            BorderEdgeKind::Bottom => {
                if let Some(side) = self.border.bottom.as_mut() {
                    side.color = color;
                }
            }
            BorderEdgeKind::Diagonal => {
                if let Some(side) = self.border.diagonal_up.as_mut() {
                    side.color = color.clone();
                }
                if let Some(side) = self.border.diagonal_down.as_mut() {
                    side.color = color;
                }
            }
        }
    }

    fn finish(self) -> BorderInfo {
        self.border
    }
}

fn parse_border_side(e: &BytesStart<'_>) -> Option<BorderSide> {
    let style = attr_value(e, b"style").filter(|value| !value.is_empty() && value != "none")?;
    let color = parse_border_color(e).unwrap_or_else(|| "#000000".to_string());
    Some(BorderSide { style, color })
}

fn parse_border_color(e: &BytesStart<'_>) -> Option<String> {
    parse_ooxml_color(e)
}

fn normalize_ooxml_rgb(rgb: &str) -> String {
    let raw = rgb.trim().trim_start_matches('#');
    let rgb = if raw.len() == 8 { &raw[2..] } else { raw };
    format!("#{}", rgb.to_ascii_uppercase())
}

fn indexed_color_hex(index: usize) -> String {
    const COLORS: [&str; 66] = [
        "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
        "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
        "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",
        "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",
        "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",
        "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
        "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696",
        "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
        "#000000", "#FFFFFF",
    ];
    COLORS.get(index).unwrap_or(&"#000000").to_string()
}

fn builtin_num_fmt(format_id: u32) -> Option<&'static str> {
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

fn parse_worksheet(
    xml: &str,
    shared_strings: &SharedStrings,
    rels: Option<&RelsGraph>,
    comments: Vec<Comment>,
    tables: Vec<Table>,
) -> Result<WorksheetData> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut dimension = None;
    let mut merged_ranges = Vec::new();
    let mut hyperlink_nodes = Vec::new();
    let mut freeze_panes = None;
    let mut sheet_properties = None;
    let mut sheet_view = None;
    let mut current_sheet_view: Option<SheetViewInfo> = None;
    let mut row_heights = HashMap::new();
    let mut column_widths = Vec::new();
    let mut data_validations = Vec::new();
    let mut sheet_protection = None;
    let mut page_margins = None;
    let mut page_setup = None;
    let mut header_footer = None;
    let mut row_breaks = None;
    let mut column_breaks = None;
    let mut sheet_format = None;
    let mut conditional_formats = Vec::new();
    let mut hidden_rows: HashMap<u32, bool> = HashMap::new();
    let mut hidden_columns: HashMap<u32, bool> = HashMap::new();
    let mut row_outline_levels: HashMap<u32, u8> = HashMap::new();
    let mut column_outline_levels: HashMap<u32, u8> = HashMap::new();
    let mut current_conditional_range: Option<String> = None;
    let mut current_conditional_rule: Option<ConditionalFormatBuilder> = None;
    let mut in_conditional_formula = false;
    let mut row_index: Option<u32> = None;
    let mut current: Option<CellBuilder> = None;
    let mut active_text: Option<TextTarget> = None;
    let mut current_validation: Option<DataValidationBuilder> = None;
    let mut active_validation_text: Option<DataValidationFormula> = None;
    let mut current_header_footer: Option<HeaderFooterInfo> = None;
    let mut active_header_footer: Option<HeaderFooterPart> = None;
    let mut active_breaks: Option<BreakListKind> = None;
    let mut cells = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"sheetPr" => {
                    sheet_properties = Some(SheetPropertiesInfo::from_start(&e));
                }
                b"tabColor" => {
                    sheet_properties
                        .get_or_insert_with(SheetPropertiesInfo::default)
                        .apply_tab_color(&e);
                }
                b"outlinePr" => {
                    sheet_properties
                        .get_or_insert_with(SheetPropertiesInfo::default)
                        .apply_outline(&e);
                }
                b"pageSetUpPr" => {
                    sheet_properties
                        .get_or_insert_with(SheetPropertiesInfo::default)
                        .apply_page_setup(&e);
                }
                b"dimension" => {
                    dimension = attr_value(&e, b"ref");
                }
                b"mergeCell" => {
                    if let Some(range) = attr_value(&e, b"ref") {
                        merged_ranges.push(range);
                    }
                }
                b"hyperlink" => {
                    if let Some(node) = HyperlinkNode::from_start(&e) {
                        hyperlink_nodes.push(node);
                    }
                }
                b"pane" => {
                    if let Some(pane) = parse_pane(&e) {
                        freeze_panes = Some(pane.clone());
                        if let Some(view) = current_sheet_view.as_mut() {
                            view.pane = Some(pane);
                        }
                    }
                }
                b"sheetView" => {
                    if current_sheet_view.is_none() && sheet_view.is_none() {
                        current_sheet_view = Some(SheetViewInfo::from_start(&e));
                    }
                }
                b"selection" => {
                    if let Some(view) = current_sheet_view.as_mut() {
                        view.selections.push(SelectionInfo::from_start(&e));
                    }
                }
                b"sheetFormatPr" => {
                    sheet_format = Some(SheetFormatInfo::from_start(&e));
                }
                b"row" => {
                    row_index = attr_value(&e, b"r").and_then(|v| v.parse::<u32>().ok());
                    if let Some((row, height)) = parse_row_height(&e, row_index) {
                        row_heights.insert(row, height);
                    }
                    update_row_visibility(&e, row_index, &mut hidden_rows, &mut row_outline_levels);
                }
                b"c" => {
                    current = Some(CellBuilder::from_start(&e, row_index));
                }
                b"dataValidation" => {
                    current_validation = Some(DataValidationBuilder::from_start(&e));
                }
                b"sheetProtection" => {
                    sheet_protection = Some(SheetProtection::from_start(&e));
                }
                b"pageMargins" => {
                    page_margins = parse_page_margins(&e);
                }
                b"pageSetup" => {
                    page_setup = Some(PageSetupInfo::from_start(&e));
                }
                b"headerFooter" => {
                    current_header_footer = Some(HeaderFooterInfo::from_start(&e));
                }
                b"oddHeader" => active_header_footer = Some(HeaderFooterPart::OddHeader),
                b"oddFooter" => active_header_footer = Some(HeaderFooterPart::OddFooter),
                b"evenHeader" => active_header_footer = Some(HeaderFooterPart::EvenHeader),
                b"evenFooter" => active_header_footer = Some(HeaderFooterPart::EvenFooter),
                b"firstHeader" => active_header_footer = Some(HeaderFooterPart::FirstHeader),
                b"firstFooter" => active_header_footer = Some(HeaderFooterPart::FirstFooter),
                b"rowBreaks" => {
                    row_breaks = Some(PageBreakListInfo::from_start(&e));
                    active_breaks = Some(BreakListKind::Row);
                }
                b"colBreaks" => {
                    column_breaks = Some(PageBreakListInfo::from_start(&e));
                    active_breaks = Some(BreakListKind::Column);
                }
                b"brk" => append_break(&mut row_breaks, &mut column_breaks, active_breaks, &e),
                b"conditionalFormatting" => {
                    current_conditional_range = attr_value(&e, b"sqref");
                }
                b"cfRule" => {
                    current_conditional_rule = Some(ConditionalFormatBuilder::from_start(
                        &e,
                        current_conditional_range.clone().unwrap_or_default(),
                    ));
                }
                b"formula" => {
                    if current_conditional_rule.is_some() {
                        in_conditional_formula = true;
                    }
                }
                b"formula1" => {
                    if current_validation.is_some() {
                        active_validation_text = Some(DataValidationFormula::Formula1);
                    }
                }
                b"formula2" => {
                    if current_validation.is_some() {
                        active_validation_text = Some(DataValidationFormula::Formula2);
                    }
                }
                b"v" => active_text = Some(TextTarget::Value),
                b"f" => {
                    if let Some(cell) = current.as_mut() {
                        cell.start_formula(&e);
                    }
                    active_text = Some(TextTarget::Formula);
                }
                b"t" => {
                    if current
                        .as_ref()
                        .is_some_and(|c| c.data_type == CellDataType::InlineString)
                    {
                        active_text = Some(TextTarget::InlineString);
                    }
                }
                b"r" => {
                    if let Some(cell) = current
                        .as_mut()
                        .filter(|c| c.data_type == CellDataType::InlineString)
                    {
                        cell.start_inline_run();
                    }
                }
                b"rPr" => {
                    if let Some(cell) = current
                        .as_mut()
                        .filter(|c| c.data_type == CellDataType::InlineString)
                    {
                        cell.start_inline_props();
                    }
                }
                other => {
                    if let Some(cell) = current
                        .as_mut()
                        .filter(|c| c.data_type == CellDataType::InlineString)
                    {
                        cell.apply_inline_font_tag(other, e.attributes());
                    }
                }
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"sheetPr" => {
                    sheet_properties = Some(SheetPropertiesInfo::from_start(&e));
                }
                b"tabColor" => {
                    sheet_properties
                        .get_or_insert_with(SheetPropertiesInfo::default)
                        .apply_tab_color(&e);
                }
                b"outlinePr" => {
                    sheet_properties
                        .get_or_insert_with(SheetPropertiesInfo::default)
                        .apply_outline(&e);
                }
                b"pageSetUpPr" => {
                    sheet_properties
                        .get_or_insert_with(SheetPropertiesInfo::default)
                        .apply_page_setup(&e);
                }
                b"col" => {
                    if let Some(width) = parse_column_width(&e) {
                        column_widths.push(width);
                    }
                    update_column_visibility(&e, &mut hidden_columns, &mut column_outline_levels);
                }
                b"dimension" => {
                    dimension = attr_value(&e, b"ref");
                }
                b"mergeCell" => {
                    if let Some(range) = attr_value(&e, b"ref") {
                        merged_ranges.push(range);
                    }
                }
                b"hyperlink" => {
                    if let Some(node) = HyperlinkNode::from_start(&e) {
                        hyperlink_nodes.push(node);
                    }
                }
                b"pane" => {
                    if let Some(pane) = parse_pane(&e) {
                        freeze_panes = Some(pane.clone());
                        if let Some(view) = current_sheet_view.as_mut() {
                            view.pane = Some(pane);
                        }
                    }
                }
                b"sheetView" => {
                    if sheet_view.is_none() {
                        sheet_view = Some(SheetViewInfo::from_start(&e));
                    }
                }
                b"selection" => {
                    if let Some(view) = current_sheet_view.as_mut() {
                        view.selections.push(SelectionInfo::from_start(&e));
                    }
                }
                b"sheetFormatPr" => {
                    sheet_format = Some(SheetFormatInfo::from_start(&e));
                }
                b"c" => {
                    let builder = CellBuilder::from_start(&e, row_index);
                    cells.push(builder.finish(shared_strings)?);
                }
                b"f" => {
                    if let Some(cell) = current.as_mut() {
                        cell.start_formula(&e);
                    }
                }
                b"row" => {
                    let row = attr_value(&e, b"r").and_then(|v| v.parse::<u32>().ok());
                    if let Some((row, height)) = parse_row_height(&e, row) {
                        row_heights.insert(row, height);
                    }
                    update_row_visibility(&e, row, &mut hidden_rows, &mut row_outline_levels);
                }
                b"dataValidation" => {
                    let validation = DataValidationBuilder::from_start(&e).finish();
                    if !validation.range.trim().is_empty() {
                        data_validations.push(validation);
                    }
                }
                b"sheetProtection" => {
                    sheet_protection = Some(SheetProtection::from_start(&e));
                }
                b"pageMargins" => {
                    page_margins = parse_page_margins(&e);
                }
                b"pageSetup" => {
                    page_setup = Some(PageSetupInfo::from_start(&e));
                }
                b"headerFooter" => {
                    header_footer = Some(HeaderFooterInfo::from_start(&e));
                }
                b"rowBreaks" => {
                    row_breaks = Some(PageBreakListInfo::from_start(&e));
                }
                b"colBreaks" => {
                    column_breaks = Some(PageBreakListInfo::from_start(&e));
                }
                b"brk" => append_break(&mut row_breaks, &mut column_breaks, active_breaks, &e),
                b"conditionalFormatting" => {
                    current_conditional_range = attr_value(&e, b"sqref");
                }
                b"cfRule" => {
                    let rule = ConditionalFormatBuilder::from_start(
                        &e,
                        current_conditional_range.clone().unwrap_or_default(),
                    )
                    .finish();
                    if !rule.range.trim().is_empty() && !rule.rule_type.trim().is_empty() {
                        conditional_formats.push(rule);
                    }
                }
                other => {
                    if let Some(cell) = current
                        .as_mut()
                        .filter(|c| c.data_type == CellDataType::InlineString)
                    {
                        cell.apply_inline_font_tag(other, e.attributes());
                    }
                }
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"c" => {
                    if let Some(builder) = current.take() {
                        cells.push(builder.finish(shared_strings)?);
                    }
                }
                b"formula1" => {
                    active_validation_text = None;
                    if let Some(validation) = current_validation.as_mut() {
                        validation.finish_formula1();
                    }
                }
                b"formula2" => {
                    active_validation_text = None;
                    if let Some(validation) = current_validation.as_mut() {
                        validation.finish_formula2();
                    }
                }
                b"dataValidation" => {
                    active_validation_text = None;
                    if let Some(validation) = current_validation.take() {
                        let validation = validation.finish();
                        if !validation.range.trim().is_empty() {
                            data_validations.push(validation);
                        }
                    }
                }
                b"oddHeader" | b"oddFooter" | b"evenHeader" | b"evenFooter" | b"firstHeader"
                | b"firstFooter" => {
                    active_header_footer = None;
                }
                b"headerFooter" => {
                    header_footer = current_header_footer.take();
                    active_header_footer = None;
                }
                b"rowBreaks" | b"colBreaks" => {
                    active_breaks = None;
                }
                b"sheetView" => {
                    if sheet_view.is_none() {
                        sheet_view = current_sheet_view.take();
                    } else {
                        current_sheet_view = None;
                    }
                }
                b"conditionalFormatting" => {
                    current_conditional_range = None;
                }
                b"formula" => {
                    if in_conditional_formula {
                        in_conditional_formula = false;
                        if let Some(rule) = current_conditional_rule.as_mut() {
                            rule.finish_formula();
                        }
                    } else {
                        active_text = None;
                    }
                }
                b"cfRule" => {
                    in_conditional_formula = false;
                    if let Some(rule) = current_conditional_rule.take() {
                        let rule = rule.finish();
                        if !rule.range.trim().is_empty() && !rule.rule_type.trim().is_empty() {
                            conditional_formats.push(rule);
                        }
                    }
                }
                b"v" | b"f" | b"t" => active_text = None,
                b"rPr" => {
                    if let Some(cell) = current.as_mut() {
                        cell.end_inline_props();
                    }
                }
                b"r" => {
                    if let Some(cell) = current.as_mut() {
                        cell.end_inline_run();
                    }
                }
                b"row" => row_index = None,
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_conditional_formula {
                    if let Some(rule) = current_conditional_rule.as_mut() {
                        let text = e
                            .unescape()
                            .map_err(|err| ReaderError::Xml(format!("worksheet text: {err}")))?;
                        rule.push_formula_text(&text);
                    }
                } else if let (Some(target), Some(validation)) =
                    (active_validation_text, current_validation.as_mut())
                {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("worksheet text: {err}")))?;
                    validation.push_text(target, &text);
                } else if let (Some(part), Some(header_footer)) =
                    (active_header_footer, current_header_footer.as_mut())
                {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("worksheet text: {err}")))?;
                    header_footer.set_part(part, parse_header_footer_item_text(&text));
                } else if let (Some(target), Some(cell)) = (active_text, current.as_mut()) {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("worksheet text: {err}")))?;
                    cell.push_text(target, &text);
                }
            }
            Ok(Event::CData(e)) => {
                if in_conditional_formula {
                    if let Some(rule) = current_conditional_rule.as_mut() {
                        rule.push_formula_text(&String::from_utf8_lossy(e.as_ref()));
                    }
                } else if let (Some(target), Some(validation)) =
                    (active_validation_text, current_validation.as_mut())
                {
                    validation.push_text(target, &String::from_utf8_lossy(e.as_ref()));
                } else if let (Some(part), Some(header_footer)) =
                    (active_header_footer, current_header_footer.as_mut())
                {
                    header_footer.set_part(
                        part,
                        parse_header_footer_item_text(&String::from_utf8_lossy(e.as_ref())),
                    );
                } else if let (Some(target), Some(cell)) = (active_text, current.as_mut()) {
                    cell.push_text(target, &String::from_utf8_lossy(e.as_ref()));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse worksheet: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    let mut hidden_rows: Vec<u32> = hidden_rows
        .into_iter()
        .filter_map(|(row, hidden)| hidden.then_some(row))
        .collect();
    let mut hidden_columns: Vec<u32> = hidden_columns
        .into_iter()
        .filter_map(|(col, hidden)| hidden.then_some(col))
        .collect();
    let mut row_outline_levels: Vec<(u32, u8)> = row_outline_levels.into_iter().collect();
    let mut column_outline_levels: Vec<(u32, u8)> = column_outline_levels.into_iter().collect();
    hidden_rows.sort_unstable();
    hidden_columns.sort_unstable();
    row_outline_levels.sort_unstable_by_key(|(row, _)| *row);
    column_outline_levels.sort_unstable_by_key(|(col, _)| *col);

    let auto_filter = parse_auto_filter(xml)?;
    resolve_shared_formulas(&mut cells);
    let array_formulas = build_array_formula_map(&cells);

    Ok(WorksheetData {
        dimension,
        merged_ranges,
        hyperlinks: resolve_hyperlinks(hyperlink_nodes, rels, &cells),
        freeze_panes,
        sheet_properties,
        sheet_view,
        comments,
        row_heights,
        column_widths,
        data_validations,
        sheet_protection,
        auto_filter,
        page_margins,
        page_setup,
        header_footer,
        row_breaks,
        column_breaks,
        sheet_format,
        images: Vec::new(),
        charts: Vec::new(),
        tables,
        conditional_formats,
        hidden_rows,
        hidden_columns,
        row_outline_levels,
        column_outline_levels,
        array_formulas,
        cells,
    })
}

fn parse_row_height(e: &BytesStart<'_>, row: Option<u32>) -> Option<(u32, RowHeight)> {
    let row = row?;
    let height = attr_value(e, b"ht")?.parse::<f64>().ok()?;
    Some((
        row,
        RowHeight {
            height,
            custom_height: attr_truthy(attr_value(e, b"customHeight").as_deref()),
        },
    ))
}

fn parse_column_width(e: &BytesStart<'_>) -> Option<ColumnWidth> {
    Some(ColumnWidth {
        min: attr_value(e, b"min")?.parse::<u32>().ok()?,
        max: attr_value(e, b"max")?.parse::<u32>().ok()?,
        width: attr_value(e, b"width")?.parse::<f64>().ok()?,
        custom_width: attr_truthy(attr_value(e, b"customWidth").as_deref()),
    })
}

fn parse_page_margins(e: &BytesStart<'_>) -> Option<PageMarginsInfo> {
    Some(PageMarginsInfo {
        left: attr_f64(e, b"left")?,
        right: attr_f64(e, b"right")?,
        top: attr_f64(e, b"top")?,
        bottom: attr_f64(e, b"bottom")?,
        header: attr_f64(e, b"header")?,
        footer: attr_f64(e, b"footer")?,
    })
}

fn parse_header_footer_item_text(text: &str) -> HeaderFooterItemInfo {
    let mut item = HeaderFooterItemInfo::default();
    let mut current: Option<u8> = None;
    let mut buf = String::new();
    let mut chars = text.chars().peekable();

    while let Some(ch) = chars.next() {
        if ch == '&' {
            match chars.peek().copied() {
                Some('L') | Some('C') | Some('R') => {
                    assign_header_footer_segment(&mut item, current, std::mem::take(&mut buf));
                    current = chars.next().map(|marker| marker as u8);
                    continue;
                }
                _ => {}
            }
        }
        buf.push(ch);
    }
    assign_header_footer_segment(&mut item, current, buf);
    item
}

fn assign_header_footer_segment(
    item: &mut HeaderFooterItemInfo,
    segment: Option<u8>,
    value: String,
) {
    if value.is_empty() {
        return;
    }
    match segment.unwrap_or(b'C') {
        b'L' => item.left = Some(value),
        b'R' => item.right = Some(value),
        _ => item.center = Some(value),
    }
}

fn append_break(
    row_breaks: &mut Option<PageBreakListInfo>,
    column_breaks: &mut Option<PageBreakListInfo>,
    active: Option<BreakListKind>,
    e: &BytesStart<'_>,
) {
    let Some(active) = active else {
        return;
    };
    let Some(id) = attr_u32(e, b"id") else {
        return;
    };
    let break_info = BreakInfo {
        id,
        min: attr_u32(e, b"min"),
        max: attr_u32(e, b"max"),
        man: attr_bool_default(e, b"man", true),
        pt: attr_bool_default(e, b"pt", false),
    };
    match active {
        BreakListKind::Row => {
            row_breaks
                .get_or_insert_with(PageBreakListInfo::default)
                .breaks
                .push(break_info);
        }
        BreakListKind::Column => {
            column_breaks
                .get_or_insert_with(PageBreakListInfo::default)
                .breaks
                .push(break_info);
        }
    }
}

fn update_row_visibility(
    e: &BytesStart<'_>,
    row: Option<u32>,
    hidden_rows: &mut HashMap<u32, bool>,
    row_outline_levels: &mut HashMap<u32, u8>,
) {
    let Some(row) = row else {
        return;
    };
    hidden_rows.insert(row, attr_truthy(attr_value(e, b"hidden").as_deref()));
    let outline_level = attr_value(e, b"outlineLevel")
        .and_then(|value| value.parse::<u8>().ok())
        .unwrap_or(0);
    if outline_level > 0 {
        row_outline_levels.insert(row, outline_level);
    } else {
        row_outline_levels.remove(&row);
    }
}

fn update_column_visibility(
    e: &BytesStart<'_>,
    hidden_columns: &mut HashMap<u32, bool>,
    column_outline_levels: &mut HashMap<u32, u8>,
) {
    let min = attr_value(e, b"min")
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(1);
    let max = attr_value(e, b"max")
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(min);
    let hidden = attr_truthy(attr_value(e, b"hidden").as_deref());
    let outline_level = attr_value(e, b"outlineLevel")
        .and_then(|value| value.parse::<u8>().ok())
        .unwrap_or(0);
    for col in min..=max {
        hidden_columns.insert(col, hidden);
        if outline_level > 0 {
            column_outline_levels.insert(col, outline_level);
        } else {
            column_outline_levels.remove(&col);
        }
    }
}

fn build_array_formula_map(cells: &[Cell]) -> HashMap<(u32, u32), ArrayFormulaInfo> {
    let mut out = HashMap::new();
    for cell in cells {
        let Some(info) = &cell.array_formula else {
            continue;
        };
        out.insert((cell.row, cell.col), info.clone());
        match info {
            ArrayFormulaInfo::Array { ref_range, .. }
            | ArrayFormulaInfo::DataTable { ref_range, .. } => {
                mark_array_spill_children(&mut out, ref_range, (cell.row, cell.col));
            }
            ArrayFormulaInfo::SpillChild => {}
        }
    }
    out
}

#[derive(Debug, Clone)]
struct SharedFormulaMaster {
    row: u32,
    col: u32,
    formula: String,
}

fn resolve_shared_formulas(cells: &mut [Cell]) {
    let mut masters: HashMap<String, SharedFormulaMaster> = HashMap::new();
    for cell in cells.iter() {
        if cell.formula_kind.as_deref() != Some("shared") {
            continue;
        }
        let (Some(index), Some(formula)) = (&cell.formula_shared_index, &cell.formula) else {
            continue;
        };
        masters
            .entry(index.clone())
            .or_insert_with(|| SharedFormulaMaster {
                row: cell.row,
                col: cell.col,
                formula: formula.clone(),
            });
    }

    for cell in cells.iter_mut() {
        if cell.formula_kind.as_deref() != Some("shared") || cell.formula.is_some() {
            continue;
        }
        let Some(index) = &cell.formula_shared_index else {
            continue;
        };
        let Some(master) = masters.get(index) else {
            continue;
        };
        cell.formula = Some(translate_shared_formula(
            &master.formula,
            cell.row as i32 - master.row as i32,
            cell.col as i32 - master.col as i32,
        ));
    }
}

fn translate_shared_formula(formula: &str, rows: i32, cols: i32) -> String {
    if rows == 0 && cols == 0 {
        return formula.to_string();
    }
    let had_equals = formula.starts_with('=');
    let wrapped;
    let input = if had_equals {
        formula
    } else {
        wrapped = format!("={formula}");
        &wrapped
    };
    let mut delta = RefDelta::empty();
    delta.rows = rows;
    delta.cols = cols;
    delta.anchor_row = 1;
    delta.anchor_col = 1;
    delta.respect_dollar = true;
    let translated = translate_formula(input, &delta).unwrap_or_else(|_| input.to_string());
    if had_equals {
        translated
    } else {
        translated
            .strip_prefix('=')
            .unwrap_or(&translated)
            .to_string()
    }
}

fn mark_array_spill_children(
    out: &mut HashMap<(u32, u32), ArrayFormulaInfo>,
    ref_range: &str,
    master: (u32, u32),
) {
    let Some((min_row, min_col, max_row, max_col)) = parse_range_bounds_1based(ref_range) else {
        return;
    };
    for row in min_row..=max_row {
        for col in min_col..=max_col {
            if (row, col) != master {
                out.entry((row, col))
                    .or_insert(ArrayFormulaInfo::SpillChild);
            }
        }
    }
}

fn parse_range_bounds_1based(range: &str) -> Option<(u32, u32, u32, u32)> {
    let clean = range.replace('$', "").to_ascii_uppercase();
    let mut parts = clean.split(':');
    let start = parts.next()?;
    let end = parts.next().unwrap_or(start);
    let (start_row, start_col) = a1_to_row_col(start)?;
    let (end_row, end_col) = a1_to_row_col(end)?;
    Some((
        start_row.min(end_row),
        start_col.min(end_col),
        start_row.max(end_row),
        start_col.max(end_col),
    ))
}

fn parse_auto_filter(xml: &str) -> Result<Option<AutoFilterInfo>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut auto_filter: Option<AutoFilterInfo> = None;
    let mut current_col: Option<FilterColumnInfo> = None;
    let mut filters: Option<FiltersBuilder> = None;
    let mut custom_filters: Option<CustomFiltersBuilder> = None;
    let mut sort_state: Option<SortStateInfo> = None;
    let mut in_auto_filter = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"autoFilter" => {
                    in_auto_filter = true;
                    auto_filter = attr_value(&e, b"ref")
                        .filter(|value| !value.trim().is_empty())
                        .map(|ref_range| AutoFilterInfo {
                            ref_range,
                            filter_columns: Vec::new(),
                            sort_state: None,
                        });
                }
                b"filterColumn" if in_auto_filter => {
                    current_col = Some(FilterColumnInfo::from_start(&e));
                }
                b"filters" if current_col.is_some() => {
                    filters = Some(FiltersBuilder::from_start(&e));
                }
                b"customFilters" if current_col.is_some() => {
                    custom_filters = Some(CustomFiltersBuilder::from_start(&e));
                }
                b"sortState" if in_auto_filter => {
                    sort_state = Some(SortStateInfo::from_start(&e));
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"autoFilter" => {
                    auto_filter = attr_value(&e, b"ref")
                        .filter(|value| !value.trim().is_empty())
                        .map(|ref_range| AutoFilterInfo {
                            ref_range,
                            filter_columns: Vec::new(),
                            sort_state: None,
                        });
                }
                b"filterColumn" if in_auto_filter => {
                    if let Some(auto_filter) = auto_filter.as_mut() {
                        auto_filter
                            .filter_columns
                            .push(FilterColumnInfo::from_start(&e));
                    }
                }
                b"filters" if current_col.is_some() => {
                    if let Some(col) = current_col.as_mut() {
                        let filter = FiltersBuilder::from_start(&e).finish();
                        if filter.is_some() {
                            col.filter = filter;
                        }
                    }
                }
                b"filter" if filters.is_some() => {
                    if let (Some(filters), Some(value)) = (filters.as_mut(), attr_value(&e, b"val"))
                    {
                        filters.values.push(value);
                    }
                }
                b"dateGroupItem" if current_col.is_some() => {
                    if let Some(item) = DateGroupItemInfo::from_start(&e) {
                        current_col.as_mut().unwrap().date_group_items.push(item);
                    }
                }
                b"customFilters" if current_col.is_some() => {
                    if let Some(col) = current_col.as_mut() {
                        col.filter = Some(CustomFiltersBuilder::from_start(&e).finish());
                    }
                }
                b"customFilter" if custom_filters.is_some() => {
                    if let Some(filter) = CustomFilterInfo::from_start(&e) {
                        custom_filters.as_mut().unwrap().filters.push(filter);
                    }
                }
                b"dynamicFilter" if current_col.is_some() => {
                    current_col.as_mut().unwrap().filter = Some(FilterInfo::Dynamic {
                        filter_type: attr_value(&e, b"type").unwrap_or_else(|| "null".to_string()),
                        val: attr_f64(&e, b"val"),
                        val_iso: attr_value(&e, b"valIso"),
                        max_val_iso: attr_value(&e, b"maxValIso"),
                    });
                }
                b"colorFilter" if current_col.is_some() => {
                    current_col.as_mut().unwrap().filter = Some(FilterInfo::Color {
                        dxf_id: attr_u32(&e, b"dxfId").unwrap_or(0),
                        cell_color: attr_bool_default(&e, b"cellColor", true),
                    });
                }
                b"iconFilter" if current_col.is_some() => {
                    current_col.as_mut().unwrap().filter = Some(FilterInfo::Icon {
                        icon_set: attr_value(&e, b"iconSet")
                            .unwrap_or_else(|| "3Arrows".to_string()),
                        icon_id: attr_u32(&e, b"iconId").unwrap_or(0),
                    });
                }
                b"top10" if current_col.is_some() => {
                    current_col.as_mut().unwrap().filter = Some(FilterInfo::Top10 {
                        top: attr_bool_default(&e, b"top", true),
                        percent: attr_bool_default(&e, b"percent", false),
                        val: attr_f64(&e, b"val").unwrap_or(10.0),
                        filter_val: attr_f64(&e, b"filterVal"),
                    });
                }
                b"sortCondition" if sort_state.is_some() => {
                    if let Some(condition) = SortConditionInfo::from_start(&e) {
                        sort_state.as_mut().unwrap().sort_conditions.push(condition);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"autoFilter" => {
                    if let Some(pending_sort_state) = sort_state.take() {
                        if let Some(auto_filter) = auto_filter.as_mut() {
                            auto_filter.sort_state = Some(pending_sort_state);
                        }
                    }
                    break;
                }
                b"filterColumn" if in_auto_filter => {
                    if let (Some(auto_filter), Some(col)) =
                        (auto_filter.as_mut(), current_col.take())
                    {
                        auto_filter.filter_columns.push(col);
                    }
                }
                b"filters" if current_col.is_some() => {
                    if let (Some(col), Some(filters)) = (current_col.as_mut(), filters.take()) {
                        if let Some(filter) = filters.finish() {
                            col.filter = Some(filter);
                        }
                    }
                }
                b"customFilters" if current_col.is_some() => {
                    if let (Some(col), Some(filters)) =
                        (current_col.as_mut(), custom_filters.take())
                    {
                        col.filter = Some(filters.finish());
                    }
                }
                b"sortState" if auto_filter.is_some() => {
                    if let Some(auto_filter) = auto_filter.as_mut() {
                        auto_filter.sort_state = sort_state.take();
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse autoFilter: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(auto_filter)
}

#[derive(Debug)]
struct FiltersBuilder {
    blank: bool,
    values: Vec<String>,
}

impl FiltersBuilder {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            blank: attr_bool_default(e, b"blank", false),
            values: Vec::new(),
        }
    }

    fn finish(self) -> Option<FilterInfo> {
        if !self.values.is_empty() {
            Some(FilterInfo::String {
                values: self.values,
            })
        } else if self.blank {
            Some(FilterInfo::Blank)
        } else {
            None
        }
    }
}

#[derive(Debug)]
struct CustomFiltersBuilder {
    and_: bool,
    filters: Vec<CustomFilterInfo>,
}

impl CustomFiltersBuilder {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            and_: attr_bool_default(e, b"and", false),
            filters: Vec::new(),
        }
    }

    fn finish(self) -> FilterInfo {
        FilterInfo::Custom {
            and_: self.and_,
            filters: self.filters,
        }
    }
}

impl FilterColumnInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            col_id: attr_u32(e, b"colId").unwrap_or(0),
            hidden_button: attr_bool_default(e, b"hiddenButton", false),
            show_button: attr_bool_default(e, b"showButton", true),
            filter: None,
            date_group_items: Vec::new(),
        }
    }
}

impl CustomFilterInfo {
    fn from_start(e: &BytesStart<'_>) -> Option<Self> {
        Some(Self {
            operator: attr_value(e, b"operator").unwrap_or_else(|| "equal".to_string()),
            val: attr_value(e, b"val")?,
        })
    }
}

impl DateGroupItemInfo {
    fn from_start(e: &BytesStart<'_>) -> Option<Self> {
        Some(Self {
            year: attr_u32(e, b"year")?,
            month: attr_u32(e, b"month"),
            day: attr_u32(e, b"day"),
            hour: attr_u32(e, b"hour"),
            minute: attr_u32(e, b"minute"),
            second: attr_u32(e, b"second"),
            date_time_grouping: attr_value(e, b"dateTimeGrouping")
                .unwrap_or_else(|| "year".to_string()),
        })
    }
}

impl SortStateInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            sort_conditions: Vec::new(),
            column_sort: attr_bool_default(e, b"columnSort", false),
            case_sensitive: attr_bool_default(e, b"caseSensitive", false),
            ref_range: attr_value(e, b"ref"),
        }
    }
}

impl SortConditionInfo {
    fn from_start(e: &BytesStart<'_>) -> Option<Self> {
        Some(Self {
            ref_range: attr_value(e, b"ref")?,
            descending: attr_bool_default(e, b"descending", false),
            sort_by: attr_value(e, b"sortBy").unwrap_or_else(|| "value".to_string()),
            custom_list: attr_value(e, b"customList"),
            dxf_id: attr_u32(e, b"dxfId"),
            icon_set: attr_value(e, b"iconSet"),
            icon_id: attr_u32(e, b"iconId"),
        })
    }
}

impl SheetProtection {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            sheet: attr_bool_default(e, b"sheet", false),
            objects: attr_bool_default(e, b"objects", false),
            scenarios: attr_bool_default(e, b"scenarios", false),
            format_cells: attr_bool_default(e, b"formatCells", true),
            format_columns: attr_bool_default(e, b"formatColumns", true),
            format_rows: attr_bool_default(e, b"formatRows", true),
            insert_columns: attr_bool_default(e, b"insertColumns", true),
            insert_rows: attr_bool_default(e, b"insertRows", true),
            insert_hyperlinks: attr_bool_default(e, b"insertHyperlinks", true),
            delete_columns: attr_bool_default(e, b"deleteColumns", true),
            delete_rows: attr_bool_default(e, b"deleteRows", true),
            select_locked_cells: attr_bool_default(e, b"selectLockedCells", false),
            sort: attr_bool_default(e, b"sort", true),
            auto_filter: attr_bool_default(e, b"autoFilter", true),
            pivot_tables: attr_bool_default(e, b"pivotTables", true),
            select_unlocked_cells: attr_bool_default(e, b"selectUnlockedCells", false),
            password_hash: attr_value(e, b"password"),
        }
    }
}

#[derive(Debug)]
struct DataValidationBuilder {
    range: String,
    validation_type: String,
    operator: Option<String>,
    formula1: Option<String>,
    formula2: Option<String>,
    formula1_text: String,
    formula2_text: String,
    allow_blank: bool,
    error_title: Option<String>,
    error: Option<String>,
}

impl DataValidationBuilder {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            range: attr_value(e, b"sqref").unwrap_or_default(),
            validation_type: attr_value(e, b"type").unwrap_or_else(|| "any".to_string()),
            operator: attr_value(e, b"operator"),
            formula1: None,
            formula2: None,
            formula1_text: String::new(),
            formula2_text: String::new(),
            allow_blank: attr_value(e, b"allowBlank")
                .map(|value| attr_truthy(Some(&value)))
                .unwrap_or(true),
            error_title: attr_value(e, b"errorTitle").filter(|value| !value.is_empty()),
            error: attr_value(e, b"error").filter(|value| !value.is_empty()),
        }
    }

    fn push_text(&mut self, target: DataValidationFormula, text: &str) {
        match target {
            DataValidationFormula::Formula1 => self.formula1_text.push_str(text),
            DataValidationFormula::Formula2 => self.formula2_text.push_str(text),
        }
    }

    fn finish_formula1(&mut self) {
        let formula = self.formula1_text.trim();
        if !formula.is_empty() {
            self.formula1 = Some(ensure_formula_prefix(formula));
        }
    }

    fn finish_formula2(&mut self) {
        let formula = self.formula2_text.trim();
        if !formula.is_empty() {
            self.formula2 = Some(ensure_formula_prefix(formula));
        }
    }

    fn finish(self) -> DataValidation {
        DataValidation {
            range: self.range,
            validation_type: self.validation_type,
            operator: self.operator,
            formula1: self.formula1,
            formula2: self.formula2,
            allow_blank: self.allow_blank,
            error_title: self.error_title,
            error: self.error,
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum DataValidationFormula {
    Formula1,
    Formula2,
}

#[derive(Debug)]
struct ConditionalFormatBuilder {
    range: String,
    rule_type: String,
    operator: Option<String>,
    formula: Option<String>,
    formula_text: String,
    priority: Option<i64>,
    stop_if_true: Option<bool>,
}

impl ConditionalFormatBuilder {
    fn from_start(e: &BytesStart<'_>, range: String) -> Self {
        Self {
            range,
            rule_type: attr_value(e, b"type").unwrap_or_default(),
            operator: attr_value(e, b"operator"),
            formula: None,
            formula_text: String::new(),
            priority: attr_value(e, b"priority").and_then(|value| value.parse::<i64>().ok()),
            stop_if_true: attr_value(e, b"stopIfTrue").map(|value| attr_truthy(Some(&value))),
        }
    }

    fn push_formula_text(&mut self, text: &str) {
        self.formula_text.push_str(text);
    }

    fn finish_formula(&mut self) {
        let formula = self.formula_text.trim();
        if !formula.is_empty() && self.formula.is_none() {
            self.formula = Some(ensure_formula_prefix(formula));
        }
    }

    fn finish(self) -> ConditionalFormatRule {
        ConditionalFormatRule {
            range: self.range,
            rule_type: self.rule_type,
            operator: self.operator,
            formula: self.formula,
            priority: self.priority,
            stop_if_true: self.stop_if_true,
        }
    }
}

fn parse_pane(e: &BytesStart<'_>) -> Option<FreezePane> {
    let mode = match attr_value(e, b"state")?.to_ascii_lowercase().as_str() {
        "split" => PaneMode::Split,
        state if state.starts_with("frozen") => PaneMode::Freeze,
        _ => return None,
    };
    Some(FreezePane {
        mode,
        top_left_cell: attr_value(e, b"topLeftCell").filter(|value| !value.is_empty()),
        x_split: attr_value(e, b"xSplit")
            .and_then(|value| value.parse::<f64>().ok())
            .map(|value| value as i64),
        y_split: attr_value(e, b"ySplit")
            .and_then(|value| value.parse::<f64>().ok())
            .map(|value| value as i64),
        active_pane: attr_value(e, b"activePane").filter(|value| !value.is_empty()),
    })
}

#[derive(Debug)]
struct HyperlinkNode {
    cell: String,
    rid: Option<String>,
    location: Option<String>,
    display: Option<String>,
    tooltip: Option<String>,
}

impl HyperlinkNode {
    fn from_start(e: &BytesStart<'_>) -> Option<Self> {
        let cell = attr_value(e, b"ref").filter(|value| !value.is_empty())?;
        Some(Self {
            cell,
            rid: attr_value(e, b"r:id"),
            location: attr_value(e, b"location"),
            display: attr_value(e, b"display"),
            tooltip: attr_value(e, b"tooltip"),
        })
    }
}

fn resolve_hyperlinks(
    nodes: Vec<HyperlinkNode>,
    rels: Option<&RelsGraph>,
    cells: &[Cell],
) -> Vec<Hyperlink> {
    let mut out = Vec::new();
    for node in nodes {
        let internal = node.location.is_some() && node.rid.is_none();
        let target = if let Some(rid) = &node.rid {
            rels.and_then(|rels| rels.get(&RelId(rid.clone())))
                .map(|rel| rel.target.clone())
                .unwrap_or_default()
        } else {
            node.location.clone().unwrap_or_default()
        };
        if target.is_empty() {
            continue;
        }
        let display = match node.display {
            Some(display) if !display.is_empty() => display,
            _ => cell_display_text(cells, &node.cell),
        };
        let tooltip = node
            .tooltip
            .and_then(|t| if t.is_empty() { None } else { Some(t) });
        out.push(Hyperlink {
            cell: node.cell,
            target,
            display,
            tooltip,
            internal,
        });
    }
    out
}

fn cell_display_text(cells: &[Cell], coordinate: &str) -> String {
    let Some(cell) = cells.iter().find(|cell| cell.coordinate == coordinate) else {
        return String::new();
    };
    match &cell.value {
        CellValue::Empty | CellValue::Error(_) => String::new(),
        CellValue::String(value) => value.clone(),
        CellValue::Number(value) => value.to_string(),
        CellValue::Bool(value) => value.to_string(),
    }
}

fn comments_target(rels: &RelsGraph) -> Option<String> {
    rels.iter()
        .find(|rel| rel.rel_type.ends_with("/comments") || rel.rel_type == "comments")
        .map(|rel| rel.target.clone())
}

fn parse_comments(xml: &str) -> Result<Vec<Comment>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut authors = Vec::new();
    let mut comments = Vec::new();
    let mut in_author = false;
    let mut in_comment = false;
    let mut in_t = false;
    let mut current_cell = String::new();
    let mut current_author_id = 0usize;
    let mut current_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"author" => in_author = true,
                b"comment" => {
                    in_comment = true;
                    current_text.clear();
                    current_cell = attr_value(&e, b"ref").unwrap_or_default();
                    current_author_id = attr_value(&e, b"authorId")
                        .and_then(|value| value.parse::<usize>().ok())
                        .unwrap_or(0);
                }
                b"t" => in_t = true,
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"author" => in_author = false,
                b"comment" => {
                    in_comment = false;
                    comments.push(Comment {
                        cell: current_cell.clone(),
                        text: current_text.clone(),
                        author: authors.get(current_author_id).cloned().unwrap_or_default(),
                        threaded: false,
                    });
                }
                b"t" => in_t = false,
                _ => {}
            },
            Ok(Event::Text(e)) => {
                let text = e
                    .unescape()
                    .map_err(|err| ReaderError::Xml(format!("comments text: {err}")))?
                    .to_string();
                if in_author {
                    authors.push(text);
                } else if in_comment && in_t {
                    current_text.push_str(&text);
                }
            }
            Ok(Event::CData(e)) => {
                if in_comment && in_t {
                    current_text.push_str(&String::from_utf8_lossy(e.as_ref()));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ReaderError::Xml(format!(
                    "failed to parse comments XML: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(comments)
}

fn parse_doc_properties_into(
    xml: &str,
    out: &mut HashMap<String, String>,
    key_for_tag: fn(&[u8]) -> Option<&'static str>,
) {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut current_tag: Option<Vec<u8>> = None;
    let mut current_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                current_tag = Some(e.local_name().as_ref().to_vec());
                current_text.clear();
            }
            Ok(Event::Text(e)) => {
                if current_tag.is_some() {
                    current_text.push_str(&e.unescape().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => {
                let name = e.local_name();
                let name = name.as_ref();
                if current_tag.as_deref() == Some(name) {
                    if let Some(key) = key_for_tag(name) {
                        let value = current_text.trim();
                        if !value.is_empty() {
                            out.entry(key.to_string())
                                .or_insert_with(|| value.to_string());
                        }
                    }
                }
                current_tag = None;
                current_text.clear();
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
}

fn doc_property_core_key(name: &[u8]) -> Option<&'static str> {
    match name {
        b"title" => Some("title"),
        b"subject" => Some("subject"),
        b"creator" => Some("creator"),
        b"keywords" => Some("keywords"),
        b"description" => Some("description"),
        b"lastModifiedBy" => Some("lastModifiedBy"),
        b"category" => Some("category"),
        b"contentStatus" => Some("contentStatus"),
        b"identifier" => Some("identifier"),
        b"language" => Some("language"),
        b"revision" => Some("revision"),
        b"version" => Some("version"),
        b"created" => Some("created"),
        b"modified" => Some("modified"),
        _ => None,
    }
}

fn doc_property_app_key(name: &[u8]) -> Option<&'static str> {
    match name {
        b"Company" => Some("company"),
        b"Manager" => Some("manager"),
        b"Application" => Some("application"),
        _ => None,
    }
}

fn read_tables<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    sheet_path: &str,
    sheet_xml: &str,
    rels: Option<&RelsGraph>,
) -> Result<Vec<Table>> {
    let Some(rels) = rels else {
        return Ok(Vec::new());
    };
    let table_rids = parse_table_rids(sheet_xml)?;
    if table_rids.is_empty() {
        return Ok(Vec::new());
    }

    let sheet_dir = part_dir(sheet_path);
    let mut out = Vec::new();
    for rid in table_rids {
        let Some(rel) = rels.get(&RelId(rid)) else {
            continue;
        };
        let table_path = join_and_normalize(&sheet_dir, &rel.target);
        let Some(table_xml) = read_part_optional(zip, &table_path)? else {
            continue;
        };
        if let Ok(table) = parse_table_xml(&table_xml) {
            out.push(table);
        }
    }
    Ok(out)
}

fn read_images<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    sheet_path: &str,
    rels: Option<&RelsGraph>,
) -> Result<Vec<ImageInfo>> {
    let Some(rels) = rels else {
        return Ok(Vec::new());
    };
    let sheet_dir = part_dir(sheet_path);
    let mut out = Vec::new();

    for drawing_rel in rels.find_by_type(wolfxl_rels::rt::DRAWING) {
        let drawing_path = join_and_normalize(&sheet_dir, &drawing_rel.target);
        let Some(drawing_xml) = read_part_optional(zip, &drawing_path)? else {
            continue;
        };
        let drawing_rels_path = sheet_rels_path(&drawing_path);
        let image_rels = read_part_optional(zip, &drawing_rels_path)?
            .map(|xml| {
                RelsGraph::parse(xml.as_bytes())
                    .map_err(|e| ReaderError::Xml(format!("failed to parse drawing rels: {e}")))
            })
            .transpose()?;
        let Some(image_rels) = image_rels else {
            continue;
        };
        let image_targets: HashMap<String, String> = image_rels
            .iter()
            .filter(|rel| rel.rel_type == wolfxl_rels::rt::IMAGE)
            .map(|rel| (rel.id.0.clone(), rel.target.clone()))
            .collect();
        let drawing_dir = part_dir(&drawing_path);

        for image_ref in parse_drawing_images(&drawing_xml)? {
            let Some(target) = image_targets.get(&image_ref.rid) else {
                continue;
            };
            let image_path = join_and_normalize(&drawing_dir, target);
            let Some(data) = read_part_optional_bytes(zip, &image_path)? else {
                continue;
            };
            out.push(ImageInfo {
                data,
                ext: image_ext_from_path(&image_path),
                anchor: image_ref.anchor,
            });
        }
    }

    Ok(out)
}

fn read_charts<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    sheet_path: &str,
    rels: Option<&RelsGraph>,
) -> Result<Vec<ChartInfo>> {
    let Some(rels) = rels else {
        return Ok(Vec::new());
    };
    let sheet_dir = part_dir(sheet_path);
    let mut out = Vec::new();

    for drawing_rel in rels.find_by_type(wolfxl_rels::rt::DRAWING) {
        let drawing_path = join_and_normalize(&sheet_dir, &drawing_rel.target);
        let Some(drawing_xml) = read_part_optional(zip, &drawing_path)? else {
            continue;
        };
        let drawing_rels_path = sheet_rels_path(&drawing_path);
        let chart_rels = read_part_optional(zip, &drawing_rels_path)?
            .map(|xml| {
                RelsGraph::parse(xml.as_bytes())
                    .map_err(|e| ReaderError::Xml(format!("failed to parse drawing rels: {e}")))
            })
            .transpose()?;
        let Some(chart_rels) = chart_rels else {
            continue;
        };
        let chart_targets: HashMap<String, String> = chart_rels
            .iter()
            .filter(|rel| rel.rel_type == wolfxl_rels::rt::CHART)
            .map(|rel| (rel.id.0.clone(), rel.target.clone()))
            .collect();
        let drawing_dir = part_dir(&drawing_path);

        for chart_ref in parse_drawing_charts(&drawing_xml)? {
            let Some(target) = chart_targets.get(&chart_ref.rid) else {
                continue;
            };
            let chart_path = join_and_normalize(&drawing_dir, target);
            let Some(chart_xml) = read_part_optional(zip, &chart_path)? else {
                continue;
            };
            if let Some(chart) = parse_chart_xml(&chart_xml, chart_ref.anchor)? {
                out.push(chart);
            }
        }
    }

    Ok(out)
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DrawingImageRef {
    rid: String,
    anchor: ImageAnchorInfo,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DrawingChartRef {
    rid: String,
    anchor: ImageAnchorInfo,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DrawingAnchorKind {
    OneCell,
    TwoCell,
    Absolute,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DrawingImageBuilder {
    kind: DrawingAnchorKind,
    from: AnchorMarkerInfo,
    to: AnchorMarkerInfo,
    pos: AnchorPositionInfo,
    ext: Option<AnchorExtentInfo>,
    edit_as: Option<String>,
    rid: Option<String>,
}

impl DrawingImageBuilder {
    fn new(kind: DrawingAnchorKind, e: &BytesStart<'_>) -> Self {
        Self {
            kind,
            from: AnchorMarkerInfo::default(),
            to: AnchorMarkerInfo::default(),
            pos: AnchorPositionInfo::default(),
            ext: None,
            edit_as: attr_value(e, b"editAs"),
            rid: None,
        }
    }

    fn finish(self) -> Option<DrawingImageRef> {
        let rid = self.rid?;
        let anchor = match self.kind {
            DrawingAnchorKind::OneCell => ImageAnchorInfo::OneCell {
                from: self.from,
                ext: self.ext,
            },
            DrawingAnchorKind::TwoCell => ImageAnchorInfo::TwoCell {
                from: self.from,
                to: self.to,
                edit_as: self.edit_as.unwrap_or_else(|| "oneCell".to_string()),
            },
            DrawingAnchorKind::Absolute => ImageAnchorInfo::Absolute {
                pos: self.pos,
                ext: self.ext.unwrap_or_default(),
            },
        };
        Some(DrawingImageRef { rid, anchor })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DrawingChartBuilder {
    kind: DrawingAnchorKind,
    from: AnchorMarkerInfo,
    to: AnchorMarkerInfo,
    pos: AnchorPositionInfo,
    ext: Option<AnchorExtentInfo>,
    edit_as: Option<String>,
    rid: Option<String>,
}

impl DrawingChartBuilder {
    fn new(kind: DrawingAnchorKind, e: &BytesStart<'_>) -> Self {
        Self {
            kind,
            from: AnchorMarkerInfo::default(),
            to: AnchorMarkerInfo::default(),
            pos: AnchorPositionInfo::default(),
            ext: None,
            edit_as: attr_value(e, b"editAs"),
            rid: None,
        }
    }

    fn finish(self) -> Option<DrawingChartRef> {
        let rid = self.rid?;
        let anchor = match self.kind {
            DrawingAnchorKind::OneCell => ImageAnchorInfo::OneCell {
                from: self.from,
                ext: self.ext,
            },
            DrawingAnchorKind::TwoCell => ImageAnchorInfo::TwoCell {
                from: self.from,
                to: self.to,
                edit_as: self.edit_as.unwrap_or_else(|| "oneCell".to_string()),
            },
            DrawingAnchorKind::Absolute => ImageAnchorInfo::Absolute {
                pos: self.pos,
                ext: self.ext.unwrap_or_default(),
            },
        };
        Some(DrawingChartRef { rid, anchor })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MarkerSlot {
    From,
    To,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MarkerTextTarget {
    Col,
    Row,
    ColOff,
    RowOff,
}

fn parse_drawing_images(xml: &str) -> Result<Vec<DrawingImageRef>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    let mut current: Option<DrawingImageBuilder> = None;
    let mut marker_slot: Option<MarkerSlot> = None;
    let mut marker_text: Option<MarkerTextTarget> = None;
    let mut in_pic = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"oneCellAnchor" => {
                    current = Some(DrawingImageBuilder::new(DrawingAnchorKind::OneCell, &e));
                }
                b"twoCellAnchor" => {
                    current = Some(DrawingImageBuilder::new(DrawingAnchorKind::TwoCell, &e));
                }
                b"absoluteAnchor" => {
                    current = Some(DrawingImageBuilder::new(DrawingAnchorKind::Absolute, &e));
                }
                b"from" if current.is_some() => marker_slot = Some(MarkerSlot::From),
                b"to" if current.is_some() => marker_slot = Some(MarkerSlot::To),
                b"col" if marker_slot.is_some() => marker_text = Some(MarkerTextTarget::Col),
                b"row" if marker_slot.is_some() => marker_text = Some(MarkerTextTarget::Row),
                b"colOff" if marker_slot.is_some() => {
                    marker_text = Some(MarkerTextTarget::ColOff);
                }
                b"rowOff" if marker_slot.is_some() => {
                    marker_text = Some(MarkerTextTarget::RowOff);
                }
                b"pic" if current.is_some() => in_pic = true,
                b"pos" => apply_anchor_pos(&mut current, &e),
                b"ext" if !in_pic => apply_anchor_ext(&mut current, &e),
                b"blip" => apply_blip_rid(&mut current, &e),
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"pos" => apply_anchor_pos(&mut current, &e),
                b"ext" if !in_pic => apply_anchor_ext(&mut current, &e),
                b"blip" => apply_blip_rid(&mut current, &e),
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if let (Some(slot), Some(target), Some(builder)) =
                    (marker_slot, marker_text, current.as_mut())
                {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("drawing text: {err}")))?;
                    if let Ok(value) = text.parse::<i64>() {
                        apply_marker_value(builder, slot, target, value);
                    }
                }
            }
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"col" | b"row" | b"colOff" | b"rowOff" => marker_text = None,
                b"from" | b"to" => {
                    marker_slot = None;
                    marker_text = None;
                }
                b"pic" => in_pic = false,
                b"oneCellAnchor" | b"twoCellAnchor" | b"absoluteAnchor" => {
                    marker_slot = None;
                    marker_text = None;
                    in_pic = false;
                    if let Some(builder) = current.take().and_then(DrawingImageBuilder::finish) {
                        out.push(builder);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse drawing: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_drawing_charts(xml: &str) -> Result<Vec<DrawingChartRef>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    let mut current: Option<DrawingChartBuilder> = None;
    let mut marker_slot: Option<MarkerSlot> = None;
    let mut marker_text: Option<MarkerTextTarget> = None;
    let mut in_graphic_frame = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"oneCellAnchor" => {
                    current = Some(DrawingChartBuilder::new(DrawingAnchorKind::OneCell, &e));
                }
                b"twoCellAnchor" => {
                    current = Some(DrawingChartBuilder::new(DrawingAnchorKind::TwoCell, &e));
                }
                b"absoluteAnchor" => {
                    current = Some(DrawingChartBuilder::new(DrawingAnchorKind::Absolute, &e));
                }
                b"from" if current.is_some() => marker_slot = Some(MarkerSlot::From),
                b"to" if current.is_some() => marker_slot = Some(MarkerSlot::To),
                b"col" if marker_slot.is_some() => marker_text = Some(MarkerTextTarget::Col),
                b"row" if marker_slot.is_some() => marker_text = Some(MarkerTextTarget::Row),
                b"colOff" if marker_slot.is_some() => {
                    marker_text = Some(MarkerTextTarget::ColOff);
                }
                b"rowOff" if marker_slot.is_some() => {
                    marker_text = Some(MarkerTextTarget::RowOff);
                }
                b"graphicFrame" if current.is_some() => in_graphic_frame = true,
                b"pos" => apply_chart_anchor_pos(&mut current, &e),
                b"ext" if !in_graphic_frame => apply_chart_anchor_ext(&mut current, &e),
                b"chart" => apply_chart_rid(&mut current, &e),
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"pos" => apply_chart_anchor_pos(&mut current, &e),
                b"ext" if !in_graphic_frame => apply_chart_anchor_ext(&mut current, &e),
                b"chart" => apply_chart_rid(&mut current, &e),
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if let (Some(slot), Some(target), Some(builder)) =
                    (marker_slot, marker_text, current.as_mut())
                {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("drawing text: {err}")))?;
                    if let Ok(value) = text.parse::<i64>() {
                        apply_chart_marker_value(builder, slot, target, value);
                    }
                }
            }
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"col" | b"row" | b"colOff" | b"rowOff" => marker_text = None,
                b"from" | b"to" => {
                    marker_slot = None;
                    marker_text = None;
                }
                b"graphicFrame" => in_graphic_frame = false,
                b"oneCellAnchor" | b"twoCellAnchor" | b"absoluteAnchor" => {
                    marker_slot = None;
                    marker_text = None;
                    in_graphic_frame = false;
                    if let Some(builder) = current.take().and_then(DrawingChartBuilder::finish) {
                        out.push(builder);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse drawing: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn apply_anchor_pos(builder: &mut Option<DrawingImageBuilder>, e: &BytesStart<'_>) {
    let Some(builder) = builder.as_mut() else {
        return;
    };
    builder.pos = AnchorPositionInfo {
        x: attr_i64(e, b"x").unwrap_or_default(),
        y: attr_i64(e, b"y").unwrap_or_default(),
    };
}

fn apply_chart_anchor_pos(builder: &mut Option<DrawingChartBuilder>, e: &BytesStart<'_>) {
    let Some(builder) = builder.as_mut() else {
        return;
    };
    builder.pos = AnchorPositionInfo {
        x: attr_i64(e, b"x").unwrap_or_default(),
        y: attr_i64(e, b"y").unwrap_or_default(),
    };
}

fn apply_chart_anchor_ext(builder: &mut Option<DrawingChartBuilder>, e: &BytesStart<'_>) {
    let Some(builder) = builder.as_mut() else {
        return;
    };
    let Some(cx) = attr_i64(e, b"cx") else {
        return;
    };
    let Some(cy) = attr_i64(e, b"cy") else {
        return;
    };
    builder.ext = Some(AnchorExtentInfo { cx, cy });
}

fn apply_chart_rid(builder: &mut Option<DrawingChartBuilder>, e: &BytesStart<'_>) {
    let Some(builder) = builder.as_mut() else {
        return;
    };
    builder.rid = attr_value(e, b"r:id").or_else(|| attr_value(e, b"id"));
}

fn apply_chart_marker_value(
    builder: &mut DrawingChartBuilder,
    slot: MarkerSlot,
    target: MarkerTextTarget,
    value: i64,
) {
    let marker = match slot {
        MarkerSlot::From => &mut builder.from,
        MarkerSlot::To => &mut builder.to,
    };
    match target {
        MarkerTextTarget::Col => marker.col = value,
        MarkerTextTarget::Row => marker.row = value,
        MarkerTextTarget::ColOff => marker.col_off = value,
        MarkerTextTarget::RowOff => marker.row_off = value,
    }
}

fn parse_chart_xml(xml: &str, anchor: ImageAnchorInfo) -> Result<Option<ChartInfo>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut kind: Option<String> = None;
    let mut title_parts: Vec<String> = Vec::new();
    let mut style: Option<u32> = None;
    let mut current_series: Option<ChartSeriesInfo> = None;
    let mut series = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let local = e.local_name().as_ref().to_vec();
                apply_chart_start(&local, &e, &mut kind, &mut style, &mut current_series);
                stack.push(local);
            }
            Ok(Event::Empty(e)) => {
                let local = e.local_name().as_ref().to_vec();
                apply_chart_start(&local, &e, &mut kind, &mut style, &mut current_series);
            }
            Ok(Event::Text(e)) => {
                let text = e
                    .unescape()
                    .map_err(|err| ReaderError::Xml(format!("chart text: {err}")))?
                    .to_string();
                apply_chart_text(&stack, text.trim(), &mut title_parts, &mut current_series);
            }
            Ok(Event::End(e)) => {
                if e.local_name().as_ref() == b"ser" {
                    if let Some(ser) = current_series.take() {
                        series.push(ser);
                    }
                }
                stack.pop();
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse chart: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    let Some(kind) = kind else {
        return Ok(None);
    };
    let title = if title_parts.is_empty() {
        None
    } else {
        Some(title_parts.join(""))
    };
    Ok(Some(ChartInfo {
        kind,
        title,
        style,
        anchor,
        series,
    }))
}

fn apply_chart_start(
    local: &[u8],
    e: &BytesStart<'_>,
    kind: &mut Option<String>,
    style: &mut Option<u32>,
    current_series: &mut Option<ChartSeriesInfo>,
) {
    if kind.is_none() {
        *kind = chart_kind(local).map(str::to_string);
    }
    match local {
        b"style" => {
            *style = attr_u32(e, b"val");
        }
        b"ser" => {
            *current_series = Some(ChartSeriesInfo::default());
        }
        b"idx" => {
            if let Some(series) = current_series.as_mut() {
                series.idx = attr_u32(e, b"val");
            }
        }
        b"order" => {
            if let Some(series) = current_series.as_mut() {
                series.order = attr_u32(e, b"val");
            }
        }
        _ => {}
    }
}

fn apply_chart_text(
    stack: &[Vec<u8>],
    text: &str,
    title_parts: &mut Vec<String>,
    current_series: &mut Option<ChartSeriesInfo>,
) {
    if text.is_empty() {
        return;
    }
    let Some(last) = stack.last().map(Vec::as_slice) else {
        return;
    };
    if let Some(series) = current_series.as_mut() {
        if last == b"f" {
            if chart_path_contains(stack, b"tx") && chart_path_contains(stack, b"strRef") {
                series.title_ref = Some(text.to_string());
            } else if chart_path_contains(stack, b"cat") {
                series.cat_ref = Some(text.to_string());
            } else if chart_path_contains(stack, b"val") {
                series.val_ref = Some(text.to_string());
            } else if chart_path_contains(stack, b"xVal") {
                series.x_ref = Some(text.to_string());
            } else if chart_path_contains(stack, b"yVal") {
                series.y_ref = Some(text.to_string());
            } else if chart_path_contains(stack, b"bubbleSize") {
                series.bubble_size_ref = Some(text.to_string());
            }
        } else if last == b"v" && chart_path_contains(stack, b"tx") {
            series.title_value = Some(text.to_string());
        }
        return;
    }

    if last == b"t" && chart_path_contains(stack, b"title") {
        title_parts.push(text.to_string());
    }
}

fn chart_path_contains(stack: &[Vec<u8>], name: &[u8]) -> bool {
    stack.iter().any(|part| part.as_slice() == name)
}

fn chart_kind(local: &[u8]) -> Option<&'static str> {
    match local {
        b"barChart" => Some("bar"),
        b"bar3DChart" => Some("bar3d"),
        b"lineChart" => Some("line"),
        b"line3DChart" => Some("line3d"),
        b"pieChart" => Some("pie"),
        b"pie3DChart" => Some("pie3d"),
        b"ofPieChart" => Some("of_pie"),
        b"doughnutChart" => Some("doughnut"),
        b"areaChart" => Some("area"),
        b"area3DChart" => Some("area3d"),
        b"scatterChart" => Some("scatter"),
        b"bubbleChart" => Some("bubble"),
        b"radarChart" => Some("radar"),
        b"surfaceChart" => Some("surface"),
        b"surface3DChart" => Some("surface3d"),
        b"stockChart" => Some("stock"),
        _ => None,
    }
}

fn apply_anchor_ext(builder: &mut Option<DrawingImageBuilder>, e: &BytesStart<'_>) {
    let Some(builder) = builder.as_mut() else {
        return;
    };
    let Some(cx) = attr_i64(e, b"cx") else {
        return;
    };
    let Some(cy) = attr_i64(e, b"cy") else {
        return;
    };
    builder.ext = Some(AnchorExtentInfo { cx, cy });
}

fn apply_blip_rid(builder: &mut Option<DrawingImageBuilder>, e: &BytesStart<'_>) {
    let Some(builder) = builder.as_mut() else {
        return;
    };
    builder.rid = attr_value(e, b"r:embed").or_else(|| attr_value(e, b"embed"));
}

fn apply_marker_value(
    builder: &mut DrawingImageBuilder,
    slot: MarkerSlot,
    target: MarkerTextTarget,
    value: i64,
) {
    let marker = match slot {
        MarkerSlot::From => &mut builder.from,
        MarkerSlot::To => &mut builder.to,
    };
    match target {
        MarkerTextTarget::Col => marker.col = value,
        MarkerTextTarget::Row => marker.row = value,
        MarkerTextTarget::ColOff => marker.col_off = value,
        MarkerTextTarget::RowOff => marker.row_off = value,
    }
}

fn image_ext_from_path(path: &str) -> String {
    PathBuf::from(path)
        .extension()
        .and_then(|ext| ext.to_str())
        .map(str::to_ascii_lowercase)
        .unwrap_or_default()
}

fn parse_table_rids(xml: &str) -> Result<Vec<String>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"tablePart" {
                    if let Some(rid) = attr_value(&e, b"r:id").filter(|value| !value.is_empty()) {
                        out.push(rid);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ReaderError::Xml(format!(
                    "failed to parse worksheet table parts: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_table_xml(xml: &str) -> Result<Table> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut name = String::new();
    let mut ref_range = String::new();
    let mut header_row = true;
    let mut totals_row = false;
    let mut style = None;
    let mut columns = Vec::new();
    let mut autofilter = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"table" => {
                    name = attr_value(&e, b"name")
                        .or_else(|| attr_value(&e, b"displayName"))
                        .unwrap_or_default();
                    ref_range = attr_value(&e, b"ref").unwrap_or_default();
                    header_row = attr_value(&e, b"headerRowCount")
                        .map(|value| value != "0")
                        .unwrap_or(true);
                    totals_row = attr_value(&e, b"totalsRowCount")
                        .map(|value| value != "0")
                        .unwrap_or(false);
                }
                b"tableStyleInfo" => {
                    style = attr_value(&e, b"name").filter(|value| !value.is_empty());
                }
                b"tableColumn" => {
                    if let Some(column) = attr_value(&e, b"name") {
                        columns.push(column);
                    }
                }
                b"autoFilter" => {
                    autofilter = true;
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse table XML: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(Table {
        name,
        ref_range,
        header_row,
        totals_row,
        style,
        columns,
        autofilter,
    })
}

#[derive(Debug)]
struct CellBuilder {
    row: u32,
    col: u32,
    coordinate: String,
    style_id: Option<u32>,
    data_type: CellDataType,
    value_text: String,
    inline_text: String,
    formula_text: String,
    formula_kind: Option<String>,
    formula_shared_index: Option<String>,
    formula_ref: Option<String>,
    formula_ca: bool,
    formula_dt2_d: bool,
    formula_dtr: bool,
    formula_r1: Option<String>,
    formula_r2: Option<String>,
    inline_runs: Vec<RichTextRun>,
    inline_current_run: Option<RichTextRun>,
    inline_current_props: Option<InlineFontProps>,
    inline_saw_r: bool,
}

impl CellBuilder {
    fn from_start(e: &BytesStart<'_>, current_row: Option<u32>) -> Self {
        let coordinate = attr_value(e, b"r");
        let (row, col) = coordinate
            .as_deref()
            .and_then(a1_to_row_col)
            .unwrap_or((current_row.unwrap_or(1), 1));
        let coordinate = coordinate.unwrap_or_else(|| row_col_to_a1(row, col));
        let data_type = match attr_value(e, b"t").as_deref() {
            Some("s") => CellDataType::SharedString,
            Some("inlineStr") => CellDataType::InlineString,
            Some("str") => CellDataType::FormulaString,
            Some("b") => CellDataType::Bool,
            Some("e") => CellDataType::Error,
            _ => CellDataType::Number,
        };
        let style_id = attr_value(e, b"s").and_then(|v| v.parse::<u32>().ok());
        Self {
            row,
            col,
            coordinate,
            style_id,
            data_type,
            value_text: String::new(),
            inline_text: String::new(),
            formula_text: String::new(),
            formula_kind: None,
            formula_shared_index: None,
            formula_ref: None,
            formula_ca: false,
            formula_dt2_d: false,
            formula_dtr: false,
            formula_r1: None,
            formula_r2: None,
            inline_runs: Vec::new(),
            inline_current_run: None,
            inline_current_props: None,
            inline_saw_r: false,
        }
    }

    fn start_formula(&mut self, e: &BytesStart<'_>) {
        self.formula_kind = attr_value(e, b"t");
        self.formula_shared_index = attr_value(e, b"si");
        self.formula_ref = attr_value(e, b"ref");
        self.formula_ca = attr_truthy(attr_value(e, b"ca").as_deref());
        self.formula_dt2_d = attr_truthy(attr_value(e, b"dt2D").as_deref());
        self.formula_dtr = attr_truthy(attr_value(e, b"dtr").as_deref());
        self.formula_r1 = attr_value(e, b"r1");
        self.formula_r2 = attr_value(e, b"r2");
    }

    fn push_text(&mut self, target: TextTarget, text: &str) {
        match target {
            TextTarget::Value => self.value_text.push_str(text),
            TextTarget::Formula => self.formula_text.push_str(text),
            TextTarget::InlineString => {
                let text = normalize_ooxml_text(text);
                self.inline_text.push_str(&text);
                if let Some(run) = self.inline_current_run.as_mut() {
                    run.text.push_str(&text);
                }
            }
        }
    }

    fn start_inline_run(&mut self) {
        self.inline_saw_r = true;
        self.inline_current_run = Some(RichTextRun::default());
        self.inline_current_props = None;
    }

    fn start_inline_props(&mut self) {
        if self.inline_current_run.is_some() {
            self.inline_current_props = Some(InlineFontProps::default());
        }
    }

    fn apply_inline_font_tag(
        &mut self,
        tag: &[u8],
        attrs: quick_xml::events::attributes::Attributes<'_>,
    ) {
        if let Some(props) = self.inline_current_props.as_mut() {
            apply_rich_font_attr(props, tag, attrs);
        }
    }

    fn end_inline_props(&mut self) {
        if let (Some(run), Some(props)) = (
            self.inline_current_run.as_mut(),
            self.inline_current_props.take(),
        ) {
            run.font = Some(props);
        }
    }

    fn end_inline_run(&mut self) {
        if let Some(run) = self.inline_current_run.take() {
            self.inline_runs.push(run);
        }
        self.inline_current_props = None;
    }

    fn finish(mut self, shared_strings: &SharedStrings) -> Result<Cell> {
        if self.inline_current_props.is_some() {
            self.end_inline_props();
        }
        if self.inline_current_run.is_some() {
            self.end_inline_run();
        }
        let shared_string_idx = if self.data_type == CellDataType::SharedString {
            self.value_text.trim().parse::<usize>().ok()
        } else {
            None
        };
        let rich_text = match self.data_type {
            CellDataType::SharedString => shared_string_idx
                .and_then(|idx| shared_strings.rich_text.get(idx).cloned().flatten()),
            CellDataType::InlineString if self.inline_saw_r => Some(self.inline_runs),
            _ => None,
        };
        let value = match self.data_type {
            CellDataType::SharedString => shared_string_idx
                .and_then(|i| shared_strings.values.get(i).cloned())
                .map(CellValue::String)
                .unwrap_or(CellValue::Empty),
            CellDataType::InlineString => {
                if self.inline_text.is_empty() {
                    CellValue::Empty
                } else {
                    CellValue::String(self.inline_text)
                }
            }
            CellDataType::Bool => CellValue::Bool(matches!(self.value_text.trim(), "1" | "true")),
            CellDataType::Error => CellValue::Error(self.value_text),
            CellDataType::FormulaString => CellValue::String(self.value_text),
            CellDataType::Number => {
                let raw = self.value_text.trim();
                if raw.is_empty() {
                    CellValue::Empty
                } else {
                    raw.parse::<f64>().map(CellValue::Number).map_err(|e| {
                        ReaderError::Xml(format!("invalid numeric cell {}: {e}", self.coordinate))
                    })?
                }
            }
        };
        let array_formula = match self.formula_kind.as_deref() {
            Some("array") => Some(ArrayFormulaInfo::Array {
                ref_range: self.formula_ref.clone().unwrap_or_default(),
                text: self.formula_text.clone(),
            }),
            Some("dataTable") => Some(ArrayFormulaInfo::DataTable {
                ref_range: self.formula_ref.clone().unwrap_or_default(),
                ca: self.formula_ca,
                dt2_d: self.formula_dt2_d,
                dtr: self.formula_dtr,
                r1: self.formula_r1.clone(),
                r2: self.formula_r2.clone(),
            }),
            _ => None,
        };
        Ok(Cell {
            row: self.row,
            col: self.col,
            coordinate: self.coordinate,
            style_id: self.style_id,
            data_type: self.data_type,
            value,
            formula: if self.formula_text.is_empty() {
                None
            } else {
                Some(self.formula_text)
            },
            formula_kind: self.formula_kind,
            formula_shared_index: self.formula_shared_index,
            array_formula,
            rich_text,
        })
    }
}

#[derive(Debug, Clone, Copy)]
enum TextTarget {
    Value,
    Formula,
    InlineString,
}

fn zip_from_bytes(bytes: &[u8]) -> Result<ZipArchive<Cursor<&[u8]>>> {
    ZipArchive::new(Cursor::new(bytes)).map_err(ReaderError::Zip)
}

fn normalize_ooxml_text(text: &str) -> String {
    text.replace("\r\n", "\n").replace('\r', "\n")
}

fn read_part_required<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> Result<String> {
    read_part_optional(zip, name)?.ok_or_else(|| ReaderError::MissingPart(name.to_string()))
}

fn read_part_optional<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<String>> {
    match zip.by_name(name) {
        Ok(mut file) => {
            let mut out = String::new();
            file.read_to_string(&mut out)?;
            Ok(Some(out))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(e) => Err(ReaderError::Zip(e)),
    }
}

fn read_part_optional_bytes<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<Vec<u8>>> {
    match zip.by_name(name) {
        Ok(mut file) => {
            let mut out = Vec::new();
            file.read_to_end(&mut out)?;
            Ok(Some(out))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(e) => Err(ReaderError::Zip(e)),
    }
}

fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for a in e.attributes().with_checks(false).flatten() {
        if a.key.as_ref() == key {
            if let Ok(v) = a.unescape_value() {
                return Some(v.to_string());
            }
            return Some(String::from_utf8_lossy(a.value.as_ref()).into_owned());
        }
    }
    None
}

fn attr_truthy(value: Option<&str>) -> bool {
    matches!(value, Some(v) if v != "0" && !v.eq_ignore_ascii_case("false"))
}

fn attr_bool_default(e: &BytesStart<'_>, key: &[u8], default: bool) -> bool {
    attr_value(e, key)
        .as_deref()
        .map_or(default, |value| attr_truthy(Some(value)))
}

fn attr_bool(e: &BytesStart<'_>, key: &[u8]) -> Option<bool> {
    attr_value(e, key)
        .as_deref()
        .map(|value| attr_truthy(Some(value)))
}

fn attr_u32(e: &BytesStart<'_>, key: &[u8]) -> Option<u32> {
    attr_value(e, key).and_then(|value| value.parse::<u32>().ok())
}

fn attr_i64(e: &BytesStart<'_>, key: &[u8]) -> Option<i64> {
    attr_value(e, key).and_then(|value| value.parse::<i64>().ok())
}

fn attr_f64(e: &BytesStart<'_>, key: &[u8]) -> Option<f64> {
    attr_value(e, key).and_then(|value| value.parse::<f64>().ok())
}

fn parse_sheet_state(value: Option<&str>) -> SheetState {
    match value {
        Some("hidden") => SheetState::Hidden,
        Some("veryHidden") => SheetState::VeryHidden,
        _ => SheetState::Visible,
    }
}

fn ensure_formula_prefix(formula: &str) -> String {
    if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    }
}

fn normalize_zip_path(path: &str) -> String {
    let mut stack = Vec::new();
    for part in path.split('/') {
        if part.is_empty() || part == "." {
            continue;
        }
        if part == ".." {
            stack.pop();
            continue;
        }
        stack.push(part);
    }
    stack.join("/")
}

fn join_and_normalize(base_dir: &str, target: &str) -> String {
    let t = target.trim_start_matches('/');
    let combined = if t.starts_with("xl/") {
        t.to_string()
    } else {
        format!("{base_dir}{t}")
    };
    normalize_zip_path(&combined)
}

fn part_dir(path: &str) -> String {
    let normalized = normalize_zip_path(path);
    normalized
        .rsplit_once('/')
        .map(|(dir, _)| format!("{dir}/"))
        .unwrap_or_default()
}

fn sheet_rels_path(sheet_path: &str) -> String {
    let normalized = normalize_zip_path(sheet_path);
    let Some((dir, file)) = normalized.rsplit_once('/') else {
        return format!("_rels/{normalized}.rels");
    };
    format!("{dir}/_rels/{file}.rels")
}

fn a1_to_row_col(coord: &str) -> Option<(u32, u32)> {
    let mut col = 0u32;
    let mut row = 0u32;
    let mut saw_digit = false;
    for ch in coord.chars() {
        if ch == '$' {
            continue;
        }
        if ch.is_ascii_alphabetic() && !saw_digit {
            col = col
                .checked_mul(26)?
                .checked_add((ch.to_ascii_uppercase() as u8 - b'A' + 1) as u32)?;
        } else if ch.is_ascii_digit() {
            saw_digit = true;
            row = row.checked_mul(10)?.checked_add(ch.to_digit(10)?)?;
        } else {
            return None;
        }
    }
    if row == 0 || col == 0 {
        None
    } else {
        Some((row, col))
    }
}

fn row_col_to_a1(row: u32, col: u32) -> String {
    let mut c = col;
    let mut letters = Vec::new();
    while c > 0 {
        c -= 1;
        letters.push((b'A' + (c % 26) as u8) as char);
        c /= 26;
    }
    letters.reverse();
    format!("{}{row}", letters.into_iter().collect::<String>())
}

#[cfg(test)]
mod tests {
    use super::*;

    const XLSX_BYTES: &[u8] = include_bytes!("../../../tests/fixtures/sprint_kappa_smoke.xlsx");

    #[test]
    fn opens_committed_xlsx_fixture() {
        let book = NativeXlsxBook::open_bytes(XLSX_BYTES).expect("fixture opens");
        assert!(!book.sheet_names().is_empty());
        let first = book.sheet_names()[0].to_string();
        let sheet = book.worksheet(&first).expect("first sheet parses");
        assert!(!sheet.cells.is_empty(), "fixture should have cells");
    }

    #[test]
    fn parses_workbook_sheet_order_and_state() {
        let xml = r#"<workbook xmlns:r="r">
        <fileSharing readOnlyRecommended="1" userName="Wolf" algorithmName="SHA-512" hashValue="FILEHASH" saltValue="FILESALT" spinCount="100000"/>
        <workbookPr date1904="1" codeName="ThisWorkbook" defaultThemeVersion="164011"/>
        <workbookProtection lockStructure="1" workbookAlgorithmName="SHA-512" workbookHashValue="HASH" workbookSaltValue="SALT" workbookSpinCount="100000"/>
        <calcPr calcId="191029" calcMode="manual" fullCalcOnLoad="1" iterate="1" iterateCount="25" iterateDelta="0.01" forceFullCalc="1"/>
        <sheets>
            <sheet name="Visible" sheetId="1" r:id="rId1"/>
            <sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/>
            <sheet name="Very" sheetId="3" state="veryHidden" r:id="rId3"/>
        </sheets><definedNames>
            <definedName name="GlobalName">Visible!$A$1</definedName>
            <definedName name="LocalName" localSheetId="1">$B$2</definedName>
            <definedName name="_xlnm.Print_Area">Visible!$A$1:$B$2</definedName>
            <definedName name="_xlnm.Print_Titles" localSheetId="0">Visible!$1:$2,Visible!$A:$B</definedName>
        </definedNames></workbook>"#;
        let (
            sheets,
            date1904,
            named_ranges,
            print_areas,
            print_titles,
            security,
            workbook_properties,
            calc_properties,
        ) = parse_workbook(xml).expect("parse workbook");
        assert!(date1904);
        let workbook_properties = workbook_properties.expect("workbookPr");
        assert!(workbook_properties.date1904);
        assert_eq!(
            workbook_properties.code_name.as_deref(),
            Some("ThisWorkbook")
        );
        assert_eq!(workbook_properties.default_theme_version, Some(164011));
        let calc_properties = calc_properties.expect("calcPr");
        assert_eq!(calc_properties.calc_id, Some(191029));
        assert_eq!(calc_properties.calc_mode.as_deref(), Some("manual"));
        assert_eq!(calc_properties.full_calc_on_load, Some(true));
        assert_eq!(calc_properties.iterate_count, Some(25));
        assert_eq!(calc_properties.iterate_delta, Some(0.01));
        assert_eq!(calc_properties.force_full_calc, Some(true));
        assert_eq!(sheets[0].name, "Visible");
        assert_eq!(sheets[1].state, SheetState::Hidden);
        assert_eq!(sheets[2].state, SheetState::VeryHidden);
        assert_eq!(
            security.workbook_protection,
            Some(WorkbookProtection {
                lock_structure: true,
                lock_windows: false,
                lock_revision: false,
                workbook_algorithm_name: Some("SHA-512".to_string()),
                workbook_hash_value: Some("HASH".to_string()),
                workbook_salt_value: Some("SALT".to_string()),
                workbook_spin_count: Some(100000),
                revisions_algorithm_name: None,
                revisions_hash_value: None,
                revisions_salt_value: None,
                revisions_spin_count: None,
            })
        );
        assert_eq!(
            security.file_sharing,
            Some(FileSharing {
                read_only_recommended: true,
                user_name: Some("Wolf".to_string()),
                algorithm_name: Some("SHA-512".to_string()),
                hash_value: Some("FILEHASH".to_string()),
                salt_value: Some("FILESALT".to_string()),
                spin_count: Some(100000),
            })
        );
        assert_eq!(
            named_ranges,
            vec![
                NamedRange {
                    name: "GlobalName".to_string(),
                    scope: "workbook".to_string(),
                    refers_to: "Visible!$A$1".to_string(),
                },
                NamedRange {
                    name: "LocalName".to_string(),
                    scope: "sheet".to_string(),
                    refers_to: "Hidden!$B$2".to_string(),
                },
            ]
        );
        assert_eq!(
            print_areas.get("Visible").map(String::as_str),
            Some("Visible!$A$1:$B$2")
        );
        assert_eq!(
            print_titles.get("Visible"),
            Some(&PrintTitlesInfo {
                rows: Some("1:2".to_string()),
                cols: Some("A:B".to_string()),
            })
        );
    }

    #[test]
    fn parses_shared_strings_plain_and_rich_text() {
        let xml = r#"<sst><si><t>Plain</t></si><si><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>Rich&#13;&#10;</t></r><r><t>Text</t></r></si></sst>"#;
        let strings = parse_shared_strings(xml).expect("parse sst");
        assert_eq!(strings.values, vec!["Plain", "Rich\nText"]);
        assert_eq!(strings.rich_text[0], None);
        let rich = strings.rich_text[1].as_ref().expect("rich runs");
        assert_eq!(
            rich.iter().map(|run| run.text.as_str()).collect::<Vec<_>>(),
            vec!["Rich\n", "Text"]
        );
        assert_eq!(rich[0].font.as_ref().and_then(|font| font.bold), Some(true));
        assert_eq!(
            rich[0].font.as_ref().and_then(|font| font.color.clone()),
            Some("FFFF0000".to_string())
        );
    }

    #[test]
    fn parses_custom_and_builtin_number_formats() {
        let xml = r#"<styleSheet>
            <numFmts count="1"><numFmt numFmtId="165" formatCode="$#,##0.00"/></numFmts>
            <borders count="2">
                <border><left/><right/><top/><bottom/><diagonal/></border>
                <border diagonalUp="1"><left style="thin"><color rgb="FFFF0000"/></left><right style="medium"/><top style="double"/><bottom style="dashed"><color indexed="4"/></bottom><diagonal style="hair"><color rgb="FF00FF00"/></diagonal></border>
            </borders>
            <cellXfs count="3">
                <xf numFmtId="0"/>
                <xf numFmtId="4"/>
                <xf numFmtId="165" borderId="1"/>
            </cellXfs>
        </styleSheet>"#;
        let styles = parse_style_tables(xml).expect("parse styles");
        assert_eq!(styles.number_format_for_style_id(0), None);
        assert_eq!(styles.number_format_for_style_id(1), Some("#,##0.00"));
        assert_eq!(styles.number_format_for_style_id(2), Some("$#,##0.00"));
        let border = styles.border_for_style_id(2).expect("border");
        assert_eq!(
            border.left,
            Some(BorderSide {
                style: "thin".to_string(),
                color: "#FF0000".to_string(),
            })
        );
        assert_eq!(
            border.right,
            Some(BorderSide {
                style: "medium".to_string(),
                color: "#000000".to_string(),
            })
        );
        assert_eq!(
            border.bottom,
            Some(BorderSide {
                style: "dashed".to_string(),
                color: "#0000FF".to_string(),
            })
        );
        assert_eq!(
            border.diagonal_up,
            Some(BorderSide {
                style: "hair".to_string(),
                color: "#00FF00".to_string(),
            })
        );
        assert_eq!(border.diagonal_down, None);
    }

    #[test]
    fn parses_sheet_values_formulas_and_types() {
        let xml = r#"<worksheet><dimension ref="A1:D2"/><sheetViews><sheetView>
            <pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
        </sheetView></sheetViews><cols><col min="2" max="3" width="18.5" customWidth="1" hidden="1" outlineLevel="2"/></cols><sheetData>
            <row r="1">
                <c r="A1" t="s"><v>0</v></c>
                <c r="B1"><v>42</v></c>
                <c r="C1" t="b"><v>1</v></c>
                <c r="D1"><f>SUM(B1:B1)</f><v>42</v></c>
            </row>
            <row r="2" ht="24" customHeight="1" hidden="1" outlineLevel="1"><c r="A2" t="inlineStr"><is><r><rPr><i/></rPr><t>In</t></r><r><t>line</t></r></is></c></row>
        </sheetData><mergeCells count="1"><mergeCell ref="A3:B3"/></mergeCells>
        <dataValidations count="1">
            <dataValidation type="whole" operator="between" allowBlank="1" sqref="B2:B5" errorTitle="Invalid" error="Use 1-10">
                <formula1>1</formula1><formula2>10</formula2>
            </dataValidation>
        </dataValidations>
        <conditionalFormatting sqref="C2:C5">
            <cfRule type="cellIs" operator="greaterThan" priority="1" stopIfTrue="1"><formula>50</formula></cfRule>
        </conditionalFormatting>
        <sheetProtection sheet="1" objects="1" formatCells="0" sort="0" password="C258"/>
        <autoFilter ref="A1:D10">
            <filterColumn colId="0"><filters><filter val="Label"/></filters></filterColumn>
            <filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThan" val="10"/></customFilters></filterColumn>
            <sortState ref="A2:D10"><sortCondition ref="B2:B10" descending="1"/></sortState>
        </autoFilter>
        </worksheet>"#;
        let shared_strings = SharedStrings {
            values: vec!["Shared".to_string()],
            rich_text: vec![None],
        };
        let sheet = parse_worksheet(xml, &shared_strings, None, Vec::new(), Vec::new())
            .expect("parse worksheet");
        assert_eq!(sheet.dimension.as_deref(), Some("A1:D2"));
        assert_eq!(sheet.merged_ranges, vec!["A3:B3"]);
        assert_eq!(
            sheet.freeze_panes,
            Some(FreezePane {
                mode: PaneMode::Freeze,
                top_left_cell: Some("B2".to_string()),
                x_split: Some(1),
                y_split: Some(1),
                active_pane: Some("bottomRight".to_string()),
            })
        );
        assert_eq!(
            sheet.row_heights.get(&2),
            Some(&RowHeight {
                height: 24.0,
                custom_height: true,
            })
        );
        assert_eq!(
            sheet.column_widths,
            vec![ColumnWidth {
                min: 2,
                max: 3,
                width: 18.5,
                custom_width: true,
            }]
        );
        assert_eq!(sheet.hidden_rows, vec![2]);
        assert_eq!(sheet.hidden_columns, vec![2, 3]);
        assert_eq!(sheet.row_outline_levels, vec![(2, 1)]);
        assert_eq!(sheet.column_outline_levels, vec![(2, 2), (3, 2)]);
        assert_eq!(
            sheet.data_validations,
            vec![DataValidation {
                range: "B2:B5".to_string(),
                validation_type: "whole".to_string(),
                operator: Some("between".to_string()),
                formula1: Some("=1".to_string()),
                formula2: Some("=10".to_string()),
                allow_blank: true,
                error_title: Some("Invalid".to_string()),
                error: Some("Use 1-10".to_string()),
            }]
        );
        assert_eq!(
            sheet.conditional_formats,
            vec![ConditionalFormatRule {
                range: "C2:C5".to_string(),
                rule_type: "cellIs".to_string(),
                operator: Some("greaterThan".to_string()),
                formula: Some("=50".to_string()),
                priority: Some(1),
                stop_if_true: Some(true),
            }]
        );
        assert_eq!(
            sheet.sheet_protection,
            Some(SheetProtection {
                sheet: true,
                objects: true,
                scenarios: false,
                format_cells: false,
                format_columns: true,
                format_rows: true,
                insert_columns: true,
                insert_rows: true,
                insert_hyperlinks: true,
                delete_columns: true,
                delete_rows: true,
                select_locked_cells: false,
                sort: false,
                auto_filter: true,
                pivot_tables: true,
                select_unlocked_cells: false,
                password_hash: Some("C258".to_string()),
            })
        );
        assert_eq!(
            sheet.auto_filter,
            Some(AutoFilterInfo {
                ref_range: "A1:D10".to_string(),
                filter_columns: vec![
                    FilterColumnInfo {
                        col_id: 0,
                        hidden_button: false,
                        show_button: true,
                        filter: Some(FilterInfo::String {
                            values: vec!["Label".to_string()],
                        }),
                        date_group_items: Vec::new(),
                    },
                    FilterColumnInfo {
                        col_id: 1,
                        hidden_button: false,
                        show_button: true,
                        filter: Some(FilterInfo::Custom {
                            and_: true,
                            filters: vec![CustomFilterInfo {
                                operator: "greaterThan".to_string(),
                                val: "10".to_string(),
                            }],
                        }),
                        date_group_items: Vec::new(),
                    },
                ],
                sort_state: Some(SortStateInfo {
                    sort_conditions: vec![SortConditionInfo {
                        ref_range: "B2:B10".to_string(),
                        descending: true,
                        sort_by: "value".to_string(),
                        custom_list: None,
                        dxf_id: None,
                        icon_set: None,
                        icon_id: None,
                    }],
                    column_sort: false,
                    case_sensitive: false,
                    ref_range: Some("A2:D10".to_string()),
                }),
            })
        );
        assert_eq!(
            sheet.cells[0].value,
            CellValue::String("Shared".to_string())
        );
        assert_eq!(sheet.cells[1].value, CellValue::Number(42.0));
        assert_eq!(sheet.cells[2].value, CellValue::Bool(true));
        assert_eq!(sheet.cells[3].formula.as_deref(), Some("SUM(B1:B1)"));
        assert_eq!(
            sheet.cells[4].value,
            CellValue::String("Inline".to_string())
        );
        let inline_rich = sheet.cells[4].rich_text.as_ref().expect("inline rich text");
        assert_eq!(
            inline_rich
                .iter()
                .map(|run| run.text.as_str())
                .collect::<Vec<_>>(),
            vec!["In", "line"]
        );
        assert_eq!(
            inline_rich[0].font.as_ref().and_then(|font| font.italic),
            Some(true)
        );
    }

    #[test]
    fn expands_shared_formula_children() {
        let xml = r#"<worksheet><sheetData>
            <row r="1">
                <c r="A1"><v>1</v></c>
                <c r="B1"><f t="shared" si="0" ref="B1:B3">A1*2</f><v/></c>
                <c r="C1"><f t="shared" si="1" ref="C1:C3">$A$1+A1+B$1+$A1</f><v/></c>
            </row>
            <row r="2">
                <c r="A2"><v>2</v></c>
                <c r="B2"><f t="shared" si="0"/><v/></c>
                <c r="C2"><f t="shared" si="1"/><v/></c>
            </row>
            <row r="3">
                <c r="A3"><v>3</v></c>
                <c r="B3"><f t="shared" si="0"/><v/></c>
                <c r="C3"><f t="shared" si="1"/><v/></c>
            </row>
        </sheetData></worksheet>"#;

        let sheet = parse_worksheet(xml, &SharedStrings::default(), None, Vec::new(), Vec::new())
            .expect("parse worksheet");
        let formulas: HashMap<_, _> = sheet
            .cells
            .iter()
            .filter_map(|cell| {
                cell.formula
                    .as_deref()
                    .map(|formula| (cell.coordinate.as_str(), formula))
            })
            .collect();

        assert_eq!(formulas.get("B1"), Some(&"A1*2"));
        assert_eq!(formulas.get("B2"), Some(&"A2*2"));
        assert_eq!(formulas.get("B3"), Some(&"A3*2"));
        assert_eq!(formulas.get("C1"), Some(&"$A$1+A1+B$1+$A1"));
        assert_eq!(formulas.get("C2"), Some(&"$A$1+A2+B$1+$A2"));
        assert_eq!(formulas.get("C3"), Some(&"$A$1+A3+B$1+$A3"));
    }

    #[test]
    fn parses_sheet_hyperlinks_with_relationship_targets() {
        let xml = r#"<worksheet xmlns:r="r"><sheetData><row r="1">
            <c r="A1" t="inlineStr"><is><t>Website</t></is></c>
            <c r="B1" t="inlineStr"><is><t>Internal</t></is></c>
        </row></sheetData><hyperlinks>
            <hyperlink ref="A1" r:id="rId1" tooltip="External site"/>
            <hyperlink ref="B1" location="Other!A1" display="Jump"/>
        </hyperlinks></worksheet>"#;
        let rels = RelsGraph::parse(
            br#"<Relationships>
                <Relationship Id="rId1" Type="hyperlink" Target="https://example.com" TargetMode="External"/>
            </Relationships>"#,
        )
        .expect("parse rels");

        let sheet = parse_worksheet(
            xml,
            &SharedStrings::default(),
            Some(&rels),
            Vec::new(),
            Vec::new(),
        )
        .expect("parse worksheet");
        assert_eq!(
            sheet.hyperlinks,
            vec![
                Hyperlink {
                    cell: "A1".to_string(),
                    target: "https://example.com".to_string(),
                    display: "Website".to_string(),
                    tooltip: Some("External site".to_string()),
                    internal: false,
                },
                Hyperlink {
                    cell: "B1".to_string(),
                    target: "Other!A1".to_string(),
                    display: "Jump".to_string(),
                    tooltip: None,
                    internal: true,
                },
            ]
        );
    }

    #[test]
    fn parses_comments_authors_and_rich_text() {
        let xml = r#"<comments><authors><author>Alice</author><author>Bob</author></authors>
            <commentList>
                <comment ref="A1" authorId="0"><text><t>First note</t></text></comment>
                <comment ref="B2" authorId="1"><text><r><t>Second</t></r><r><t> note</t></r></text></comment>
            </commentList></comments>"#;

        assert_eq!(
            parse_comments(xml).expect("parse comments"),
            vec![
                Comment {
                    cell: "A1".to_string(),
                    text: "First note".to_string(),
                    author: "Alice".to_string(),
                    threaded: false,
                },
                Comment {
                    cell: "B2".to_string(),
                    text: "Second note".to_string(),
                    author: "Bob".to_string(),
                    threaded: false,
                },
            ]
        );
    }

    #[test]
    fn parses_table_metadata() {
        let xml = r#"<table name="SalesTable" displayName="SalesTable" ref="A1:B3" headerRowCount="1" totalsRowCount="0">
            <autoFilter ref="A1:B3"/>
            <tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Sales"/></tableColumns>
            <tableStyleInfo name="TableStyleLight9"/>
        </table>"#;

        assert_eq!(
            parse_table_xml(xml).expect("parse table"),
            Table {
                name: "SalesTable".to_string(),
                ref_range: "A1:B3".to_string(),
                header_row: true,
                totals_row: false,
                style: Some("TableStyleLight9".to_string()),
                columns: vec!["Name".to_string(), "Sales".to_string()],
                autofilter: true,
            }
        );
    }

    #[test]
    fn parses_document_properties() {
        let mut props = HashMap::new();
        parse_doc_properties_into(
            r#"<cp:coreProperties>
                <dc:title>Q3 Report</dc:title>
                <dc:creator>Alice</dc:creator>
                <cp:lastModifiedBy>Bob</cp:lastModifiedBy>
                <dcterms:created>2024-01-01T00:00:00Z</dcterms:created>
            </cp:coreProperties>"#,
            &mut props,
            doc_property_core_key,
        );
        parse_doc_properties_into(
            r#"<Properties><Application>Excel</Application><Company>SynthGL</Company></Properties>"#,
            &mut props,
            doc_property_app_key,
        );

        assert_eq!(props.get("title").map(String::as_str), Some("Q3 Report"));
        assert_eq!(props.get("creator").map(String::as_str), Some("Alice"));
        assert_eq!(props.get("lastModifiedBy").map(String::as_str), Some("Bob"));
        assert_eq!(
            props.get("created").map(String::as_str),
            Some("2024-01-01T00:00:00Z")
        );
        assert_eq!(props.get("application").map(String::as_str), Some("Excel"));
        assert_eq!(props.get("company").map(String::as_str), Some("SynthGL"));
    }
}
