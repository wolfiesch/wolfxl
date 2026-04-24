//! Worksheet data — rows, cells, merges, freeze, and all sheet-scope features.

use std::collections::BTreeMap;

use super::cell::WriteCell;
use super::comment::Comment;
use super::conditional::ConditionalFormat;
use super::table::Table;
use super::validation::DataValidation;

/// A single worksheet within a workbook.
///
/// `BTreeMap` row/cell keys are deliberate: OOXML requires rows inside
/// `<sheetData>` to be sorted ascending by `r`, and cells inside each
/// `<row>` to be sorted ascending by column letter. Using `BTreeMap`
/// means the emitter iterates them in the right order without an
/// explicit pre-sort pass.
#[derive(Debug, Clone)]
pub struct Worksheet {
    /// Sheet display name. Maximum 31 chars; `/\?*[]:` must be stripped
    /// before reaching here. See [`crate::refs::sanitize_sheet_name`].
    pub name: String,

    /// Sparse row storage, keyed by 1-based row index.
    pub rows: BTreeMap<u32, Row>,

    /// Ranges of cells merged into one visual cell.
    pub merges: Vec<Merge>,

    /// The freeze or split pane configuration, if any.
    pub freeze: Option<FreezePane>,
    pub split: Option<SplitPane>,

    /// Column metadata: widths, hidden flags, style defaults.
    /// Keyed by 1-based column index.
    pub columns: BTreeMap<u32, Column>,

    /// Cell-scope hyperlinks. Keyed by A1 reference (e.g. `"A1"`).
    pub hyperlinks: BTreeMap<String, Hyperlink>,

    /// Cell comments. Keyed by A1 reference. Author insertion order
    /// is preserved by the `IndexMap` inside the workbook-level emitter.
    pub comments: BTreeMap<String, Comment>,

    /// Conditional formatting blocks. Order is preserved (Excel honors
    /// first-matching priority).
    pub conditional_formats: Vec<ConditionalFormat>,

    /// Data validation rules. Order is preserved.
    pub validations: Vec<DataValidation>,

    /// Tables (`<table>` OOXML, not HTML tables) — named structured ranges.
    pub tables: Vec<Table>,

    /// Print area, stored as a range reference string (e.g. `"A1:D20"`).
    /// Written both into `<definedNames>` at workbook scope and into the
    /// sheet's `<pageSetup>`-adjacent blocks.
    pub print_area: Option<String>,

    /// Whether the sheet tab is visible, hidden, or very-hidden.
    pub visibility: SheetVisibility,
}

impl Worksheet {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            rows: BTreeMap::new(),
            merges: Vec::new(),
            freeze: None,
            split: None,
            columns: BTreeMap::new(),
            hyperlinks: BTreeMap::new(),
            comments: BTreeMap::new(),
            conditional_formats: Vec::new(),
            validations: Vec::new(),
            tables: Vec::new(),
            print_area: None,
            visibility: SheetVisibility::Visible,
        }
    }

    /// Set a cell by 1-based row/column. Any row in between is left untouched.
    pub fn set_cell(&mut self, row: u32, col: u32, cell: WriteCell) {
        self.rows.entry(row).or_default().cells.insert(col, cell);
    }

    pub fn set_row_height(&mut self, row: u32, height: f64) {
        self.rows.entry(row).or_default().custom_height = Some(height);
    }

    pub fn set_column(&mut self, col: u32, column: Column) {
        self.columns.insert(col, column);
    }

    pub fn merge(&mut self, range: Merge) {
        self.merges.push(range);
    }
}

/// One row's metadata and its cells.
#[derive(Debug, Clone, Default)]
pub struct Row {
    /// If set, the row is emitted with `ht="…" customHeight="1"`.
    pub custom_height: Option<f64>,

    /// If set, the row is hidden (`hidden="1"`).
    pub hidden: bool,

    /// Optional style_id for whole-row default style.
    pub style_id: Option<u32>,

    /// Sparse cell storage, keyed by 1-based column index.
    pub cells: BTreeMap<u32, WriteCell>,
}

/// One column's metadata. Excel stores column widths per-range, but the
/// most common API shape is per-column, so we normalize to that and let
/// the emitter coalesce adjacent identical ranges if desired.
#[derive(Debug, Clone, Default)]
pub struct Column {
    /// Column width in "max-digit-width" units (Excel's oddball measure).
    /// Pass-through whatever the Python caller provides.
    pub width: Option<f64>,
    pub hidden: bool,
    pub style_id: Option<u32>,
    /// Outline-grouping level (used by Data → Group).
    pub outline_level: u8,
}

/// A rectangular range of cells merged into one visual cell.
///
/// `top_row` and `bottom_row` are 1-based; `left_col` and `right_col` are
/// 1-based. For a single cell, all four values are equal — but OOXML
/// allows that and Excel treats it as a no-op.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Merge {
    pub top_row: u32,
    pub left_col: u32,
    pub bottom_row: u32,
    pub right_col: u32,
}

/// Freeze-pane configuration: rows above `freeze_row` and columns left
/// of `freeze_col` stay fixed while the user scrolls.
///
/// Pure split panes (resizable dividers, no freeze) use [`SplitPane`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FreezePane {
    /// 1-based row; 0/1 means no horizontal freeze.
    pub freeze_row: u32,
    /// 1-based column; 0/1 means no vertical freeze.
    pub freeze_col: u32,
    /// The top-left cell shown in the bottom-right pane.
    /// If `None`, defaults to `(freeze_row, freeze_col)`.
    pub top_left: Option<(u32, u32)>,
}

/// Split-pane configuration (draggable dividers, no freeze).
///
/// Values are in "twentieths of a point" per the OOXML spec. This is
/// rare enough that callers usually just use [`FreezePane`] instead.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct SplitPane {
    pub x_split: f64,
    pub y_split: f64,
    pub top_left: Option<(u32, u32)>,
}

/// A hyperlink pointing from a cell (or range) to a target URL or location.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hyperlink {
    /// Either an external URL or an internal workbook reference
    /// (e.g. `"#Sheet2!A1"`).
    pub target: String,
    pub display: Option<String>,
    pub tooltip: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SheetVisibility {
    #[default]
    Visible,
    Hidden,
    /// Very-hidden sheets can only be un-hidden via the Excel VBA editor.
    VeryHidden,
}
