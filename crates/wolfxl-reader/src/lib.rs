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
use std::path::Path;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
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
    pub comments: Vec<Comment>,
    pub row_heights: HashMap<u32, RowHeight>,
    pub column_widths: Vec<ColumnWidth>,
    pub data_validations: Vec<DataValidation>,
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

/// OOXML pane mode relevant to read compatibility.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PaneMode {
    Freeze,
    Split,
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

/// Native XLSX/XLSM workbook reader.
#[derive(Debug, Clone)]
pub struct NativeXlsxBook {
    bytes: Vec<u8>,
    sheets: Vec<SheetInfo>,
    shared_strings: Vec<String>,
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
        let (sheet_refs, date1904) = parse_workbook(&workbook_xml)?;
        let sheets = resolve_sheet_paths(sheet_refs, &rels)?;
        let shared_strings = match read_part_optional(&mut zip, "xl/sharedStrings.xml")? {
            Some(xml) => parse_shared_strings(&xml)?,
            None => Vec::new(),
        };
        let styles = match read_part_optional(&mut zip, "xl/styles.xml")? {
            Some(xml) => parse_style_tables(&xml)?,
            None => StyleTables::default(),
        };

        Ok(Self {
            bytes,
            sheets,
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

    /// Whether the workbook uses the 1904 date system.
    pub fn date1904(&self) -> bool {
        self.date1904
    }

    /// Shared-string table as plain strings.
    pub fn shared_strings(&self) -> &[String] {
        &self.shared_strings
    }

    /// Resolve a style id to an Excel number format code.
    pub fn number_format_for_style_id(&self, style_id: u32) -> Option<&str> {
        self.styles.number_format_for_style_id(style_id)
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
        parse_worksheet(&xml, &self.shared_strings, rels.as_ref(), comments)
    }
}

#[derive(Debug)]
struct SheetRef {
    name: String,
    sheet_id: Option<String>,
    state: SheetState,
    rid: String,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct StyleTables {
    custom_num_fmts: HashMap<u32, String>,
    cell_xfs: Vec<XfEntry>,
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
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct XfEntry {
    num_fmt_id: u32,
}

fn parse_workbook(xml: &str) -> Result<(Vec<SheetRef>, bool)> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut sheets = Vec::new();
    let mut date1904 = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"workbookPr" => {
                    date1904 = attr_truthy(attr_value(&e, b"date1904").as_deref());
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
                _ => {}
            },
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

    Ok((sheets, date1904))
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

fn parse_shared_strings(xml: &str) -> Result<Vec<String>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    let mut current = String::new();
    let mut in_si = false;
    let mut in_t = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"si" => {
                    in_si = true;
                    current.clear();
                }
                b"t" => {
                    if in_si {
                        in_t = true;
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"si" => {
                    out.push(std::mem::take(&mut current));
                    in_si = false;
                }
                b"t" => in_t = false,
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_si && in_t {
                    current.push_str(
                        &e.unescape().map_err(|err| {
                            ReaderError::Xml(format!("shared string text: {err}"))
                        })?,
                    );
                }
            }
            Ok(Event::CData(e)) => {
                if in_si && in_t {
                    current.push_str(&String::from_utf8_lossy(e.as_ref()));
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

fn parse_style_tables(xml: &str) -> Result<StyleTables> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut styles = StyleTables::default();
    let mut in_num_fmts = false;
    let mut in_cell_xfs = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"numFmts" => in_num_fmts = true,
                b"cellXfs" => in_cell_xfs = true,
                b"numFmt" if in_num_fmts => {
                    push_num_fmt(&mut styles, &e);
                }
                b"xf" if in_cell_xfs => {
                    styles.cell_xfs.push(parse_xf_entry(&e));
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"numFmt" if in_num_fmts => push_num_fmt(&mut styles, &e),
                b"xf" if in_cell_xfs => styles.cell_xfs.push(parse_xf_entry(&e)),
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"numFmts" => in_num_fmts = false,
                b"cellXfs" => in_cell_xfs = false,
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
    }
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
    shared_strings: &[String],
    rels: Option<&RelsGraph>,
    comments: Vec<Comment>,
) -> Result<WorksheetData> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut dimension = None;
    let mut merged_ranges = Vec::new();
    let mut hyperlink_nodes = Vec::new();
    let mut freeze_panes = None;
    let mut row_heights = HashMap::new();
    let mut column_widths = Vec::new();
    let mut data_validations = Vec::new();
    let mut row_index: Option<u32> = None;
    let mut current: Option<CellBuilder> = None;
    let mut active_text: Option<TextTarget> = None;
    let mut current_validation: Option<DataValidationBuilder> = None;
    let mut active_validation_text: Option<DataValidationFormula> = None;
    let mut cells = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
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
                    freeze_panes = parse_pane(&e);
                }
                b"row" => {
                    row_index = attr_value(&e, b"r").and_then(|v| v.parse::<u32>().ok());
                    if let Some((row, height)) = parse_row_height(&e, row_index) {
                        row_heights.insert(row, height);
                    }
                }
                b"c" => {
                    current = Some(CellBuilder::from_start(&e, row_index));
                }
                b"dataValidation" => {
                    current_validation = Some(DataValidationBuilder::from_start(&e));
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
                b"f" => active_text = Some(TextTarget::Formula),
                b"t" => {
                    if current
                        .as_ref()
                        .is_some_and(|c| c.data_type == CellDataType::InlineString)
                    {
                        active_text = Some(TextTarget::InlineString);
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"col" => {
                    if let Some(width) = parse_column_width(&e) {
                        column_widths.push(width);
                    }
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
                    freeze_panes = parse_pane(&e);
                }
                b"c" => {
                    let builder = CellBuilder::from_start(&e, row_index);
                    cells.push(builder.finish(shared_strings)?);
                }
                b"row" => {
                    let row = attr_value(&e, b"r").and_then(|v| v.parse::<u32>().ok());
                    if let Some((row, height)) = parse_row_height(&e, row) {
                        row_heights.insert(row, height);
                    }
                }
                b"dataValidation" => {
                    let validation = DataValidationBuilder::from_start(&e).finish();
                    if !validation.range.trim().is_empty() {
                        data_validations.push(validation);
                    }
                }
                _ => {}
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
                b"v" | b"f" | b"t" => active_text = None,
                b"row" => row_index = None,
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if let (Some(target), Some(validation)) =
                    (active_validation_text, current_validation.as_mut())
                {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("worksheet text: {err}")))?;
                    validation.push_text(target, &text);
                } else if let (Some(target), Some(cell)) = (active_text, current.as_mut()) {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("worksheet text: {err}")))?;
                    cell.push_text(target, &text);
                }
            }
            Ok(Event::CData(e)) => {
                if let (Some(target), Some(validation)) =
                    (active_validation_text, current_validation.as_mut())
                {
                    validation.push_text(target, &String::from_utf8_lossy(e.as_ref()));
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

    Ok(WorksheetData {
        dimension,
        merged_ranges,
        hyperlinks: resolve_hyperlinks(hyperlink_nodes, rels, &cells),
        freeze_panes,
        comments,
        row_heights,
        column_widths,
        data_validations,
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
        }
    }

    fn push_text(&mut self, target: TextTarget, text: &str) {
        match target {
            TextTarget::Value => self.value_text.push_str(text),
            TextTarget::Formula => self.formula_text.push_str(text),
            TextTarget::InlineString => self.inline_text.push_str(text),
        }
    }

    fn finish(self, shared_strings: &[String]) -> Result<Cell> {
        let value = match self.data_type {
            CellDataType::SharedString => {
                let idx = self.value_text.trim().parse::<usize>().ok();
                idx.and_then(|i| shared_strings.get(i).cloned())
                    .map(CellValue::String)
                    .unwrap_or(CellValue::Empty)
            }
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
        let xml = r#"<workbook xmlns:r="r"><workbookPr date1904="1"/><sheets>
            <sheet name="Visible" sheetId="1" r:id="rId1"/>
            <sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/>
            <sheet name="Very" sheetId="3" state="veryHidden" r:id="rId3"/>
        </sheets></workbook>"#;
        let (sheets, date1904) = parse_workbook(xml).expect("parse workbook");
        assert!(date1904);
        assert_eq!(sheets[0].name, "Visible");
        assert_eq!(sheets[1].state, SheetState::Hidden);
        assert_eq!(sheets[2].state, SheetState::VeryHidden);
    }

    #[test]
    fn parses_shared_strings_plain_and_rich_text() {
        let xml =
            r#"<sst><si><t>Plain</t></si><si><r><t>Rich</t></r><r><t> Text</t></r></si></sst>"#;
        let strings = parse_shared_strings(xml).expect("parse sst");
        assert_eq!(strings, vec!["Plain", "Rich Text"]);
    }

    #[test]
    fn parses_custom_and_builtin_number_formats() {
        let xml = r#"<styleSheet>
            <numFmts count="1"><numFmt numFmtId="165" formatCode="$#,##0.00"/></numFmts>
            <cellXfs count="3">
                <xf numFmtId="0"/>
                <xf numFmtId="4"/>
                <xf numFmtId="165"/>
            </cellXfs>
        </styleSheet>"#;
        let styles = parse_style_tables(xml).expect("parse styles");
        assert_eq!(styles.number_format_for_style_id(0), None);
        assert_eq!(styles.number_format_for_style_id(1), Some("#,##0.00"));
        assert_eq!(styles.number_format_for_style_id(2), Some("$#,##0.00"));
    }

    #[test]
    fn parses_sheet_values_formulas_and_types() {
        let xml = r#"<worksheet><dimension ref="A1:D2"/><sheetViews><sheetView>
            <pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
        </sheetView></sheetViews><cols><col min="2" max="3" width="18.5" customWidth="1"/></cols><sheetData>
            <row r="1">
                <c r="A1" t="s"><v>0</v></c>
                <c r="B1"><v>42</v></c>
                <c r="C1" t="b"><v>1</v></c>
                <c r="D1"><f>SUM(B1:B1)</f><v>42</v></c>
            </row>
            <row r="2" ht="24" customHeight="1"><c r="A2" t="inlineStr"><is><t>Inline</t></is></c></row>
        </sheetData><mergeCells count="1"><mergeCell ref="A3:B3"/></mergeCells>
        <dataValidations count="1">
            <dataValidation type="whole" operator="between" allowBlank="1" sqref="B2:B5" errorTitle="Invalid" error="Use 1-10">
                <formula1>1</formula1><formula2>10</formula2>
            </dataValidation>
        </dataValidations></worksheet>"#;
        let sheet = parse_worksheet(xml, &["Shared".to_string()], None, Vec::new())
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

        let sheet = parse_worksheet(xml, &[], Some(&rels), Vec::new()).expect("parse worksheet");
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
}
