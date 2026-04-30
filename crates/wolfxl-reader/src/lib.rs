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
    pub cells: Vec<Cell>,
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
        parse_worksheet(&xml, &self.shared_strings)
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

fn parse_worksheet(xml: &str, shared_strings: &[String]) -> Result<WorksheetData> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut dimension = None;
    let mut merged_ranges = Vec::new();
    let mut row_index: Option<u32> = None;
    let mut current: Option<CellBuilder> = None;
    let mut active_text: Option<TextTarget> = None;
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
                b"row" => {
                    row_index = attr_value(&e, b"r").and_then(|v| v.parse::<u32>().ok());
                }
                b"c" => {
                    current = Some(CellBuilder::from_start(&e, row_index));
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
                b"dimension" => {
                    dimension = attr_value(&e, b"ref");
                }
                b"mergeCell" => {
                    if let Some(range) = attr_value(&e, b"ref") {
                        merged_ranges.push(range);
                    }
                }
                b"c" => {
                    let builder = CellBuilder::from_start(&e, row_index);
                    cells.push(builder.finish(shared_strings)?);
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"c" => {
                    if let Some(builder) = current.take() {
                        cells.push(builder.finish(shared_strings)?);
                    }
                }
                b"v" | b"f" | b"t" => active_text = None,
                b"row" => row_index = None,
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if let (Some(target), Some(cell)) = (active_text, current.as_mut()) {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("worksheet text: {err}")))?;
                    cell.push_text(target, &text);
                }
            }
            Ok(Event::CData(e)) => {
                if let (Some(target), Some(cell)) = (active_text, current.as_mut()) {
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
        cells,
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
        let xml = r#"<worksheet><dimension ref="A1:D2"/><sheetData>
            <row r="1">
                <c r="A1" t="s"><v>0</v></c>
                <c r="B1"><v>42</v></c>
                <c r="C1" t="b"><v>1</v></c>
                <c r="D1"><f>SUM(B1:B1)</f><v>42</v></c>
            </row>
            <row r="2"><c r="A2" t="inlineStr"><is><t>Inline</t></is></c></row>
        </sheetData><mergeCells count="1"><mergeCell ref="A3:B3"/></mergeCells></worksheet>"#;
        let sheet = parse_worksheet(xml, &["Shared".to_string()]).expect("parse worksheet");
        assert_eq!(sheet.dimension.as_deref(), Some("A1:D2"));
        assert_eq!(sheet.merged_ranges, vec!["A3:B3"]);
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
}
