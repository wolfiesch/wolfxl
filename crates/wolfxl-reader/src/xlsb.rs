use std::collections::HashMap;
use std::fs;
use std::io::{Cursor, Read};
use std::path::Path;

use wolfxl_rels::{RelId, RelsGraph};
use zip::ZipArchive;

use crate::{
    row_col_to_a1, AlignmentInfo, AutoFilterInfo, BorderInfo, Cell, CellDataType, CellValue,
    ColumnWidth, Comment, ConditionalFormatRule, DataValidation, FillInfo, FontInfo, FreezePane,
    HeaderFooterInfo, Hyperlink, ImageInfo, PageBreakListInfo, PageMarginsInfo, PageSetupInfo,
    RowHeight, SheetFormatInfo, SheetPropertiesInfo, SheetProtection, SheetState, SheetViewInfo,
    StyleTables, Table, WorksheetData, XfEntry,
};

type Result<T> = std::result::Result<T, XlsbError>;

#[derive(Debug)]
pub enum XlsbError {
    Io(std::io::Error),
    Zip(zip::result::ZipError),
    Xml(String),
    Format(String),
    SheetNotFound(String),
}

impl std::fmt::Display for XlsbError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(e) => write!(f, "I/O error: {e}"),
            Self::Zip(e) => write!(f, "ZIP error: {e}"),
            Self::Xml(e) => write!(f, "XML error: {e}"),
            Self::Format(e) => write!(f, "XLSB format error: {e}"),
            Self::SheetNotFound(name) => write!(f, "sheet not found: {name}"),
        }
    }
}

impl std::error::Error for XlsbError {}

impl From<std::io::Error> for XlsbError {
    fn from(value: std::io::Error) -> Self {
        Self::Io(value)
    }
}

impl From<zip::result::ZipError> for XlsbError {
    fn from(value: zip::result::ZipError) -> Self {
        Self::Zip(value)
    }
}

#[derive(Debug, Clone)]
pub struct NativeXlsbBook {
    bytes: Vec<u8>,
    sheets: Vec<XlsbSheet>,
    shared_strings: Vec<String>,
    styles: StyleTables,
    date1904: bool,
}

#[derive(Debug, Clone)]
struct XlsbSheet {
    name: String,
    path: String,
    state: SheetState,
}

#[derive(Debug)]
struct Record<'a> {
    typ: u16,
    payload: &'a [u8],
}

impl NativeXlsbBook {
    pub fn open_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_bytes(fs::read(path)?)
    }

    pub fn open_bytes(bytes: Vec<u8>) -> Result<Self> {
        let mut zip = ZipArchive::new(Cursor::new(bytes.clone()))?;
        let workbook_rels = read_rels(&mut zip, "xl/_rels/workbook.bin.rels")?;
        let shared_strings = read_shared_strings(&mut zip)?;
        let styles = read_styles(&mut zip)?;
        let (sheets, date1904) = read_workbook(&mut zip, &workbook_rels)?;
        Ok(Self {
            bytes,
            sheets,
            shared_strings,
            styles,
            date1904,
        })
    }

    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name.as_str()).collect()
    }

    pub fn sheet_state(&self, sheet_name: &str) -> Result<SheetState> {
        self.sheets
            .iter()
            .find(|s| s.name == sheet_name)
            .map(|s| s.state)
            .ok_or_else(|| XlsbError::SheetNotFound(sheet_name.to_string()))
    }

    pub fn date1904(&self) -> bool {
        self.date1904
    }

    pub fn number_format_for_style_id(&self, style_id: u32) -> Option<&str> {
        self.styles.number_format_for_style_id(style_id)
    }

    pub fn border_for_style_id(&self, style_id: u32) -> Option<&BorderInfo> {
        self.styles.border_for_style_id(style_id)
    }

    pub fn font_for_style_id(&self, style_id: u32) -> Option<&FontInfo> {
        self.styles.font_for_style_id(style_id)
    }

    pub fn fill_for_style_id(&self, style_id: u32) -> Option<&FillInfo> {
        self.styles.fill_for_style_id(style_id)
    }

    pub fn alignment_for_style_id(&self, style_id: u32) -> Option<&AlignmentInfo> {
        self.styles.alignment_for_style_id(style_id)
    }

    pub fn worksheet(&self, sheet_name: &str) -> Result<WorksheetData> {
        let sheet = self
            .sheets
            .iter()
            .find(|s| s.name == sheet_name)
            .ok_or_else(|| XlsbError::SheetNotFound(sheet_name.to_string()))?;
        let mut zip = ZipArchive::new(Cursor::new(self.bytes.clone()))?;
        let data = read_zip_part(&mut zip, &sheet.path)?;
        Ok(parse_worksheet(&data, &self.shared_strings)?)
    }
}

fn read_rels(zip: &mut ZipArchive<Cursor<Vec<u8>>>, path: &str) -> Result<RelsGraph> {
    let xml = read_zip_part(zip, path)?;
    RelsGraph::parse(&xml).map_err(|e| XlsbError::Xml(format!("failed to parse {path}: {e}")))
}

fn read_zip_part(zip: &mut ZipArchive<Cursor<Vec<u8>>>, path: &str) -> Result<Vec<u8>> {
    let name = zip_part_name(zip, path).unwrap_or_else(|| path.to_string());
    let mut file = zip.by_name(&name)?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)?;
    Ok(bytes)
}

fn read_zip_part_optional(
    zip: &mut ZipArchive<Cursor<Vec<u8>>>,
    path: &str,
) -> Result<Option<Vec<u8>>> {
    let Some(name) = zip_part_name(zip, path) else {
        return Ok(None);
    };
    match zip.by_name(&name) {
        Ok(mut file) => {
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes)?;
            Ok(Some(bytes))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(e) => Err(e.into()),
    }
}

fn zip_part_name(zip: &mut ZipArchive<Cursor<Vec<u8>>>, path: &str) -> Option<String> {
    if zip.by_name(path).is_ok() {
        return Some(path.to_string());
    }
    let needle = path.to_ascii_lowercase();
    zip.file_names()
        .find(|name| name.to_ascii_lowercase() == needle)
        .map(str::to_string)
}

fn read_workbook(
    zip: &mut ZipArchive<Cursor<Vec<u8>>>,
    rels: &RelsGraph,
) -> Result<(Vec<XlsbSheet>, bool)> {
    let data = read_zip_part(zip, "xl/workbook.bin")?;
    let mut sheets = Vec::new();
    let mut date1904 = false;
    for record in Records::new(&data) {
        let record = record?;
        match record.typ {
            0x0099 => {
                date1904 = record
                    .payload
                    .first()
                    .is_some_and(|value| value & 0x01 != 0);
            }
            0x009c => {
                if record.payload.len() < 16 {
                    continue;
                }
                let state = match le_u32(&record.payload[0..4]) {
                    1 => SheetState::Hidden,
                    2 => SheetState::VeryHidden,
                    _ => SheetState::Visible,
                };
                let rel_len = le_u32(&record.payload[8..12]) as usize;
                let rel_start = 12;
                let rel_end = rel_start + rel_len * 2;
                if rel_end > record.payload.len() {
                    continue;
                }
                let rid = utf16_string(&record.payload[rel_start..rel_end]);
                let mut consumed = 0;
                let name = wide_string(&record.payload[rel_end..], &mut consumed)?;
                let target = rels
                    .get(&RelId(rid.clone()))
                    .ok_or_else(|| XlsbError::Format(format!("missing workbook rel {rid}")))?;
                let path = join_xl_target(&target.target);
                sheets.push(XlsbSheet { name, path, state });
            }
            _ => {}
        }
    }
    Ok((sheets, date1904))
}

fn read_shared_strings(zip: &mut ZipArchive<Cursor<Vec<u8>>>) -> Result<Vec<String>> {
    let Some(data) = read_zip_part_optional(zip, "xl/sharedStrings.bin")? else {
        return Ok(Vec::new());
    };
    let mut strings = Vec::new();
    for record in Records::new(&data) {
        let record = record?;
        if record.typ == 0x0013 && record.payload.len() > 1 {
            let mut consumed = 0;
            strings.push(wide_string(&record.payload[1..], &mut consumed)?);
        }
    }
    Ok(strings)
}

fn read_styles(zip: &mut ZipArchive<Cursor<Vec<u8>>>) -> Result<StyleTables> {
    let Some(data) = read_zip_part_optional(zip, "xl/styles.bin")? else {
        return Ok(StyleTables::default());
    };
    let mut styles = StyleTables::default();
    let mut in_cell_xfs = false;
    for record in Records::new(&data) {
        let record = record?;
        match record.typ {
            0x0269 => in_cell_xfs = true,
            0x026a => in_cell_xfs = false,
            0x002c => {
                if record.payload.len() >= 6 {
                    let id = le_u16(&record.payload[0..2]) as u32;
                    let mut consumed = 0;
                    let code = wide_string(&record.payload[2..], &mut consumed)?;
                    styles.custom_num_fmts.insert(id, code);
                }
            }
            0x002b => styles.fonts.push(parse_font(record.payload)),
            0x002d => styles.fills.push(parse_fill(record.payload)),
            0x002e => styles.borders.push(parse_border(record.payload)),
            0x002f if in_cell_xfs => styles.cell_xfs.push(parse_xf(record.payload)),
            _ => {}
        }
    }
    if styles.cell_xfs.is_empty() {
        styles.cell_xfs.push(XfEntry::default());
    }
    Ok(styles)
}

fn parse_worksheet(data: &[u8], shared_strings: &[String]) -> Result<WorksheetData> {
    let mut cells = Vec::new();
    let mut dimension = None;
    let mut row_heights = HashMap::new();
    let mut column_widths = Vec::new();
    let mut merged_ranges = Vec::new();
    let mut hidden_rows = Vec::new();
    let mut hidden_columns = Vec::new();
    let mut row_outline_levels = Vec::new();
    let mut column_outline_levels = Vec::new();
    let mut row = 0u32;
    for record in Records::new(data) {
        let record = record?;
        match record.typ {
            0x0094 => {
                dimension = parse_ws_dimension(record.payload);
            }
            0x0000 => {
                if record.payload.len() >= 4 {
                    row = le_u32(&record.payload[0..4]);
                }
                if let Some(info) = parse_row_header(record.payload) {
                    if info.hidden {
                        hidden_rows.push(info.row);
                    }
                    if info.outline_level > 0 {
                        row_outline_levels.push((info.row, info.outline_level));
                    }
                    if let Some(height) = info.height {
                        row_heights.insert(info.row, height);
                    }
                }
            }
            0x0001 => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::Empty,
                        CellDataType::Number,
                        None,
                    ));
                }
            }
            0x0002 => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let value = parse_rk(record.payload.get(8..12).unwrap_or(&[]));
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::Number(value),
                        CellDataType::Number,
                        None,
                    ));
                }
            }
            0x0003 | 0x000b => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let err = record
                        .payload
                        .get(8)
                        .copied()
                        .map(error_code)
                        .unwrap_or("#ERROR!");
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::Error(err.to_string()),
                        CellDataType::Error,
                        None,
                    ));
                }
            }
            0x0004 | 0x000a => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let value = record.payload.get(8).copied().unwrap_or_default() != 0;
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::Bool(value),
                        CellDataType::Bool,
                        None,
                    ));
                }
            }
            0x0005 | 0x0009 => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let value = record.payload.get(8..16).map(le_f64).unwrap_or_default();
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::Number(value),
                        CellDataType::Number,
                        None,
                    ));
                }
            }
            0x0006 | 0x0008 => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let mut consumed = 0;
                    let value = wide_string(record.payload.get(8..).unwrap_or(&[]), &mut consumed)?;
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::String(value),
                        CellDataType::InlineString,
                        None,
                    ));
                }
            }
            0x0007 => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let idx = record.payload.get(8..12).map(le_u32).unwrap_or_default() as usize;
                    let value = shared_strings.get(idx).cloned().unwrap_or_default();
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::String(value),
                        CellDataType::SharedString,
                        None,
                    ));
                }
            }
            0x003c => {
                if let Some(info) = parse_column_info(record.payload) {
                    if info.width.custom_width {
                        column_widths.push(info.width);
                    }
                    for col in info.width.min..=info.width.max {
                        if info.hidden {
                            hidden_columns.push(col);
                        }
                        if info.outline_level > 0 {
                            column_outline_levels.push((col, info.outline_level));
                        }
                    }
                }
            }
            0x00b0 => {
                if let Some(range) = parse_merged_range(record.payload) {
                    merged_ranges.push(range);
                }
            }
            _ => {}
        }
    }
    hidden_rows.sort_unstable();
    hidden_rows.dedup();
    hidden_columns.sort_unstable();
    hidden_columns.dedup();
    row_outline_levels.sort_unstable_by_key(|(row, _)| *row);
    column_outline_levels.sort_unstable_by_key(|(col, _)| *col);
    Ok(worksheet_data(
        dimension.or_else(|| cells_dimension(&cells)),
        row_heights,
        column_widths,
        hidden_rows,
        hidden_columns,
        row_outline_levels,
        column_outline_levels,
        merged_ranges,
        cells,
    ))
}

#[derive(Debug, Clone, Copy, PartialEq)]
struct RowHeaderInfo {
    row: u32,
    hidden: bool,
    outline_level: u8,
    height: Option<RowHeight>,
}

#[derive(Debug, Clone, Copy, PartialEq)]
struct ColumnInfo {
    width: ColumnWidth,
    hidden: bool,
    outline_level: u8,
}

fn parse_ws_dimension(payload: &[u8]) -> Option<String> {
    if payload.len() < 16 {
        return None;
    }
    let min_row = le_u32(&payload[0..4]) + 1;
    let max_row = le_u32(&payload[4..8]) + 1;
    let min_col = le_u32(&payload[8..12]) + 1;
    let max_col = le_u32(&payload[12..16]) + 1;
    if max_row < min_row || max_col < min_col {
        return None;
    }
    Some(format!(
        "{}:{}",
        row_col_to_a1(min_row, min_col),
        row_col_to_a1(max_row, max_col)
    ))
}

fn parse_row_header(payload: &[u8]) -> Option<RowHeaderInfo> {
    if payload.len() < 12 {
        return None;
    }
    let row = le_u32(&payload[0..4]) + 1;
    let height_twips = le_u16(&payload[8..10]);
    let flags = le_u16(&payload[10..12]);
    let outline_level = ((flags >> 8) & 0x07) as u8;
    let hidden = flags & 0x1000 != 0;
    let custom_height = flags & 0x2000 != 0;
    let height = custom_height.then_some(RowHeight {
        height: height_twips as f64 / 20.0,
        custom_height: true,
    });
    Some(RowHeaderInfo {
        row,
        hidden,
        outline_level,
        height,
    })
}

fn parse_column_info(payload: &[u8]) -> Option<ColumnInfo> {
    if payload.len() < 18 {
        return None;
    }
    let min = le_u32(&payload[0..4]) + 1;
    let max = le_u32(&payload[4..8]) + 1;
    let width_raw = le_u32(&payload[8..12]);
    let flags = le_u16(&payload[16..18]);
    Some(ColumnInfo {
        width: ColumnWidth {
            min,
            max,
            width: width_raw as f64 / 256.0,
            custom_width: flags & 0x0002 != 0,
        },
        hidden: flags & 0x0001 != 0,
        outline_level: ((flags >> 8) & 0x07) as u8,
    })
}

fn parse_merged_range(payload: &[u8]) -> Option<String> {
    if payload.len() < 16 {
        return None;
    }
    let first_row = le_u32(&payload[0..4]) + 1;
    let last_row = le_u32(&payload[4..8]) + 1;
    let first_col = le_u32(&payload[8..12]) + 1;
    let last_col = le_u32(&payload[12..16]) + 1;
    if last_row < first_row || last_col < first_col {
        return None;
    }
    Some(format!(
        "{}:{}",
        row_col_to_a1(first_row, first_col),
        row_col_to_a1(last_row, last_col)
    ))
}

fn parse_cell_header(payload: &[u8]) -> Option<(u32, Option<u32>)> {
    if payload.len() < 8 {
        return None;
    }
    let col = le_u32(&payload[0..4]);
    let style_id = le_u32(&payload[4..8]);
    Some((col, (style_id != 0).then_some(style_id)))
}

fn make_cell(
    row0: u32,
    col0: u32,
    style_id: Option<u32>,
    value: CellValue,
    data_type: CellDataType,
    formula: Option<String>,
) -> Cell {
    let row = row0 + 1;
    let col = col0 + 1;
    Cell {
        row,
        col,
        coordinate: row_col_to_a1(row, col),
        style_id,
        data_type,
        value,
        formula,
        formula_kind: None,
        formula_shared_index: None,
        array_formula: None,
        rich_text: None,
    }
}

fn worksheet_data(
    dimension: Option<String>,
    row_heights: HashMap<u32, RowHeight>,
    column_widths: Vec<ColumnWidth>,
    hidden_rows: Vec<u32>,
    hidden_columns: Vec<u32>,
    row_outline_levels: Vec<(u32, u8)>,
    column_outline_levels: Vec<(u32, u8)>,
    merged_ranges: Vec<String>,
    cells: Vec<Cell>,
) -> WorksheetData {
    WorksheetData {
        dimension,
        merged_ranges,
        hyperlinks: Vec::<Hyperlink>::new(),
        freeze_panes: None::<FreezePane>,
        sheet_properties: None::<SheetPropertiesInfo>,
        sheet_view: None::<SheetViewInfo>,
        comments: Vec::<Comment>::new(),
        row_heights,
        column_widths,
        data_validations: Vec::<DataValidation>::new(),
        sheet_protection: None::<SheetProtection>,
        auto_filter: None::<AutoFilterInfo>,
        page_margins: None::<PageMarginsInfo>,
        page_setup: None::<PageSetupInfo>,
        header_footer: None::<HeaderFooterInfo>,
        row_breaks: None::<PageBreakListInfo>,
        column_breaks: None::<PageBreakListInfo>,
        sheet_format: None::<SheetFormatInfo>,
        images: Vec::<ImageInfo>::new(),
        charts: Vec::new(),
        tables: Vec::<Table>::new(),
        conditional_formats: Vec::<ConditionalFormatRule>::new(),
        hidden_rows,
        hidden_columns,
        row_outline_levels,
        column_outline_levels,
        array_formulas: HashMap::new(),
        cells,
    }
}

fn cells_dimension(cells: &[Cell]) -> Option<String> {
    let mut bounds: Option<(u32, u32, u32, u32)> = None;
    for cell in cells {
        bounds = match bounds {
            Some((min_r, min_c, max_r, max_c)) => Some((
                min_r.min(cell.row),
                min_c.min(cell.col),
                max_r.max(cell.row),
                max_c.max(cell.col),
            )),
            None => Some((cell.row, cell.col, cell.row, cell.col)),
        };
    }
    bounds.map(|(min_r, min_c, max_r, max_c)| {
        format!(
            "{}:{}",
            row_col_to_a1(min_r, min_c),
            row_col_to_a1(max_r, max_c)
        )
    })
}

fn parse_xf(payload: &[u8]) -> XfEntry {
    XfEntry {
        num_fmt_id: payload.get(2..4).map(le_u16).unwrap_or_default() as u32,
        font_id: payload.get(4..6).map(le_u16).unwrap_or_default() as u32,
        border_id: payload.get(6..8).map(le_u16).unwrap_or_default() as u32,
        fill_id: payload.get(8..10).map(le_u16).unwrap_or_default() as u32,
        alignment: parse_binary_alignment(payload),
    }
}

fn parse_font(payload: &[u8]) -> FontInfo {
    let mut font = FontInfo::default();
    if payload.len() >= 2 {
        let size_twips = le_u16(&payload[0..2]);
        if size_twips > 0 {
            font.size = Some(size_twips as f64 / 20.0);
        }
    }
    if let Some((_, name)) = find_trailing_wide_string(payload) {
        if !name.is_empty() {
            font.name = Some(name);
        }
    }
    font
}

fn parse_fill(payload: &[u8]) -> FillInfo {
    let _ = payload;
    FillInfo::default()
}

fn parse_border(payload: &[u8]) -> BorderInfo {
    let _ = payload;
    BorderInfo::default()
}

fn parse_binary_alignment(payload: &[u8]) -> Option<AlignmentInfo> {
    let flags = payload.get(12..16).map(le_u32).unwrap_or_default();
    let wrap_text = flags & 0x20 != 0;
    wrap_text.then_some(AlignmentInfo {
        horizontal: None,
        vertical: None,
        wrap_text,
        text_rotation: None,
        indent: None,
    })
}

fn parse_rk(bytes: &[u8]) -> f64 {
    if bytes.len() < 4 {
        return 0.0;
    }
    let mut raw = [0u8; 4];
    raw.copy_from_slice(&bytes[..4]);
    let flags = raw[0] & 0x03;
    raw[0] &= 0xfc;
    let mut value = if flags & 0x02 != 0 {
        i32::from_le_bytes(raw) as f64 / 4.0
    } else {
        let mut eight = [0u8; 8];
        eight[4..8].copy_from_slice(&raw);
        f64::from_le_bytes(eight)
    };
    if flags & 0x01 != 0 {
        value /= 100.0;
    }
    value
}

fn find_trailing_wide_string(payload: &[u8]) -> Option<(usize, String)> {
    for offset in (0..payload.len().saturating_sub(4)).rev() {
        let len = le_u32(&payload[offset..offset + 4]) as usize;
        let end = offset + 4 + len * 2;
        if len > 0 && end == payload.len() {
            return Some((offset, utf16_string(&payload[offset + 4..end])));
        }
    }
    None
}

fn wide_string(bytes: &[u8], consumed: &mut usize) -> Result<String> {
    if bytes.len() < 4 {
        return Err(XlsbError::Format(
            "truncated wide string length".to_string(),
        ));
    }
    let len = le_u32(&bytes[0..4]) as usize;
    let end = 4 + len * 2;
    if bytes.len() < end {
        return Err(XlsbError::Format(
            "truncated wide string payload".to_string(),
        ));
    }
    *consumed = end;
    Ok(utf16_string(&bytes[4..end]))
}

fn utf16_string(bytes: &[u8]) -> String {
    let units: Vec<u16> = bytes.chunks_exact(2).map(le_u16).collect();
    String::from_utf16_lossy(&units)
}

fn join_xl_target(target: &str) -> String {
    let target = target.trim_start_matches('/');
    if target.starts_with("xl/") {
        target.to_string()
    } else {
        format!("xl/{target}")
    }
}

fn error_code(code: u8) -> &'static str {
    match code {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0f => "#VALUE!",
        0x17 => "#REF!",
        0x1d => "#NAME?",
        0x24 => "#NUM!",
        0x2a => "#N/A",
        _ => "#ERROR!",
    }
}

fn le_u16(bytes: &[u8]) -> u16 {
    u16::from_le_bytes([bytes[0], bytes[1]])
}

fn le_u32(bytes: &[u8]) -> u32 {
    u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]])
}

fn le_f64(bytes: &[u8]) -> f64 {
    f64::from_le_bytes([
        bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
    ])
}

struct Records<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> Records<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, offset: 0 }
    }
}

impl<'a> Iterator for Records<'a> {
    type Item = Result<Record<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.offset >= self.bytes.len() {
            return None;
        }
        let typ = match read_record_type(self.bytes, &mut self.offset) {
            Ok(value) => value,
            Err(e) => return Some(Err(e)),
        };
        let len = match read_record_len(self.bytes, &mut self.offset) {
            Ok(value) => value,
            Err(e) => return Some(Err(e)),
        };
        let end = self.offset.saturating_add(len);
        if end > self.bytes.len() {
            return Some(Err(XlsbError::Format(
                "truncated record payload".to_string(),
            )));
        }
        let payload = &self.bytes[self.offset..end];
        self.offset = end;
        Some(Ok(Record { typ, payload }))
    }
}

fn read_record_type(bytes: &[u8], offset: &mut usize) -> Result<u16> {
    let first = read_byte(bytes, offset)?;
    let mut typ = (first & 0x7f) as u16;
    if first & 0x80 != 0 {
        let second = read_byte(bytes, offset)?;
        typ += ((second & 0x7f) as u16) << 7;
    }
    Ok(typ)
}

fn read_record_len(bytes: &[u8], offset: &mut usize) -> Result<usize> {
    let mut shift = 0;
    let mut len = 0usize;
    loop {
        let byte = read_byte(bytes, offset)?;
        len += ((byte & 0x7f) as usize) << shift;
        if byte & 0x80 == 0 {
            break;
        }
        shift += 7;
        if shift > 21 {
            return Err(XlsbError::Format(
                "record length varint too long".to_string(),
            ));
        }
    }
    Ok(len)
}

fn read_byte(bytes: &[u8], offset: &mut usize) -> Result<u8> {
    let Some(value) = bytes.get(*offset).copied() else {
        return Err(XlsbError::Format("unexpected end of records".to_string()));
    };
    *offset += 1;
    Ok(value)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn put_u16(out: &mut Vec<u8>, value: u16) {
        out.extend_from_slice(&value.to_le_bytes());
    }

    fn put_u32(out: &mut Vec<u8>, value: u32) {
        out.extend_from_slice(&value.to_le_bytes());
    }

    fn push_record(out: &mut Vec<u8>, typ: u16, payload: &[u8]) {
        if typ < 0x80 {
            out.push(typ as u8);
        } else {
            out.push(((typ & 0x7f) as u8) | 0x80);
            out.push((typ >> 7) as u8);
        }
        out.push(payload.len() as u8);
        out.extend_from_slice(payload);
    }

    #[test]
    fn parses_xlsb_row_header_visibility_outline_and_height() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 4); // zero-based row 5
        put_u32(&mut payload, 0);
        put_u16(&mut payload, 420); // 21pt
        put_u16(&mut payload, (3 << 8) | 0x1000 | 0x2000);

        assert_eq!(
            parse_row_header(&payload),
            Some(RowHeaderInfo {
                row: 5,
                hidden: true,
                outline_level: 3,
                height: Some(RowHeight {
                    height: 21.0,
                    custom_height: true,
                }),
            })
        );
    }

    #[test]
    fn parses_xlsb_column_info_visibility_outline_and_width() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 1); // zero-based B
        put_u32(&mut payload, 2); // zero-based C
        put_u32(&mut payload, 12 * 256);
        put_u32(&mut payload, 0);
        put_u16(&mut payload, 0x0001 | 0x0002 | (2 << 8));

        assert_eq!(
            parse_column_info(&payload),
            Some(ColumnInfo {
                width: ColumnWidth {
                    min: 2,
                    max: 3,
                    width: 12.0,
                    custom_width: true,
                },
                hidden: true,
                outline_level: 2,
            })
        );
    }

    #[test]
    fn parses_xlsb_worksheet_dimension_and_merged_range_as_inclusive_bounds() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 3); // B4
        put_u32(&mut payload, 6); // B7
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 1);
        assert_eq!(parse_ws_dimension(&payload), Some("B4:B7".to_string()));

        let mut merge = Vec::new();
        put_u32(&mut merge, 0); // A1
        put_u32(&mut merge, 1); // B2
        put_u32(&mut merge, 0);
        put_u32(&mut merge, 1);
        assert_eq!(parse_merged_range(&merge), Some("A1:B2".to_string()));
    }

    #[test]
    fn parse_worksheet_collects_xlsb_sheet_structure_metadata() {
        let mut data = Vec::new();
        let mut row = Vec::new();
        put_u32(&mut row, 4);
        put_u32(&mut row, 0);
        put_u16(&mut row, 420);
        put_u16(&mut row, (1 << 8) | 0x1000 | 0x2000);
        put_u32(&mut row, 0);
        push_record(&mut data, 0x0000, &row);

        let mut cell = Vec::new();
        put_u32(&mut cell, 1);
        put_u32(&mut cell, 0);
        cell.extend_from_slice(&42.0f64.to_le_bytes());
        push_record(&mut data, 0x0005, &cell);

        let mut col = Vec::new();
        put_u32(&mut col, 1);
        put_u32(&mut col, 1);
        put_u32(&mut col, 20 * 256);
        put_u32(&mut col, 0);
        put_u16(&mut col, 0x0001 | 0x0002 | (2 << 8));
        push_record(&mut data, 0x003c, &col);

        let mut merge = Vec::new();
        put_u32(&mut merge, 4);
        put_u32(&mut merge, 4);
        put_u32(&mut merge, 1);
        put_u32(&mut merge, 2);
        push_record(&mut data, 0x00b0, &merge);

        let sheet = parse_worksheet(&data, &[]).expect("parse synthetic worksheet");
        assert_eq!(sheet.merged_ranges, vec!["B5:C5"]);
        assert_eq!(sheet.hidden_rows, vec![5]);
        assert_eq!(sheet.row_outline_levels, vec![(5, 1)]);
        assert_eq!(
            sheet.row_heights.get(&5),
            Some(&RowHeight {
                height: 21.0,
                custom_height: true,
            })
        );
        assert_eq!(sheet.hidden_columns, vec![2]);
        assert_eq!(sheet.column_outline_levels, vec![(2, 2)]);
        assert_eq!(
            sheet.column_widths,
            vec![ColumnWidth {
                min: 2,
                max: 2,
                width: 20.0,
                custom_width: true,
            }]
        );
    }
}
