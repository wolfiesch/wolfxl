use std::collections::HashMap;
use std::fs;
use std::io::{Cursor, Read};
use std::path::Path;

use wolfxl_rels::{RelId, RelsGraph};
use zip::ZipArchive;

use crate::{
    row_col_to_a1, AlignmentInfo, AutoFilterInfo, BorderInfo, Cell, CellDataType, CellValue,
    ColumnWidth, Comment, ConditionalFormatRule, DataValidation, FillInfo, FontInfo, FreezePane,
    HeaderFooterInfo, Hyperlink, ImageInfo, NamedRange, PageBreakListInfo, PageMarginsInfo,
    PageSetupInfo, PaneMode, PrintTitlesInfo, RowHeight, SelectionInfo, SheetFormatInfo,
    SheetPropertiesInfo, SheetProtection, SheetState, SheetViewInfo, StyleTables, Table,
    WorksheetData, XfEntry,
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
    named_ranges: Vec<NamedRange>,
    print_areas: HashMap<String, String>,
    print_titles: HashMap<String, PrintTitlesInfo>,
    extern_sheets: Vec<XtiRef>,
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct XtiRef {
    first_sheet: i32,
    last_sheet: i32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RawNamedRange {
    name: String,
    local_id: Option<usize>,
    refers_to: String,
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
        let (sheets, named_ranges, print_areas, print_titles, extern_sheets, date1904) =
            read_workbook(&mut zip, &workbook_rels)?;
        Ok(Self {
            bytes,
            sheets,
            named_ranges,
            print_areas,
            print_titles,
            extern_sheets,
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

    pub fn named_ranges(&self) -> &[NamedRange] {
        &self.named_ranges
    }

    pub fn print_area(&self, sheet_name: &str) -> Option<&str> {
        self.print_areas.get(sheet_name).map(String::as_str)
    }

    pub fn print_titles(&self, sheet_name: &str) -> Option<&PrintTitlesInfo> {
        self.print_titles.get(sheet_name)
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
        let rels = read_zip_part_optional(&mut zip, &sheet_rels_path(&sheet.path))?
            .map(|xml| {
                RelsGraph::parse(&xml).map_err(|e| {
                    XlsbError::Xml(format!("failed to parse sheet rels for {sheet_name}: {e}"))
                })
            })
            .transpose()?;
        let comments = match rels.as_ref().and_then(comments_target) {
            Some(target) => read_zip_part_optional(
                &mut zip,
                &join_and_normalize(&part_dir(&sheet.path), &target),
            )?
            .map(|data| parse_comments_bin(&data))
            .transpose()?
            .unwrap_or_default(),
            None => Vec::new(),
        };
        let context = FormulaContext {
            sheets: &self.sheets,
            extern_sheets: &self.extern_sheets,
            named_ranges: &self.named_ranges,
        };
        Ok(parse_worksheet(
            &data,
            &self.shared_strings,
            rels.as_ref(),
            comments,
            Some(&context),
        )?)
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

fn sheet_rels_path(sheet_path: &str) -> String {
    let normalized = sheet_path.trim_start_matches('/');
    let (dir, file) = match normalized.rsplit_once('/') {
        Some((dir, file)) => (dir, file),
        None => ("", normalized),
    };
    if dir.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{dir}/_rels/{file}.rels")
    }
}

fn part_dir(path: &str) -> String {
    path.trim_start_matches('/')
        .rsplit_once('/')
        .map(|(dir, _)| dir.to_string())
        .unwrap_or_default()
}

fn join_and_normalize(base_dir: &str, target: &str) -> String {
    let target = target.trim_start_matches('/');
    let raw = if target.starts_with("xl/") {
        target.to_string()
    } else if base_dir.is_empty() {
        target.to_string()
    } else {
        format!("{base_dir}/{target}")
    };
    let mut parts = Vec::new();
    for part in raw.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            value => parts.push(value),
        }
    }
    parts.join("/")
}

fn read_workbook(
    zip: &mut ZipArchive<Cursor<Vec<u8>>>,
    rels: &RelsGraph,
) -> Result<(
    Vec<XlsbSheet>,
    Vec<NamedRange>,
    HashMap<String, String>,
    HashMap<String, PrintTitlesInfo>,
    Vec<XtiRef>,
    bool,
)> {
    let data = read_zip_part(zip, "xl/workbook.bin")?;
    let mut sheets = Vec::new();
    let mut raw_names = Vec::new();
    let mut raw_print_areas = Vec::new();
    let mut raw_print_titles = Vec::new();
    let mut extern_sheets = Vec::new();
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
            0x016a => {
                extern_sheets = parse_extern_sheet(record.payload);
            }
            0x0027 => {
                if let Some(raw_name) = parse_defined_name(record.payload, &sheets, &extern_sheets)
                {
                    if raw_name.name == "_xlnm.Print_Area" {
                        raw_print_areas.push(raw_name);
                    } else if raw_name.name == "_xlnm.Print_Titles" {
                        raw_print_titles.push(raw_name);
                    } else if !raw_name.name.starts_with("_xlnm.") {
                        raw_names.push(raw_name);
                    }
                }
            }
            _ => {}
        }
    }
    let named_ranges = resolve_xlsb_named_ranges(&sheets, raw_names);
    let print_areas = resolve_xlsb_print_areas(&sheets, raw_print_areas);
    let print_titles = resolve_xlsb_print_titles(&sheets, raw_print_titles);
    Ok((
        sheets,
        named_ranges,
        print_areas,
        print_titles,
        extern_sheets,
        date1904,
    ))
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

fn parse_extern_sheet(payload: &[u8]) -> Vec<XtiRef> {
    let Some(count) = payload.get(0..4).map(le_u32) else {
        return Vec::new();
    };
    payload
        .get(4..)
        .unwrap_or_default()
        .chunks_exact(12)
        .take(count as usize)
        .map(|xti| XtiRef {
            first_sheet: le_i32(&xti[4..8]),
            last_sheet: le_i32(&xti[8..12]),
        })
        .collect()
}

fn parse_defined_name(
    payload: &[u8],
    sheets: &[XlsbSheet],
    extern_sheets: &[XtiRef],
) -> Option<RawNamedRange> {
    if payload.len() < 13 {
        return None;
    }
    let itab = le_u32(&payload[5..9]);
    let local_id = (itab != u32::MAX).then_some(itab as usize);
    let mut name_len = 0;
    let name = wide_string(&payload[9..], &mut name_len).ok()?;
    let formula_offset = 9 + name_len;
    let refers_to = parse_name_formula(payload.get(formula_offset..)?, sheets, extern_sheets)?;
    if refers_to.is_empty() {
        return None;
    }
    Some(RawNamedRange {
        name,
        local_id,
        refers_to,
    })
}

fn parse_name_formula(
    payload: &[u8],
    sheets: &[XlsbSheet],
    extern_sheets: &[XtiRef],
) -> Option<String> {
    let rgce_len = payload.get(0..4).map(le_u32)? as usize;
    let rgce = payload.get(4..4 + rgce_len)?;
    let context = FormulaContext {
        sheets,
        extern_sheets,
        named_ranges: &[],
    };
    parse_formula_rgce_with_context(rgce, Some(&context))
}

fn resolve_xlsb_named_ranges(
    sheets: &[XlsbSheet],
    raw_names: Vec<RawNamedRange>,
) -> Vec<NamedRange> {
    raw_names
        .into_iter()
        .map(|raw| {
            let (scope, sheet_name) = match raw.local_id {
                Some(index) => (
                    "sheet".to_string(),
                    sheets.get(index).map(|sheet| sheet.name.clone()),
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

fn resolve_xlsb_print_areas(
    sheets: &[XlsbSheet],
    raw_print_areas: Vec<RawNamedRange>,
) -> HashMap<String, String> {
    let mut out = HashMap::new();
    for raw in raw_print_areas {
        let sheet_name = raw
            .local_id
            .and_then(|index| sheets.get(index).map(|sheet| sheet.name.clone()))
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

fn resolve_xlsb_print_titles(
    sheets: &[XlsbSheet],
    raw_print_titles: Vec<RawNamedRange>,
) -> HashMap<String, PrintTitlesInfo> {
    let mut out = HashMap::new();
    for raw in raw_print_titles {
        let sheet_name = raw
            .local_id
            .and_then(|index| sheets.get(index).map(|sheet| sheet.name.clone()))
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

fn parse_worksheet(
    data: &[u8],
    shared_strings: &[String],
    rels: Option<&RelsGraph>,
    comments: Vec<Comment>,
    formula_context: Option<&FormulaContext<'_>>,
) -> Result<WorksheetData> {
    let mut cells = Vec::new();
    let mut hyperlink_nodes = Vec::new();
    let mut dimension = None;
    let mut row_heights = HashMap::new();
    let mut column_widths = Vec::new();
    let mut merged_ranges = Vec::new();
    let mut hidden_rows = Vec::new();
    let mut hidden_columns = Vec::new();
    let mut row_outline_levels = Vec::new();
    let mut column_outline_levels = Vec::new();
    let mut freeze_panes = None;
    let mut sheet_view = None;
    let mut current_sheet_view: Option<SheetViewInfo> = None;
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
                    let formula =
                        parse_formula_from_cell_record(record.typ, record.payload, formula_context);
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::Error(err.to_string()),
                        CellDataType::Error,
                        formula,
                    ));
                }
            }
            0x0004 | 0x000a => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let value = record.payload.get(8).copied().unwrap_or_default() != 0;
                    let formula =
                        parse_formula_from_cell_record(record.typ, record.payload, formula_context);
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::Bool(value),
                        CellDataType::Bool,
                        formula,
                    ));
                }
            }
            0x0005 | 0x0009 => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let value = record.payload.get(8..16).map(le_f64).unwrap_or_default();
                    let formula =
                        parse_formula_from_cell_record(record.typ, record.payload, formula_context);
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::Number(value),
                        CellDataType::Number,
                        formula,
                    ));
                }
            }
            0x0006 | 0x0008 => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let mut consumed = 0;
                    let value = wide_string(record.payload.get(8..).unwrap_or(&[]), &mut consumed)?;
                    let formula =
                        parse_formula_from_cell_record(record.typ, record.payload, formula_context);
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::String(value),
                        CellDataType::InlineString,
                        formula,
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
            0x01ee => {
                if let Some(node) = parse_hyperlink(record.payload) {
                    hyperlink_nodes.push(node);
                }
            }
            0x0089 => {
                if sheet_view.is_none() {
                    current_sheet_view = parse_sheet_view(record.payload);
                }
            }
            0x008a => {
                if sheet_view.is_none() {
                    sheet_view = current_sheet_view.take();
                } else {
                    current_sheet_view = None;
                }
            }
            0x0097 => {
                if let Some(pane) = parse_pane(record.payload) {
                    if freeze_panes.is_none() {
                        freeze_panes = Some(pane.clone());
                    }
                    if let Some(view) = current_sheet_view.as_mut() {
                        view.pane = Some(pane);
                    }
                }
            }
            0x0098 => {
                if let (Some(view), Some(selection)) =
                    (current_sheet_view.as_mut(), parse_selection(record.payload))
                {
                    view.selections.push(selection);
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
        freeze_panes,
        sheet_view,
        resolve_hyperlinks(hyperlink_nodes, rels, &cells),
        comments,
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

fn parse_pane(payload: &[u8]) -> Option<FreezePane> {
    if payload.len() < 29 {
        return None;
    }
    let x_split = le_f64(&payload[0..8]);
    let y_split = le_f64(&payload[8..16]);
    let top_row = le_u32(&payload[16..20]) + 1;
    let left_col = le_u32(&payload[20..24]) + 1;
    let active_pane = active_pane_name(le_u32(&payload[24..28]));
    let flags = payload[28];
    let frozen = flags & 0x03 != 0;
    Some(FreezePane {
        mode: if frozen {
            PaneMode::Freeze
        } else {
            PaneMode::Split
        },
        top_left_cell: Some(row_col_to_a1(top_row, left_col)),
        x_split: Some(x_split.round() as i64),
        y_split: Some(y_split.round() as i64),
        active_pane,
    })
}

fn active_pane_name(pnn: u32) -> Option<String> {
    let name = match pnn {
        0 => "bottomRight",
        1 => "topRight",
        2 => "bottomLeft",
        3 => "topLeft",
        _ => return None,
    };
    Some(name.to_string())
}

fn parse_sheet_view(payload: &[u8]) -> Option<SheetViewInfo> {
    if payload.len() < 30 {
        return None;
    }
    let flags = le_u16(&payload[0..2]);
    let view = match le_u32(&payload[2..6]) {
        1 => "pageBreakPreview",
        2 => "pageLayout",
        _ => "normal",
    };
    let top_row = le_u32(&payload[6..10]) + 1;
    let left_col = le_u32(&payload[10..14]) + 1;
    let top_left_cell = (top_row != 1 || left_col != 1).then(|| row_col_to_a1(top_row, left_col));
    Some(SheetViewInfo {
        zoom_scale: le_u16(&payload[18..20]) as u32,
        zoom_scale_normal: match le_u16(&payload[20..22]) as u32 {
            0 => 100,
            value => value,
        },
        view: view.to_string(),
        show_grid_lines: flags & 0x0004 != 0,
        show_row_col_headers: flags & 0x0008 != 0,
        show_outline_symbols: flags & 0x0100 != 0,
        show_zeros: flags & 0x0010 != 0,
        right_to_left: flags & 0x0020 != 0,
        tab_selected: flags & 0x0040 != 0,
        top_left_cell,
        workbook_view_id: le_u32(&payload[26..30]),
        pane: None,
        selections: Vec::new(),
    })
}

fn parse_selection(payload: &[u8]) -> Option<SelectionInfo> {
    if payload.len() < 20 {
        return None;
    }
    let pane = active_pane_name(le_u32(&payload[0..4]));
    let active_row = le_u32(&payload[4..8]) + 1;
    let active_col = le_u32(&payload[8..12]) + 1;
    let active_cell_id = le_u32(&payload[12..16]);
    let active_cell = row_col_to_a1(active_row, active_col);
    Some(SelectionInfo {
        pane,
        active_cell: Some(active_cell.clone()),
        sqref: parse_sqref(payload.get(16..)?).or(Some(active_cell)),
        active_cell_id: Some(active_cell_id),
    })
}

fn parse_sqref(payload: &[u8]) -> Option<String> {
    let count = payload.get(0..4).map(le_u32)? as usize;
    if count == 0 {
        return None;
    }
    let ranges: Vec<String> = payload
        .get(4..)?
        .chunks_exact(16)
        .take(count)
        .filter_map(parse_rfx)
        .collect();
    (!ranges.is_empty()).then(|| ranges.join(" "))
}

fn parse_rfx(payload: &[u8]) -> Option<String> {
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
    let first = row_col_to_a1(first_row, first_col);
    let last = row_col_to_a1(last_row, last_col);
    if first == last {
        Some(first)
    } else {
        Some(format!("{first}:{last}"))
    }
}

#[derive(Debug)]
struct HyperlinkNode {
    cell: String,
    rid: Option<String>,
    location: Option<String>,
    display: Option<String>,
    tooltip: Option<String>,
}

fn parse_hyperlink(payload: &[u8]) -> Option<HyperlinkNode> {
    let cell = parse_rfx(payload.get(0..16)?)?;
    let mut offset = 16;
    let rid = read_wide_string_at(payload, &mut offset).filter(|value| !value.is_empty());
    let location = read_wide_string_at(payload, &mut offset).filter(|value| !value.is_empty());
    let tooltip = read_wide_string_at(payload, &mut offset).filter(|value| !value.is_empty());
    let display = read_wide_string_at(payload, &mut offset).filter(|value| !value.is_empty());
    Some(HyperlinkNode {
        cell,
        rid,
        location,
        display,
        tooltip,
    })
}

fn read_wide_string_at(payload: &[u8], offset: &mut usize) -> Option<String> {
    let mut consumed = 0;
    let value = wide_string(payload.get(*offset..)?, &mut consumed).ok()?;
    *offset += consumed;
    Some(value)
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
        out.push(Hyperlink {
            cell: node.cell,
            target,
            display,
            tooltip: node.tooltip,
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

fn parse_comments_bin(data: &[u8]) -> Result<Vec<Comment>> {
    let mut authors = Vec::new();
    let mut comments = Vec::new();
    let mut current: Option<PendingComment> = None;

    for record in Records::new(data) {
        let record = record?;
        match record.typ {
            0x0278 => {
                let mut consumed = 0;
                authors.push(wide_string(record.payload, &mut consumed)?);
            }
            0x027b => {
                current =
                    parse_comment_header(record.payload).map(|(author_id, cell)| PendingComment {
                        cell,
                        author_id,
                        text: String::new(),
                    });
            }
            0x027d => {
                if let Some(comment) = current.as_mut() {
                    comment.text = parse_rich_string(record.payload)?;
                }
            }
            0x027c => {
                if let Some(comment) = current.take() {
                    comments.push(Comment {
                        cell: comment.cell,
                        text: comment.text,
                        author: authors.get(comment.author_id).cloned().unwrap_or_default(),
                        threaded: false,
                    });
                }
            }
            _ => {}
        }
    }
    Ok(comments)
}

#[derive(Debug)]
struct PendingComment {
    cell: String,
    author_id: usize,
    text: String,
}

fn parse_comment_header(payload: &[u8]) -> Option<(usize, String)> {
    let author_id = payload.get(0..4).map(le_u32)? as usize;
    let cell = parse_rfx(payload.get(4..20)?)?;
    Some((author_id, cell))
}

fn parse_rich_string(payload: &[u8]) -> Result<String> {
    let mut consumed = 0;
    if payload.len() > 1 {
        if let Ok(value) = wide_string(&payload[1..], &mut consumed) {
            return Ok(value);
        }
    }
    wide_string(payload, &mut consumed)
}

fn parse_formula_from_cell_record(
    record_type: u16,
    payload: &[u8],
    context: Option<&FormulaContext<'_>>,
) -> Option<String> {
    let formula_payload = match record_type {
        0x0008 => {
            let len = payload.get(8..12).map(le_u32)? as usize;
            payload.get(14 + len * 2..)?
        }
        0x0009 => payload.get(18..)?,
        0x000a | 0x000b => payload.get(11..)?,
        _ => return None,
    };
    let rgce_len = formula_payload.get(0..4).map(le_u32)? as usize;
    let rgce = formula_payload.get(4..4 + rgce_len)?;
    parse_formula_rgce_with_context(rgce, context)
}

struct FormulaContext<'a> {
    sheets: &'a [XlsbSheet],
    extern_sheets: &'a [XtiRef],
    named_ranges: &'a [NamedRange],
}

#[cfg(test)]
fn parse_formula_rgce(rgce: &[u8]) -> Option<String> {
    parse_formula_rgce_with_context(rgce, None)
}

fn parse_formula_rgce_with_context(
    mut rgce: &[u8],
    context: Option<&FormulaContext<'_>>,
) -> Option<String> {
    if rgce.is_empty() {
        return Some(String::new());
    }
    let mut formula = String::with_capacity(rgce.len());
    let mut stack: Vec<usize> = Vec::new();
    while !rgce.is_empty() {
        let ptg = rgce[0];
        rgce = &rgce[1..];
        match ptg {
            0x03..=0x11 => apply_binary_formula_op(ptg, &mut formula, &mut stack)?,
            0x12 => {
                let start = *stack.last()?;
                formula.insert(start, '+');
            }
            0x13 => {
                let start = *stack.last()?;
                formula.insert(start, '-');
            }
            0x14 => formula.push('%'),
            0x15 => {
                let start = *stack.last()?;
                formula.insert(start, '(');
                formula.push(')');
            }
            0x16 => stack.push(formula.len()),
            0x17 => {
                let len = rgce.get(0..2).map(le_u16)? as usize;
                let text = utf16_string(rgce.get(2..2 + len * 2)?);
                stack.push(formula.len());
                formula.push('"');
                formula.push_str(&text);
                formula.push('"');
                rgce = rgce.get(2 + len * 2..)?;
            }
            0x19 => {
                let eptg = *rgce.first()?;
                rgce = rgce.get(1..)?;
                match eptg {
                    0x01 | 0x02 | 0x08 | 0x20 | 0x21 | 0x40 | 0x41 | 0x80 => {
                        rgce = rgce.get(2..)?;
                    }
                    0x04 => rgce = rgce.get(10..)?,
                    0x10 => {
                        let start = *stack.last()?;
                        let args = formula.split_off(start);
                        formula.push_str("SUM(");
                        formula.push_str(&args);
                        formula.push(')');
                    }
                    _ => return None,
                }
            }
            0x1c => {
                let err = error_code(*rgce.first()?);
                stack.push(formula.len());
                formula.push_str(err);
                rgce = rgce.get(1..)?;
            }
            0x1d => {
                stack.push(formula.len());
                formula.push_str(if *rgce.first()? == 0 { "FALSE" } else { "TRUE" });
                rgce = rgce.get(1..)?;
            }
            0x1e => {
                let value = rgce.get(0..2).map(le_u16)?;
                stack.push(formula.len());
                formula.push_str(&value.to_string());
                rgce = rgce.get(2..)?;
            }
            0x1f => {
                let value = rgce.get(0..8).map(le_f64)?;
                stack.push(formula.len());
                formula.push_str(&format_formula_number(value));
                rgce = rgce.get(8..)?;
            }
            0x21 | 0x41 | 0x61 => {
                let iftab = rgce.get(0..2).map(le_u16)? as usize;
                let argc = fixed_formula_arg_count(iftab)?;
                rgce = rgce.get(2..)?;
                apply_formula_function(iftab, argc, &mut formula, &mut stack)?;
            }
            0x22 | 0x42 | 0x62 => {
                let argc = *rgce.first()? as usize;
                let iftab = rgce.get(1..3).map(le_u16)? as usize;
                rgce = rgce.get(3..)?;
                apply_formula_function(iftab, argc, &mut formula, &mut stack)?;
            }
            0x24 | 0x44 | 0x64 => {
                let reference = parse_formula_ref(rgce.get(0..6)?)?;
                stack.push(formula.len());
                formula.push_str(&reference);
                rgce = rgce.get(6..)?;
            }
            0x25 | 0x45 | 0x65 => {
                let area = parse_formula_area(rgce.get(0..12)?)?;
                stack.push(formula.len());
                formula.push_str(&area);
                rgce = rgce.get(12..)?;
            }
            0x23 | 0x43 | 0x63 => {
                let name = parse_formula_name(rgce.get(0..4)?, context?)?;
                stack.push(formula.len());
                formula.push_str(&name);
                rgce = rgce.get(4..)?;
            }
            0x3a | 0x5a | 0x7a => {
                let reference = parse_formula_ref3d(rgce.get(0..8)?, context?)?;
                stack.push(formula.len());
                formula.push_str(&reference);
                rgce = rgce.get(8..)?;
            }
            0x3b | 0x5b | 0x7b => {
                let area = parse_formula_area3d(rgce.get(0..14)?, context?)?;
                stack.push(formula.len());
                formula.push_str(&area);
                rgce = rgce.get(14..)?;
            }
            0x2a | 0x4a | 0x6a => {
                stack.push(formula.len());
                formula.push_str("#REF!");
                rgce = rgce.get(6..)?;
            }
            0x2b | 0x4b | 0x6b => {
                stack.push(formula.len());
                formula.push_str("#REF!");
                rgce = rgce.get(12..)?;
            }
            _ => return None,
        }
    }
    (stack.len() == 1).then_some(formula)
}

fn apply_binary_formula_op(ptg: u8, formula: &mut String, stack: &mut Vec<usize>) -> Option<()> {
    let right_start = stack.pop()?;
    let right = formula.split_off(right_start);
    let op = match ptg {
        0x03 => "+",
        0x04 => "-",
        0x05 => "*",
        0x06 => "/",
        0x07 => "^",
        0x08 => "&",
        0x09 => "<",
        0x0a => "<=",
        0x0b => "=",
        0x0c => ">",
        0x0d => ">=",
        0x0e => "<>",
        0x0f => " ",
        0x10 => ",",
        0x11 => ":",
        _ => return None,
    };
    formula.push_str(op);
    formula.push_str(&right);
    Some(())
}

fn apply_formula_function(
    iftab: usize,
    argc: usize,
    formula: &mut String,
    stack: &mut Vec<usize>,
) -> Option<()> {
    let name = formula_function_name(iftab)?;
    if stack.len() < argc {
        return None;
    }
    if argc == 0 {
        stack.push(formula.len());
        formula.push_str(name);
        formula.push_str("()");
        return Some(());
    }
    let args_start = stack.len() - argc;
    let mut arg_offsets = stack.split_off(args_start);
    let start = *arg_offsets.first()?;
    for offset in &mut arg_offsets {
        *offset -= start;
    }
    let args = formula.split_off(start);
    stack.push(formula.len());
    arg_offsets.push(args.len());
    formula.push_str(name);
    formula.push('(');
    for window in arg_offsets.windows(2) {
        formula.push_str(&args[window[0]..window[1]]);
        formula.push(',');
    }
    formula.pop();
    formula.push(')');
    Some(())
}

fn fixed_formula_arg_count(iftab: usize) -> Option<usize> {
    match iftab {
        1 => Some(3), // IF
        2 | 3 | 8 | 9 | 15..=18 | 20..=26 | 32 | 33 | 38 => Some(1),
        10 | 19 | 34 | 35 => Some(0),
        13 | 27 | 30 | 39 | 48 => Some(2),
        14 | 31 | 40..=45 => Some(3),
        29 | 49..=52 => Some(4),
        _ => None,
    }
}

fn formula_function_name(iftab: usize) -> Option<&'static str> {
    match iftab {
        0 => Some("COUNT"),
        1 => Some("IF"),
        2 => Some("ISNA"),
        3 => Some("ISERROR"),
        4 => Some("SUM"),
        5 => Some("AVERAGE"),
        6 => Some("MIN"),
        7 => Some("MAX"),
        8 => Some("ROW"),
        9 => Some("COLUMN"),
        10 => Some("NA"),
        15 => Some("SIN"),
        16 => Some("COS"),
        17 => Some("TAN"),
        19 => Some("PI"),
        20 => Some("SQRT"),
        21 => Some("EXP"),
        22 => Some("LN"),
        23 => Some("LOG10"),
        24 => Some("ABS"),
        25 => Some("INT"),
        26 => Some("SIGN"),
        27 => Some("ROUND"),
        30 => Some("REPT"),
        31 => Some("MID"),
        32 => Some("LEN"),
        33 => Some("VALUE"),
        34 => Some("TRUE"),
        35 => Some("FALSE"),
        36 => Some("AND"),
        37 => Some("OR"),
        38 => Some("NOT"),
        39 => Some("MOD"),
        48 => Some("TEXT"),
        61 => Some("MIRR"),
        63 => Some("RAND"),
        65 => Some("DATE"),
        66 => Some("TIME"),
        67 => Some("DAY"),
        68 => Some("MONTH"),
        69 => Some("YEAR"),
        70 => Some("WEEKDAY"),
        97 => Some("ATAN2"),
        98 => Some("ASIN"),
        99 => Some("ACOS"),
        100 => Some("CHOOSE"),
        101 => Some("HLOOKUP"),
        102 => Some("VLOOKUP"),
        109 => Some("LOG"),
        111 => Some("CHAR"),
        112 => Some("LOWER"),
        113 => Some("UPPER"),
        115 => Some("LEFT"),
        116 => Some("RIGHT"),
        117 => Some("EXACT"),
        118 => Some("TRIM"),
        119 => Some("REPLACE"),
        120 => Some("SUBSTITUTE"),
        124 => Some("FIND"),
        125 => Some("CELL"),
        148 => Some("INDIRECT"),
        162 => Some("CLEAN"),
        163 => Some("MDETERM"),
        164 => Some("MINVERSE"),
        165 => Some("MMULT"),
        167 => Some("IPMT"),
        168 => Some("PPMT"),
        169 => Some("COUNTA"),
        183 => Some("PRODUCT"),
        184 => Some("FACT"),
        193 => Some("DPRODUCT"),
        194 => Some("ISNONTEXT"),
        195 => Some("STDEVP"),
        196 => Some("VARP"),
        197 => Some("DSTDEVP"),
        198 => Some("DVARP"),
        212 => Some("ROUNDUP"),
        213 => Some("ROUNDDOWN"),
        216 => Some("RANK"),
        219 => Some("ADDRESS"),
        220 => Some("DAYS360"),
        221 => Some("TODAY"),
        227 => Some("MEDIAN"),
        228 => Some("SUMPRODUCT"),
        229 => Some("SINH"),
        230 => Some("COSH"),
        231 => Some("TANH"),
        244 => Some("INFO"),
        247 => Some("DB"),
        255 => Some("GETPIVOTDATA"),
        269 => Some("AVEDEV"),
        270 => Some("BETADIST"),
        271 => Some("GAMMALN"),
        276 => Some("COMBIN"),
        279 => Some("CEILING"),
        280 => Some("FLOOR"),
        285 => Some("EVEN"),
        286 => Some("ODD"),
        300 => Some("CEILING"),
        303 => Some("SUMIFS"),
        304 => Some("COUNTIFS"),
        345 => Some("SUMIF"),
        346 => Some("COUNTIF"),
        347 => Some("AVERAGEIF"),
        350 => Some("IFERROR"),
        _ => None,
    }
}

fn parse_formula_ref(payload: &[u8]) -> Option<String> {
    let row = le_u32(&payload[0..4]) + 1;
    let col_flags = le_u16(&payload[4..6]);
    let col = (col_flags & 0x3fff) as u32 + 1;
    Some(format_cell_reference(
        row,
        col,
        col_flags & 0x4000 == 0,
        col_flags & 0x8000 == 0,
    ))
}

fn parse_formula_area(payload: &[u8]) -> Option<String> {
    let first_row = le_u32(&payload[0..4]) + 1;
    let last_row = le_u32(&payload[4..8]) + 1;
    let first_col_flags = le_u16(&payload[8..10]);
    let last_col_flags = le_u16(&payload[10..12]);
    let first_col = (first_col_flags & 0x3fff) as u32 + 1;
    let last_col = (last_col_flags & 0x3fff) as u32 + 1;
    Some(format!(
        "{}:{}",
        format_cell_reference(
            first_row,
            first_col,
            first_col_flags & 0x4000 == 0,
            first_col_flags & 0x8000 == 0,
        ),
        format_cell_reference(
            last_row,
            last_col,
            last_col_flags & 0x4000 == 0,
            last_col_flags & 0x8000 == 0,
        )
    ))
}

fn parse_formula_name(payload: &[u8], context: &FormulaContext<'_>) -> Option<String> {
    let name_index = le_u32(payload);
    if name_index == 0 {
        return None;
    }
    context
        .named_ranges
        .get(name_index as usize - 1)
        .map(|range| range.name.clone())
}

fn parse_formula_ref3d(payload: &[u8], context: &FormulaContext<'_>) -> Option<String> {
    let sheet = formula_sheet_prefix(le_u16(&payload[0..2]) as usize, context)?;
    let reference = parse_formula_ref(payload.get(2..8)?)?;
    Some(format!("{sheet}!{reference}"))
}

fn parse_formula_area3d(payload: &[u8], context: &FormulaContext<'_>) -> Option<String> {
    let sheet = formula_sheet_prefix(le_u16(&payload[0..2]) as usize, context)?;
    let area = parse_formula_area(payload.get(2..14)?)?;
    Some(format!("{sheet}!{area}"))
}

fn formula_sheet_prefix(ixti: usize, context: &FormulaContext<'_>) -> Option<String> {
    let xti = context.extern_sheets.get(ixti)?;
    if xti.first_sheet < 0 || xti.last_sheet < 0 {
        return None;
    }
    let first = context.sheets.get(xti.first_sheet as usize)?;
    let last = context.sheets.get(xti.last_sheet as usize)?;
    if xti.first_sheet == xti.last_sheet {
        Some(quote_sheet_name(&first.name))
    } else {
        Some(format!(
            "{}:{}",
            quote_sheet_name(&first.name),
            quote_sheet_name(&last.name)
        ))
    }
}

fn quote_sheet_name(name: &str) -> String {
    if name
        .chars()
        .all(|ch| ch.is_ascii_alphanumeric() || ch == '_')
    {
        return name.to_string();
    }
    let escaped = name.replace('\'', "''");
    format!("'{escaped}'")
}

fn format_cell_reference(row: u32, col: u32, row_abs: bool, col_abs: bool) -> String {
    let mut out = String::new();
    if col_abs {
        out.push('$');
    }
    push_column_label(col, &mut out);
    if row_abs {
        out.push('$');
    }
    out.push_str(&row.to_string());
    out
}

fn push_column_label(mut col: u32, out: &mut String) {
    let mut buf = Vec::new();
    while col > 0 {
        col -= 1;
        buf.push((b'A' + (col % 26) as u8) as char);
        col /= 26;
    }
    for ch in buf.iter().rev() {
        out.push(*ch);
    }
}

fn format_formula_number(value: f64) -> String {
    if value.fract() == 0.0 && value.abs() < (i64::MAX as f64) {
        (value as i64).to_string()
    } else {
        value.to_string()
    }
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
    freeze_panes: Option<FreezePane>,
    sheet_view: Option<SheetViewInfo>,
    hyperlinks: Vec<Hyperlink>,
    comments: Vec<Comment>,
    cells: Vec<Cell>,
) -> WorksheetData {
    WorksheetData {
        dimension,
        merged_ranges,
        hyperlinks,
        freeze_panes,
        sheet_properties: None::<SheetPropertiesInfo>,
        sheet_view,
        comments,
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

fn le_i32(bytes: &[u8]) -> i32 {
    i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]])
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

    fn put_i32(out: &mut Vec<u8>, value: i32) {
        out.extend_from_slice(&value.to_le_bytes());
    }

    fn put_wide_string(out: &mut Vec<u8>, value: &str) {
        let units: Vec<u16> = value.encode_utf16().collect();
        put_u32(out, units.len() as u32);
        for unit in units {
            put_u16(out, unit);
        }
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

        let sheet =
            parse_worksheet(&data, &[], None, Vec::new(), None).expect("parse synthetic worksheet");
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

    #[test]
    fn parses_xlsb_freeze_pane_record() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&1.0f64.to_le_bytes());
        payload.extend_from_slice(&1.0f64.to_le_bytes());
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 0);
        payload.push(0x02);

        assert_eq!(
            parse_pane(&payload),
            Some(FreezePane {
                mode: PaneMode::Freeze,
                top_left_cell: Some("B2".to_string()),
                x_split: Some(1),
                y_split: Some(1),
                active_pane: Some("bottomRight".to_string()),
            })
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_freeze_panes() {
        let mut data = Vec::new();
        let mut pane = Vec::new();
        pane.extend_from_slice(&1.0f64.to_le_bytes());
        pane.extend_from_slice(&0.0f64.to_le_bytes());
        put_u32(&mut pane, 1);
        put_u32(&mut pane, 0);
        put_u32(&mut pane, 2);
        pane.push(0x02);
        push_record(&mut data, 0x0097, &pane);

        let sheet =
            parse_worksheet(&data, &[], None, Vec::new(), None).expect("parse freeze worksheet");

        assert_eq!(
            sheet.freeze_panes,
            Some(FreezePane {
                mode: PaneMode::Freeze,
                top_left_cell: Some("A2".to_string()),
                x_split: Some(1),
                y_split: Some(0),
                active_pane: Some("bottomLeft".to_string()),
            })
        );
    }

    #[test]
    fn parses_xlsb_sheet_view_record() {
        let mut payload = Vec::new();
        put_u16(&mut payload, 0x0004 | 0x0008 | 0x0010 | 0x0040 | 0x0100);
        put_u32(&mut payload, 2);
        put_u32(&mut payload, 2);
        put_u32(&mut payload, 3);
        payload.push(0);
        payload.push(0);
        put_u16(&mut payload, 0);
        put_u16(&mut payload, 150);
        put_u16(&mut payload, 120);
        put_u16(&mut payload, 80);
        put_u16(&mut payload, 90);
        put_u32(&mut payload, 1);

        let view = parse_sheet_view(&payload).unwrap();

        assert_eq!(view.zoom_scale, 150);
        assert_eq!(view.zoom_scale_normal, 120);
        assert_eq!(view.view, "pageLayout");
        assert!(view.show_grid_lines);
        assert!(view.show_row_col_headers);
        assert!(view.show_outline_symbols);
        assert!(view.show_zeros);
        assert!(view.tab_selected);
        assert_eq!(view.top_left_cell.as_deref(), Some("D3"));
        assert_eq!(view.workbook_view_id, 1);
    }

    #[test]
    fn parses_xlsb_selection_record() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 2);
        put_u32(&mut payload, 3);
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 2);
        put_u32(&mut payload, 4);
        put_u32(&mut payload, 3);
        put_u32(&mut payload, 5);

        let selection = parse_selection(&payload).unwrap();

        assert_eq!(selection.pane.as_deref(), Some("bottomRight"));
        assert_eq!(selection.active_cell.as_deref(), Some("D3"));
        assert_eq!(selection.sqref.as_deref(), Some("D3:F5"));
        assert_eq!(selection.active_cell_id, Some(0));
    }

    #[test]
    fn parse_worksheet_collects_xlsb_sheet_view_metadata() {
        let mut data = Vec::new();
        let mut view = Vec::new();
        put_u16(&mut view, 0x0004 | 0x0008 | 0x0010 | 0x0100);
        put_u32(&mut view, 0);
        put_u32(&mut view, 0);
        put_u32(&mut view, 0);
        view.push(0);
        view.push(0);
        put_u16(&mut view, 0);
        put_u16(&mut view, 100);
        put_u16(&mut view, 100);
        put_u16(&mut view, 100);
        put_u16(&mut view, 100);
        put_u32(&mut view, 0);
        push_record(&mut data, 0x0089, &view);

        let mut selection = Vec::new();
        put_u32(&mut selection, 3);
        put_u32(&mut selection, 0);
        put_u32(&mut selection, 0);
        put_u32(&mut selection, 0);
        put_u32(&mut selection, 1);
        put_u32(&mut selection, 0);
        put_u32(&mut selection, 0);
        put_u32(&mut selection, 0);
        put_u32(&mut selection, 0);
        push_record(&mut data, 0x0098, &selection);
        push_record(&mut data, 0x008a, &[]);

        let sheet = parse_worksheet(&data, &[], None, Vec::new(), None)
            .expect("parse sheet view worksheet");
        let view = sheet.sheet_view.unwrap();

        assert_eq!(view.view, "normal");
        assert_eq!(view.selections.len(), 1);
        assert_eq!(view.selections[0].active_cell.as_deref(), Some("A1"));
    }

    #[test]
    fn parses_xlsb_hyperlink_record_with_relationship_target() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 0);
        put_wide_string(&mut payload, "rId5");
        put_wide_string(&mut payload, "");
        put_wide_string(&mut payload, "tip");
        put_wide_string(&mut payload, "Example");
        let rels = RelsGraph::parse(
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
            </Relationships>"#,
        )
        .expect("parse rels");

        let node = parse_hyperlink(&payload).unwrap();
        let links = resolve_hyperlinks(vec![node], Some(&rels), &[]);

        assert_eq!(
            links,
            vec![Hyperlink {
                cell: "A1".to_string(),
                target: "https://example.com".to_string(),
                display: "Example".to_string(),
                tooltip: Some("tip".to_string()),
                internal: false,
            }]
        );
    }

    #[test]
    fn parses_xlsb_internal_hyperlink_record() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 1);
        put_wide_string(&mut payload, "");
        put_wide_string(&mut payload, "Sheet2!A1");
        put_wide_string(&mut payload, "");
        put_wide_string(&mut payload, "");

        let node = parse_hyperlink(&payload).unwrap();
        let links = resolve_hyperlinks(vec![node], None, &[]);

        assert_eq!(
            links,
            vec![Hyperlink {
                cell: "B2".to_string(),
                target: "Sheet2!A1".to_string(),
                display: String::new(),
                tooltip: None,
                internal: true,
            }]
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_hyperlinks() {
        let mut data = Vec::new();
        let mut link = Vec::new();
        put_u32(&mut link, 0);
        put_u32(&mut link, 0);
        put_u32(&mut link, 0);
        put_u32(&mut link, 0);
        put_wide_string(&mut link, "rId2");
        put_wide_string(&mut link, "");
        put_wide_string(&mut link, "");
        put_wide_string(&mut link, "Docs");
        push_record(&mut data, 0x01ee, &link);
        let rels = RelsGraph::parse(
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://docs.example" TargetMode="External"/>
            </Relationships>"#,
        )
        .expect("parse rels");

        let sheet =
            parse_worksheet(&data, &[], Some(&rels), Vec::new(), None).expect("parse hyperlinks");

        assert_eq!(sheet.hyperlinks.len(), 1);
        assert_eq!(sheet.hyperlinks[0].cell, "A1");
        assert_eq!(sheet.hyperlinks[0].target, "https://docs.example");
        assert_eq!(sheet.hyperlinks[0].display, "Docs");
    }

    #[test]
    fn parses_xlsb_binary_comments_authors_and_text() {
        let mut data = Vec::new();
        let mut author = Vec::new();
        put_wide_string(&mut author, "Analyst");
        push_record(&mut data, 0x0278, &author);

        let mut comment = Vec::new();
        put_u32(&mut comment, 0);
        put_u32(&mut comment, 2);
        put_u32(&mut comment, 2);
        put_u32(&mut comment, 1);
        put_u32(&mut comment, 1);
        comment.extend_from_slice(&[0; 16]);
        push_record(&mut data, 0x027b, &comment);

        let mut text = vec![0];
        put_wide_string(&mut text, "Review EBITDA addback");
        push_record(&mut data, 0x027d, &text);
        push_record(&mut data, 0x027c, &[]);

        assert_eq!(
            parse_comments_bin(&data).expect("parse comments"),
            vec![Comment {
                cell: "B3".to_string(),
                text: "Review EBITDA addback".to_string(),
                author: "Analyst".to_string(),
                threaded: false,
            }]
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_comments() {
        let comments = vec![Comment {
            cell: "A1".to_string(),
            text: "Looks right".to_string(),
            author: "Reviewer".to_string(),
            threaded: false,
        }];

        let sheet =
            parse_worksheet(&[], &[], None, comments.clone(), None).expect("parse comments");

        assert_eq!(sheet.comments, comments);
    }

    #[test]
    fn parses_basic_formula_rgce_tokens() {
        let mut rgce = Vec::new();
        rgce.push(0x44); // PtgRef relative B5
        put_u32(&mut rgce, 4);
        put_u16(&mut rgce, 0xc001);
        rgce.push(0x1e); // PtgInt 2
        put_u16(&mut rgce, 2);
        rgce.push(0x03); // PtgAdd

        assert_eq!(parse_formula_rgce(&rgce), Some("B5+2".to_string()));
    }

    #[test]
    fn parses_variable_function_formula_rgce_tokens() {
        let mut rgce = Vec::new();
        rgce.push(0x45); // PtgArea relative A1:B2
        put_u32(&mut rgce, 0);
        put_u32(&mut rgce, 1);
        put_u16(&mut rgce, 0xc000);
        put_u16(&mut rgce, 0xc001);
        rgce.push(0x22); // PtgFuncVar SUM(argc=1)
        rgce.push(1);
        put_u16(&mut rgce, 4);

        assert_eq!(parse_formula_rgce(&rgce), Some("SUM(A1:B2)".to_string()));
    }

    #[test]
    fn parses_defined_name_formula_rgce_tokens() {
        let named_ranges = vec![NamedRange {
            name: "Revenue".to_string(),
            scope: "workbook".to_string(),
            refers_to: "Data!$A$1".to_string(),
        }];
        let context = FormulaContext {
            sheets: &[],
            extern_sheets: &[],
            named_ranges: &named_ranges,
        };
        let mut rgce = Vec::new();
        rgce.push(0x43);
        put_u32(&mut rgce, 1);

        assert_eq!(
            parse_formula_rgce_with_context(&rgce, Some(&context)),
            Some("Revenue".to_string())
        );
    }

    #[test]
    fn parse_worksheet_attaches_formula_text_to_formula_cells() {
        let mut rgce = Vec::new();
        rgce.push(0x1e);
        put_u16(&mut rgce, 40);
        rgce.push(0x1e);
        put_u16(&mut rgce, 2);
        rgce.push(0x03);

        let mut payload = Vec::new();
        put_u32(&mut payload, 0); // A
        put_u32(&mut payload, 0);
        payload.extend_from_slice(&42.0f64.to_le_bytes());
        put_u16(&mut payload, 0);
        put_u32(&mut payload, rgce.len() as u32);
        payload.extend_from_slice(&rgce);

        let mut data = Vec::new();
        let mut row = Vec::new();
        put_u32(&mut row, 0);
        put_u32(&mut row, 0);
        put_u16(&mut row, 0);
        put_u16(&mut row, 0);
        push_record(&mut data, 0x0000, &row);
        push_record(&mut data, 0x0009, &payload);

        let sheet =
            parse_worksheet(&data, &[], None, Vec::new(), None).expect("parse formula worksheet");
        assert_eq!(sheet.cells.len(), 1);
        assert_eq!(sheet.cells[0].value, CellValue::Number(42.0));
        assert_eq!(sheet.cells[0].formula.as_deref(), Some("40+2"));
    }

    #[test]
    fn parse_worksheet_uses_extern_sheets_for_formula_text() {
        let sheets = vec![
            XlsbSheet {
                name: "Inputs".to_string(),
                path: "xl/worksheets/sheet1.bin".to_string(),
                state: SheetState::Visible,
            },
            XlsbSheet {
                name: "Calc".to_string(),
                path: "xl/worksheets/sheet2.bin".to_string(),
                state: SheetState::Visible,
            },
        ];
        let extern_sheets = vec![XtiRef {
            first_sheet: 0,
            last_sheet: 0,
        }];
        let context = FormulaContext {
            sheets: &sheets,
            extern_sheets: &extern_sheets,
            named_ranges: &[],
        };

        let mut rgce = Vec::new();
        rgce.push(0x5a);
        put_u16(&mut rgce, 0);
        put_u32(&mut rgce, 0);
        put_u16(&mut rgce, 0);

        let mut payload = Vec::new();
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 0);
        payload.extend_from_slice(&42.0f64.to_le_bytes());
        put_u16(&mut payload, 0);
        put_u32(&mut payload, rgce.len() as u32);
        payload.extend_from_slice(&rgce);

        let mut data = Vec::new();
        let mut row = Vec::new();
        put_u32(&mut row, 0);
        put_u32(&mut row, 0);
        put_u16(&mut row, 0);
        put_u16(&mut row, 0);
        push_record(&mut data, 0x0000, &row);
        push_record(&mut data, 0x0009, &payload);

        let sheet = parse_worksheet(&data, &[], None, Vec::new(), Some(&context))
            .expect("parse formula worksheet");

        assert_eq!(sheet.cells[0].formula.as_deref(), Some("Inputs!$A$1"));
    }

    #[test]
    fn parses_xlsb_extern_sheet_records() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 2);
        put_u32(&mut payload, 0);
        put_i32(&mut payload, 0);
        put_i32(&mut payload, 0);
        put_u32(&mut payload, 0);
        put_i32(&mut payload, -2);
        put_i32(&mut payload, -2);

        let refs = parse_extern_sheet(&payload);

        assert_eq!(
            refs,
            vec![
                XtiRef {
                    first_sheet: 0,
                    last_sheet: 0
                },
                XtiRef {
                    first_sheet: -2,
                    last_sheet: -2
                }
            ]
        );
    }

    #[test]
    fn parses_xlsb_defined_name_area3d_formula() {
        let sheets = vec![XlsbSheet {
            name: "Data".to_string(),
            path: "xl/worksheets/sheet1.bin".to_string(),
            state: SheetState::Visible,
        }];
        let extern_sheets = vec![XtiRef {
            first_sheet: 0,
            last_sheet: 0,
        }];
        let mut payload = Vec::new();
        put_u32(&mut payload, 0);
        payload.push(0);
        put_u32(&mut payload, u32::MAX);
        put_wide_string(&mut payload, "GlobalRange");

        let mut rgce = Vec::new();
        rgce.push(0x3b);
        put_u16(&mut rgce, 0);
        put_u32(&mut rgce, 0);
        put_u32(&mut rgce, 1);
        put_u16(&mut rgce, 0);
        put_u16(&mut rgce, 0);
        put_u32(&mut payload, rgce.len() as u32);
        payload.extend_from_slice(&rgce);
        put_u32(&mut payload, 0);

        let raw = parse_defined_name(&payload, &sheets, &extern_sheets).unwrap();
        let ranges = resolve_xlsb_named_ranges(&sheets, vec![raw]);

        assert_eq!(
            ranges,
            vec![NamedRange {
                name: "GlobalRange".to_string(),
                scope: "workbook".to_string(),
                refers_to: "Data!$A$1:$A$2".to_string(),
            }]
        );
    }

    #[test]
    fn resolves_xlsb_sheet_scoped_defined_name_without_sheet_prefix() {
        let sheets = vec![XlsbSheet {
            name: "Other".to_string(),
            path: "xl/worksheets/sheet1.bin".to_string(),
            state: SheetState::Visible,
        }];
        let mut payload = Vec::new();
        put_u32(&mut payload, 0);
        payload.push(0);
        put_u32(&mut payload, 0);
        put_wide_string(&mut payload, "LocalConstant");

        let mut rgce = Vec::new();
        rgce.push(0x1e);
        put_u16(&mut rgce, 7);
        put_u32(&mut payload, rgce.len() as u32);
        payload.extend_from_slice(&rgce);
        put_u32(&mut payload, 0);

        let raw = parse_defined_name(&payload, &sheets, &[]).unwrap();
        let ranges = resolve_xlsb_named_ranges(&sheets, vec![raw]);

        assert_eq!(
            ranges[0],
            NamedRange {
                name: "LocalConstant".to_string(),
                scope: "sheet".to_string(),
                refers_to: "Other!7".to_string(),
            }
        );
    }

    #[test]
    fn resolves_xlsb_print_area_and_titles_from_builtin_names() {
        let sheets = vec![XlsbSheet {
            name: "Data".to_string(),
            path: "xl/worksheets/sheet1.bin".to_string(),
            state: SheetState::Visible,
        }];
        let print_area = RawNamedRange {
            name: "_xlnm.Print_Area".to_string(),
            local_id: Some(0),
            refers_to: "Data!$A$1:$C$10".to_string(),
        };
        let print_titles = RawNamedRange {
            name: "_xlnm.Print_Titles".to_string(),
            local_id: Some(0),
            refers_to: "Data!$1:$2,Data!$A:$B".to_string(),
        };

        let areas = resolve_xlsb_print_areas(&sheets, vec![print_area]);
        let titles = resolve_xlsb_print_titles(&sheets, vec![print_titles]);

        assert_eq!(
            areas.get("Data").map(String::as_str),
            Some("Data!$A$1:$C$10")
        );
        assert_eq!(
            titles.get("Data"),
            Some(&PrintTitlesInfo {
                rows: Some("1:2".to_string()),
                cols: Some("A:B".to_string()),
            })
        );
    }
}
