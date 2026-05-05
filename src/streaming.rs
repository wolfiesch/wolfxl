//! SAX-based streaming sheet reader (Sprint Ι Pod-β).
//!
//! Activated by `load_workbook(path, read_only=True)` (or auto-trigger when
//! a sheet has > 50k rows). Walks `xl/worksheets/sheetN.xml` one row at a
//! time using a hand-rolled byte-scanner over the sheet XML, resolving
//! shared-string-table (SST) references upfront. Style metadata is
//! exposed as a `style_id` only — Python-side `StreamingCell` resolves the
//! actual font/fill/etc. via the eager Rust reader code path
//! (which already loads `xl/styles.xml` and exposes O(1) style lookups).
//!
//! Public surface (Python):
//!
//! - `StreamingSheetReader.open(path, sheet, ...)` — constructor.
//! - `reader.read_next_row()` → `(row_index_1based, [(col_1based, value, style_id, type), ...])`.
//! - `reader.read_next_values(min_col, max_col)` → padded value tuple.
//! - `reader.close()` — eagerly drop the in-memory XML buffer.
//!
//! Memory profile: SST loaded once (typically <10MB even on huge
//! workbooks); sheet XML loaded once (peak RSS scales with sheet XML
//! size, not row × col count) — quick-xml-style streaming would force
//! per-row file-handle reopens against ZipArchive's non-seekable
//! inflate stream and is deferred to a future Pod.

use std::collections::HashMap;
use std::fs::File;

use pyo3::exceptions::{PyIOError, PyStopIteration, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyTuple};
use pyo3::IntoPyObjectExt;

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use zip::ZipArchive;

use crate::ooxml_util;
use crate::util::a1_to_row_col;

type PyObjectOwned = Py<PyAny>;

/// Parse `xl/sharedStrings.xml` into a flat `Vec<String>`. Each entry is
/// the plain-text concatenation of any nested `<r><t>...</t></r>` runs
/// (matches Excel/openpyxl's flattening for `Cell.value`).
fn load_sst(zip: &mut ZipArchive<File>) -> PyResult<Vec<String>> {
    let xml = match ooxml_util::zip_read_to_string_opt(zip, "xl/sharedStrings.xml")? {
        Some(s) => s,
        None => return Ok(Vec::new()),
    };

    let mut reader = XmlReader::from_str(&xml);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();
    let mut out: Vec<String> = Vec::new();
    let mut current = String::new();
    let mut in_si = false;
    let mut in_t = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.local_name();
                match name.as_ref() {
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
                }
            }
            Ok(Event::End(e)) => {
                let name = e.local_name();
                match name.as_ref() {
                    b"si" => {
                        out.push(std::mem::take(&mut current));
                        in_si = false;
                    }
                    b"t" => {
                        in_t = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(t)) => {
                if in_si && in_t {
                    let s = t
                        .unescape()
                        .map_err(|e| PyErr::new::<PyIOError, _>(format!("SST text decode: {e}")))?;
                    current.push_str(&s);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(PyErr::new::<PyIOError, _>(format!(
                    "Failed to parse sharedStrings.xml: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

/// Resolve `<sheet name=...>` → ZIP path (`xl/worksheets/sheetN.xml`).
fn resolve_sheet_xml_path(zip: &mut ZipArchive<File>, sheet: &str) -> PyResult<String> {
    let workbook_xml = ooxml_util::zip_read_to_string(zip, "xl/workbook.xml")?;
    let rels_xml = ooxml_util::zip_read_to_string(zip, "xl/_rels/workbook.xml.rels")?;
    let sheet_rids = ooxml_util::parse_workbook_sheet_rids(&workbook_xml)?;
    let rel_targets = ooxml_util::parse_relationship_targets(&rels_xml)?;
    for (name, rid) in sheet_rids {
        if name == sheet {
            if let Some(target) = rel_targets.get(&rid) {
                return Ok(ooxml_util::join_and_normalize("xl/", target));
            }
        }
    }
    Err(PyErr::new::<PyIOError, _>(format!(
        "Sheet not found in workbook.xml: {sheet}"
    )))
}

/// Streaming sheet reader. See module docs.
#[pyclass(unsendable, module = "wolfxl._rust")]
pub struct StreamingSheetReader {
    /// Owned XML buffer the scanner walks.
    xml: Box<str>,
    /// Byte cursor into `xml`.
    cursor: usize,
    /// Shared-strings table loaded upfront from `xl/sharedStrings.xml`.
    sst: Vec<String>,
    /// Whether the reader has been exhausted.
    exhausted: bool,
    /// Optional `min_row` bound (1-based, inclusive). Rows below are skipped.
    min_row: Option<u32>,
    /// Optional `max_row` bound (1-based, inclusive). Iteration stops past this.
    max_row: Option<u32>,
    /// Optional `min_col` bound (1-based, inclusive).
    min_col: Option<u32>,
    /// Optional `max_col` bound (1-based, inclusive).
    max_col: Option<u32>,
}

/// One parsed cell, before Python-tuple boxing.
struct ParsedCell {
    col: u32,
    value: PyObjectOwned,
    style_id: Option<u32>,
    cell_type: &'static str,
}

#[pymethods]
impl StreamingSheetReader {
    /// Open `path` and prepare to stream `sheet`.
    #[staticmethod]
    #[pyo3(signature = (path, sheet, min_row=None, max_row=None, min_col=None, max_col=None))]
    pub fn open(
        path: &str,
        sheet: &str,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
    ) -> PyResult<Self> {
        let file = File::open(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open xlsx: {e}")))?;
        let mut zip = ZipArchive::new(file)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open zip: {e}")))?;

        let sst = load_sst(&mut zip)?;
        let sheet_path = resolve_sheet_xml_path(&mut zip, sheet)?;
        let xml = ooxml_util::zip_read_to_string(&mut zip, &sheet_path)?;

        Ok(Self {
            xml: xml.into_boxed_str(),
            cursor: 0,
            sst,
            exhausted: false,
            min_row,
            max_row,
            min_col,
            max_col,
        })
    }

    /// Read the next row matching the configured bounds.
    ///
    /// Returns `None` when the stream is exhausted. Otherwise returns
    /// `(row_index_1based, [(col_1based, value, style_id_or_None, type_str), ...])`.
    pub fn read_next_row<'py>(&mut self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyTuple>>> {
        if self.exhausted {
            return Ok(None);
        }
        loop {
            match self.parse_one_row(py)? {
                StepResult::Row(row_idx, cells) => {
                    let cell_list = PyList::empty(py);
                    for c in cells {
                        let style_obj = match c.style_id {
                            Some(s) => s.into_py_any(py)?,
                            None => py.None(),
                        };
                        let tup = PyTuple::new(
                            py,
                            [
                                c.col.into_py_any(py)?,
                                c.value,
                                style_obj,
                                c.cell_type.into_py_any(py)?,
                            ],
                        )?;
                        cell_list.append(tup)?;
                    }
                    let outer =
                        PyTuple::new(py, [row_idx.into_py_any(py)?, cell_list.into_py_any(py)?])?;
                    return Ok(Some(outer));
                }
                StepResult::Skip => continue,
                StepResult::Done => {
                    self.exhausted = true;
                    return Ok(None);
                }
            }
        }
    }

    /// Read the next row as a plain tuple of values, padded by column
    /// bounds. Returns `None` when exhausted.
    pub fn read_next_values<'py>(
        &mut self,
        py: Python<'py>,
    ) -> PyResult<Option<Bound<'py, PyTuple>>> {
        if self.exhausted {
            return Ok(None);
        }
        loop {
            match self.parse_one_row(py)? {
                StepResult::Row(_row_idx, cells) => {
                    return Ok(Some(self.row_to_values_tuple(py, cells)?));
                }
                StepResult::Skip => continue,
                StepResult::Done => {
                    self.exhausted = true;
                    return Ok(None);
                }
            }
        }
    }

    /// True once the stream has been fully consumed.
    pub fn is_exhausted(&self) -> bool {
        self.exhausted
    }

    /// Eagerly drop the in-memory XML buffer (releases peak RSS).
    pub fn close(&mut self) {
        self.exhausted = true;
        self.xml = String::new().into_boxed_str();
        self.cursor = 0;
    }

    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__<'py>(
        mut slf: PyRefMut<'_, Self>,
        py: Python<'py>,
    ) -> PyResult<Bound<'py, PyTuple>> {
        match slf.read_next_row(py)? {
            Some(t) => Ok(t),
            None => Err(PyStopIteration::new_err(())),
        }
    }
}

enum StepResult {
    Row(u32, Vec<ParsedCell>),
    Skip,
    Done,
}

impl StreamingSheetReader {
    fn parse_one_row(&mut self, py: Python<'_>) -> PyResult<StepResult> {
        let bytes = self.xml.as_bytes();
        let pos = self.cursor;

        let row_start = match find_tag_open(bytes, pos, b"row") {
            Some(i) => i,
            None => return Ok(StepResult::Done),
        };
        let (attrs_str, open_end, self_closing) = read_open_tag(bytes, row_start)?;
        let row_idx = parse_row_index(attrs_str)?;

        let mut cells: Vec<ParsedCell> = Vec::new();

        if self_closing {
            self.cursor = open_end;
        } else {
            let close = find_close_tag(bytes, open_end, b"row").ok_or_else(|| {
                PyErr::new::<PyIOError, _>(format!(
                    "Streaming reader: unterminated <row r=\"{row_idx}\">"
                ))
            })?;
            let mut p = open_end;
            while p < close {
                let next = match find_tag_open(&bytes[..close], p, b"c") {
                    Some(i) => i,
                    None => break,
                };
                let (c_attrs, c_open_end, c_self_closing) = read_open_tag(bytes, next)?;
                let coord = read_attr(c_attrs, "r");
                let style_id = read_attr(c_attrs, "s").and_then(|s| s.parse::<u32>().ok());
                let t_attr = read_attr(c_attrs, "t").unwrap_or_else(|| "n".to_string());

                let (value_obj, type_token, end_pos) = if c_self_closing {
                    (py.None(), "blank", c_open_end)
                } else {
                    let c_close = find_close_tag(bytes, c_open_end, b"c").ok_or_else(|| {
                        PyErr::new::<PyIOError, _>("Streaming reader: unterminated <c>".to_string())
                    })?;
                    let inner = &bytes[c_open_end..c_close];
                    let (val, tok) = parse_cell_inner(py, inner, &t_attr, &self.sst)?;
                    (val, tok, c_close + b"</c>".len())
                };

                let col = match coord.as_deref() {
                    Some(c) => {
                        let (_r, c0) = a1_to_row_col(c).map_err(PyErr::new::<PyValueError, _>)?;
                        c0 + 1
                    }
                    None => cells.last().map(|c| c.col + 1).unwrap_or(1),
                };

                cells.push(ParsedCell {
                    col,
                    value: value_obj,
                    style_id,
                    cell_type: type_token,
                });

                p = end_pos;
            }
            self.cursor = close + b"</row>".len();
        }

        if let Some(min) = self.min_row {
            if row_idx < min {
                return Ok(StepResult::Skip);
            }
        }
        if let Some(max) = self.max_row {
            if row_idx > max {
                self.exhausted = true;
                return Ok(StepResult::Done);
            }
        }

        if let Some(cmin) = self.min_col {
            cells.retain(|c| c.col >= cmin);
        }
        if let Some(cmax) = self.max_col {
            cells.retain(|c| c.col <= cmax);
        }

        Ok(StepResult::Row(row_idx, cells))
    }

    fn row_to_values_tuple<'py>(
        &self,
        py: Python<'py>,
        cells: Vec<ParsedCell>,
    ) -> PyResult<Bound<'py, PyTuple>> {
        // Determine output width.
        let (cmin, cmax) = match (self.min_col, self.max_col) {
            (Some(a), Some(b)) => (a, b),
            (Some(a), None) => {
                let m = cells.iter().map(|c| c.col).max().unwrap_or(a);
                (a, m.max(a))
            }
            (None, Some(b)) => (1, b),
            (None, None) => {
                if cells.is_empty() {
                    return Ok(PyTuple::empty(py));
                }
                let lo = cells.iter().map(|c| c.col).min().unwrap_or(1);
                let hi = cells.iter().map(|c| c.col).max().unwrap_or(lo);
                (lo, hi)
            }
        };
        if cmax < cmin {
            return Ok(PyTuple::empty(py));
        }
        let width = (cmax - cmin + 1) as usize;
        let mut by_col: HashMap<u32, PyObjectOwned> = HashMap::with_capacity(cells.len());
        for c in cells {
            by_col.insert(c.col, c.value);
        }
        let mut out: Vec<PyObjectOwned> = Vec::with_capacity(width);
        for col in cmin..=cmax {
            match by_col.remove(&col) {
                Some(v) => out.push(v),
                None => out.push(py.None()),
            }
        }
        Ok(PyTuple::new(py, out)?)
    }
}

// ----------------------------------------------------------------------
// Low-level XML scanning helpers.
// ----------------------------------------------------------------------

/// Return the index of the next `<TAG` (followed by space / `>` / `/`).
fn find_tag_open(bytes: &[u8], start: usize, tag: &[u8]) -> Option<usize> {
    let needle_len = 1 + tag.len();
    let mut i = start;
    while i + needle_len < bytes.len() {
        if bytes[i] == b'<' && bytes[i + 1..].starts_with(tag) {
            let after = bytes[i + 1 + tag.len()];
            if matches!(after, b' ' | b'>' | b'/' | b'\t' | b'\n' | b'\r') {
                return Some(i);
            }
        }
        i += 1;
    }
    None
}

/// Return the index of the next `</TAG>` token.
fn find_close_tag(bytes: &[u8], start: usize, tag: &[u8]) -> Option<usize> {
    let needle_len = 3 + tag.len();
    let mut i = start;
    while i + needle_len <= bytes.len() {
        if bytes[i] == b'<' && bytes[i + 1] == b'/' && bytes[i + 2..].starts_with(tag) {
            let after = bytes[i + 2 + tag.len()];
            if matches!(after, b'>' | b' ' | b'\t') {
                return Some(i);
            }
        }
        i += 1;
    }
    None
}

/// Parse the open tag starting at `bytes[start]` (which should be `<`).
/// Returns `(attrs_substring, byte_index_after_open_tag, self_closing)`.
fn read_open_tag(bytes: &[u8], start: usize) -> PyResult<(&str, usize, bool)> {
    debug_assert_eq!(bytes.get(start), Some(&b'<'));
    let mut i = start + 1;
    let mut in_quote: Option<u8> = None;
    while i < bytes.len() {
        let b = bytes[i];
        if let Some(q) = in_quote {
            if b == q {
                in_quote = None;
            }
        } else if b == b'"' || b == b'\'' {
            in_quote = Some(b);
        } else if b == b'>' {
            let self_closing = i > 0 && bytes[i - 1] == b'/';
            let attrs_end = if self_closing { i - 1 } else { i };
            let mut name_end = start + 1;
            while name_end < bytes.len()
                && !matches!(bytes[name_end], b' ' | b'/' | b'>' | b'\t' | b'\n' | b'\r')
            {
                name_end += 1;
            }
            let attrs_slice = &bytes[name_end..attrs_end];
            let attrs_str = std::str::from_utf8(attrs_slice).map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("Streaming reader: invalid UTF-8 in attrs: {e}"))
            })?;
            return Ok((attrs_str, i + 1, self_closing));
        }
        i += 1;
    }
    Err(PyErr::new::<PyIOError, _>(
        "Streaming reader: unterminated open tag".to_string(),
    ))
}

/// Parse `r="A1"` from an attribute substring.
fn read_attr(attrs: &str, name: &str) -> Option<String> {
    // Build candidate offsets for `<sep>name=`.
    let mut idx: Option<usize> = None;
    let prefix = format!("{name}=");
    if attrs.starts_with(&prefix) {
        idx = Some(0);
    }
    if idx.is_none() {
        for sep in [' ', '\t', '\n', '\r'] {
            let needle: String = format!("{sep}{name}=");
            if let Some(p) = attrs.find(&needle) {
                idx = Some(p + 1);
                break;
            }
        }
    }
    let i = idx?;
    let after_eq = i + name.len() + 1;
    let bytes = attrs.as_bytes();
    if after_eq >= bytes.len() {
        return None;
    }
    let quote = bytes[after_eq];
    if quote != b'"' && quote != b'\'' {
        return None;
    }
    let value_start = after_eq + 1;
    let mut j = value_start;
    while j < bytes.len() && bytes[j] != quote {
        j += 1;
    }
    Some(
        std::str::from_utf8(&bytes[value_start..j])
            .ok()?
            .to_string(),
    )
}

fn parse_row_index(attrs: &str) -> PyResult<u32> {
    match read_attr(attrs, "r") {
        Some(s) => s.parse().map_err(|_| {
            PyErr::new::<PyIOError, _>(format!("Streaming reader: bad row index {s:?}"))
        }),
        None => Err(PyErr::new::<PyIOError, _>(
            "Streaming reader: <row> missing r= attribute".to_string(),
        )),
    }
}

fn parse_cell_inner(
    py: Python<'_>,
    inner: &[u8],
    t_attr: &str,
    sst: &[String],
) -> PyResult<(PyObjectOwned, &'static str)> {
    let v_text = extract_inner_text(inner, b"v");
    let f_text = extract_inner_text(inner, b"f");
    let is_text = extract_is_text(inner);

    let formula = f_text;
    let raw_value = v_text.or(is_text);

    match t_attr {
        "s" => {
            let v = raw_value.unwrap_or_default();
            let idx: usize = v
                .parse()
                .map_err(|_| PyErr::new::<PyValueError, _>(format!("Bad SST index: {v:?}")))?;
            let resolved = sst.get(idx).cloned().unwrap_or_default();
            if let Some(formula) = formula {
                return Ok((build_formula_dict(py, &formula, &resolved)?, "formula"));
            }
            Ok((resolved.into_py_any(py)?, "s"))
        }
        "str" => {
            let v = raw_value.unwrap_or_default();
            if let Some(formula) = formula {
                return Ok((build_formula_dict(py, &formula, &v)?, "formula"));
            }
            Ok((v.into_py_any(py)?, "str"))
        }
        "inlineStr" => {
            let v = raw_value.unwrap_or_default();
            Ok((v.into_py_any(py)?, "inlineStr"))
        }
        "b" => {
            let v = raw_value.unwrap_or_default();
            let b = matches!(v.trim(), "1" | "true" | "TRUE");
            if let Some(formula) = formula {
                return Ok((build_formula_dict(py, &formula, &v)?, "formula"));
            }
            Ok((b.into_py_any(py)?, "b"))
        }
        "e" => {
            let v = raw_value.unwrap_or_else(|| "#ERROR!".to_string());
            let d = PyDict::new(py);
            d.set_item("type", "error")?;
            d.set_item("value", &v)?;
            Ok((d.into_py_any(py)?, "e"))
        }
        "d" => {
            let v = raw_value.unwrap_or_default();
            Ok((v.into_py_any(py)?, "d"))
        }
        _ => {
            let v = match raw_value {
                Some(s) => s,
                None => return Ok((py.None(), "blank")),
            };
            if let Some(formula) = formula {
                return Ok((build_formula_dict(py, &formula, &v)?, "formula"));
            }
            if let Ok(i) = v.parse::<i64>() {
                // Excel-stored ints: surface as int when round-trip safe.
                return Ok((i.into_py_any(py)?, "n"));
            }
            let f: f64 = v
                .parse()
                .map_err(|_| PyErr::new::<PyValueError, _>(format!("Bad numeric value: {v:?}")))?;
            Ok((f.into_py_any(py)?, "n"))
        }
    }
}

fn build_formula_dict(py: Python<'_>, formula: &str, cached: &str) -> PyResult<PyObjectOwned> {
    let d = PyDict::new(py);
    let f_with_eq = if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    };
    d.set_item("type", "formula")?;
    d.set_item("formula", &f_with_eq)?;
    d.set_item("cached", cached)?;
    Ok(d.into_py_any(py)?)
}

fn extract_inner_text(inner: &[u8], tag: &[u8]) -> Option<String> {
    let open = find_tag_open(inner, 0, tag)?;
    let (_, after_open, self_closing) = read_open_tag(inner, open).ok()?;
    if self_closing {
        return Some(String::new());
    }
    let close = find_close_tag(inner, after_open, tag)?;
    let raw = &inner[after_open..close];
    Some(unescape_xml(raw))
}

fn extract_is_text(inner: &[u8]) -> Option<String> {
    let is_open = find_tag_open(inner, 0, b"is")?;
    let (_, after_is, _) = read_open_tag(inner, is_open).ok()?;
    let is_close = find_close_tag(inner, after_is, b"is")?;
    let is_inner = &inner[after_is..is_close];
    let mut out = String::new();
    let mut p = 0;
    while p < is_inner.len() {
        let t_open = match find_tag_open(is_inner, p, b"t") {
            Some(i) => i,
            None => break,
        };
        let (_, after_t, t_self) = read_open_tag(is_inner, t_open).ok()?;
        if t_self {
            p = after_t;
            continue;
        }
        let t_close = find_close_tag(is_inner, after_t, b"t")?;
        out.push_str(&unescape_xml(&is_inner[after_t..t_close]));
        p = t_close + b"</t>".len();
    }
    Some(out)
}

fn unescape_xml(b: &[u8]) -> String {
    if !b.contains(&b'&') {
        return String::from_utf8_lossy(b).into_owned();
    }
    let s = String::from_utf8_lossy(b);
    s.replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn find_tag_basic() {
        let xml = b"<row r=\"1\"><c r=\"A1\"><v>3</v></c></row>";
        assert_eq!(find_tag_open(xml, 0, b"row"), Some(0));
        assert_eq!(find_tag_open(xml, 0, b"c"), Some(11));
        assert_eq!(find_tag_open(xml, 0, b"v"), Some(21));
    }

    #[test]
    fn open_tag_attrs() {
        let xml = b"<c r=\"A1\" s=\"3\" t=\"s\"><v>0</v></c>";
        let (attrs, _end, sc) = read_open_tag(xml, 0).unwrap();
        assert!(!sc);
        assert!(attrs.contains("r=\"A1\""));
        assert_eq!(read_attr(attrs, "r").as_deref(), Some("A1"));
        assert_eq!(read_attr(attrs, "s").as_deref(), Some("3"));
        assert_eq!(read_attr(attrs, "t").as_deref(), Some("s"));
    }

    #[test]
    fn open_tag_self_closing() {
        let xml = b"<c r=\"A1\" s=\"3\"/>";
        let (_attrs, _end, sc) = read_open_tag(xml, 0).unwrap();
        assert!(sc);
    }

    #[test]
    fn close_tag_position() {
        let xml = b"<c r=\"A1\"><v>3</v></c>more";
        let close = find_close_tag(xml, 0, b"c").unwrap();
        assert_eq!(&xml[close..close + 4], b"</c>");
    }

    #[test]
    fn extract_v_text() {
        let inner = b"<v>3.14</v>";
        assert_eq!(extract_inner_text(inner, b"v").as_deref(), Some("3.14"));
    }

    #[test]
    fn extract_is_inline() {
        let inner = b"<is><t>hello</t></is>";
        assert_eq!(extract_is_text(inner).as_deref(), Some("hello"));
    }

    #[test]
    fn unescape_basic() {
        assert_eq!(unescape_xml(b"a &amp; b"), "a & b");
        assert_eq!(unescape_xml(b"plain"), "plain");
    }
}
