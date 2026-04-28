//! Worksheet XML stream-patcher.
//!
//! Takes a worksheet XML string and a set of cell patches, produces a new XML
//! string with those cells replaced or inserted.  Uses quick-xml's streaming
//! reader+writer to avoid building a full DOM.
//!
//! WolfXL uses **inline strings** (`t="str"`) for all new string values.  This
//! avoids modifying the shared string table for the common case.

use std::collections::{BTreeMap, BTreeSet};
use std::io::Write;

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

use crate::ooxml_util::attr_value;

// ---------------------------------------------------------------------------
// Cell patch types
// ---------------------------------------------------------------------------

/// What to write into a cell.
#[derive(Debug, Clone)]
pub enum CellValue {
    /// Empty / blank cell.
    Blank,
    /// Numeric value (integer or float).
    Number(f64),
    /// String value — written as inline string (`t="str"`).
    String(String),
    /// Boolean value.
    Boolean(bool),
    /// Formula string (e.g. `"SUM(A1:A2)"`).
    Formula(String),
    /// Rich-text runs — emitted as `t="inlineStr"` with `<is>...</is>`
    /// containing one `<r><rPr/><t/></r>` per run.  Sprint Ι Pod-α.
    RichText(Vec<wolfxl_writer::rich_text::RichTextRun>),
    /// RFC-057 (Sprint Ο Pod 1C): array-formula master cell.
    /// Emitted as `<c r="..."><f t="array" ref="...">text</f></c>`.
    ArrayFormula {
        /// Spill / array range, e.g. `"A1:A10"`.
        ref_range: String,
        /// Formula body without the leading `=` and without the
        /// surrounding `{}` braces.
        text: String,
    },
    /// RFC-057: data-table formula master cell.
    DataTableFormula {
        ref_range: String,
        ca: bool,
        dt2_d: bool,
        dtr: bool,
        r1: Option<String>,
        r2: Option<String>,
    },
    /// RFC-057: bare placeholder cell that lives inside an
    /// array-formula's spill range (everything except the master).
    /// Emitted as `<c r="..."/>`.
    SpillChild,
}

/// A single cell modification.
#[derive(Debug, Clone)]
pub struct CellPatch {
    /// 1-based row number.
    pub row: u32,
    /// 1-based column number.
    pub col: u32,
    /// New value (or None to keep existing).
    pub value: Option<CellValue>,
    /// New style index (or None to keep existing).
    pub style_index: Option<u32>,
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Patch a worksheet XML string with the given cell modifications.
///
/// Cells are replaced if they already exist, or inserted at the correct
/// sorted position if they don't.  Rows are created as needed.
///
/// The `shared_strings` table is used only to resolve existing shared string
/// values in cells that aren't being patched (for context — we don't modify
/// them).
pub fn patch_worksheet(xml: &str, patches: &[CellPatch]) -> Result<String, String> {
    if patches.is_empty() {
        return Ok(xml.to_string());
    }

    // Group patches by row for efficient lookup.
    // Within each row, map col -> patch.
    let mut row_patches: BTreeMap<u32, BTreeMap<u32, &CellPatch>> = BTreeMap::new();
    for p in patches {
        row_patches.entry(p.row).or_default().insert(p.col, p);
    }

    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::new());
    let mut buf: Vec<u8> = Vec::new();

    // State tracking
    let mut in_sheet_data = false;
    let mut current_row: Option<u32> = None;
    let mut current_row_cols_seen: BTreeSet<u32> = BTreeSet::new();
    let mut rows_seen: BTreeSet<u32> = BTreeSet::new();
    let mut skip_until_cell_end = false; // skip children of a cell being replaced
    let mut worksheet_prefix: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let tag = e.local_name().as_ref().to_vec();
                capture_prefix(&mut worksheet_prefix, e.name().as_ref(), &tag);

                if tag == b"sheetData" {
                    in_sheet_data = true;
                    write_event(&mut writer, Event::Start(e.to_owned()))?;
                } else if tag == b"row" && in_sheet_data {
                    let row_num = attr_value(e, b"r")
                        .and_then(|s| s.parse::<u32>().ok())
                        .unwrap_or(0);

                    // Insert any missing rows that should come before this one
                    for &pr in row_patches.keys() {
                        if pr < row_num && !rows_seen.contains(&pr) {
                            write_new_row(
                                &mut writer,
                                pr,
                                row_patches.get(&pr).unwrap(),
                                worksheet_prefix.as_deref(),
                            )?;
                            rows_seen.insert(pr);
                        }
                    }

                    current_row = Some(row_num);
                    current_row_cols_seen.clear();
                    rows_seen.insert(row_num);
                    write_event(&mut writer, Event::Start(e.to_owned()))?;
                } else if tag == b"c" && in_sheet_data {
                    let cell_ref = attr_value(e, b"r").unwrap_or_default();
                    let (_, col) = parse_cell_ref(&cell_ref);

                    current_row_cols_seen.insert(col);

                    if let Some(row_map) = current_row.and_then(|r| row_patches.get(&r)) {
                        if let Some(patch) = row_map.get(&col) {
                            // This cell is being patched.
                            // If it's style-only (no value change), preserve the original
                            // children (<v>, <f>, etc.) and only rewrite the <c ...> attrs.
                            if patch.value.is_none() && patch.style_index.is_some() {
                                write_style_only_cell_start(
                                    &mut writer,
                                    &cell_ref,
                                    e,
                                    patch,
                                    worksheet_prefix.as_deref(),
                                )?;
                                // Do NOT skip children.
                            } else {
                                // Value patch: replace the entire cell element.
                                write_patched_cell(
                                    &mut writer,
                                    &cell_ref,
                                    e,
                                    patch,
                                    worksheet_prefix.as_deref(),
                                )?;
                                skip_until_cell_end = true;
                            }
                        } else {
                            // Not patched — pass through
                            write_event(&mut writer, Event::Start(e.to_owned()))?;
                        }
                    } else {
                        write_event(&mut writer, Event::Start(e.to_owned()))?;
                    }
                } else {
                    if !skip_until_cell_end {
                        write_event(&mut writer, Event::Start(e.to_owned()))?;
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let tag = e.local_name().as_ref().to_vec();
                capture_prefix(&mut worksheet_prefix, e.name().as_ref(), &tag);

                if tag == b"row" && in_sheet_data {
                    // Self-closing empty row — handle insertions
                    let row_num = attr_value(e, b"r")
                        .and_then(|s| s.parse::<u32>().ok())
                        .unwrap_or(0);

                    for &pr in row_patches.keys() {
                        if pr < row_num && !rows_seen.contains(&pr) {
                            write_new_row(
                                &mut writer,
                                pr,
                                row_patches.get(&pr).unwrap(),
                                worksheet_prefix.as_deref(),
                            )?;
                            rows_seen.insert(pr);
                        }
                    }
                    rows_seen.insert(row_num);

                    // If this empty row has patches, expand it
                    if let Some(row_map) = row_patches.get(&row_num) {
                        write_new_row(&mut writer, row_num, row_map, worksheet_prefix.as_deref())?;
                    } else {
                        write_event(&mut writer, Event::Empty(e.to_owned()))?;
                    }
                } else if tag == b"c" && in_sheet_data {
                    // Self-closing cell (no value/formula children)
                    let cell_ref = attr_value(e, b"r").unwrap_or_default();
                    let (_, col) = parse_cell_ref(&cell_ref);

                    current_row_cols_seen.insert(col);

                    if let Some(row_map) = current_row.and_then(|r| row_patches.get(&r)) {
                        if let Some(patch) = row_map.get(&col) {
                            write_patched_cell(
                                &mut writer,
                                &cell_ref,
                                e,
                                patch,
                                worksheet_prefix.as_deref(),
                            )?;
                        } else {
                            write_event(&mut writer, Event::Empty(e.to_owned()))?;
                        }
                    } else {
                        write_event(&mut writer, Event::Empty(e.to_owned()))?;
                    }
                } else if tag == b"sheetData" {
                    // Empty <sheetData/> — need to insert all rows
                    let sheet_data_name = qname(worksheet_prefix.as_deref(), "sheetData");
                    let start = BytesStart::new(sheet_data_name.as_str());
                    write_event(&mut writer, Event::Start(start))?;
                    for (&row_num, row_map) in &row_patches {
                        write_new_row(&mut writer, row_num, row_map, worksheet_prefix.as_deref())?;
                        rows_seen.insert(row_num);
                    }
                    write_event(
                        &mut writer,
                        Event::End(BytesEnd::new(sheet_data_name.as_str())),
                    )?;
                } else {
                    if !skip_until_cell_end {
                        write_event(&mut writer, Event::Empty(e.to_owned()))?;
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let tag = e.local_name().as_ref().to_vec();

                if tag == b"c" && skip_until_cell_end {
                    skip_until_cell_end = false;
                    // Already wrote the replacement cell — don't write end tag
                } else if tag == b"row" && in_sheet_data {
                    // Before closing row, insert any new cells for this row
                    if let Some(r) = current_row {
                        if let Some(row_map) = row_patches.get(&r) {
                            for (&col, patch) in row_map.iter() {
                                if !current_row_cols_seen.contains(&col) {
                                    let cell_ref = col_row_to_a1(col, r);
                                    write_new_cell(
                                        &mut writer,
                                        &cell_ref,
                                        patch,
                                        worksheet_prefix.as_deref(),
                                    )?;
                                }
                            }
                        }
                    }
                    current_row = None;
                    write_event(&mut writer, Event::End(e.to_owned()))?;
                } else if tag == b"sheetData" {
                    // Before closing sheetData, insert any remaining rows
                    for (&row_num, row_map) in &row_patches {
                        if !rows_seen.contains(&row_num) {
                            write_new_row(
                                &mut writer,
                                row_num,
                                row_map,
                                worksheet_prefix.as_deref(),
                            )?;
                        }
                    }
                    in_sheet_data = false;
                    write_event(&mut writer, Event::End(e.to_owned()))?;
                } else if !skip_until_cell_end {
                    write_event(&mut writer, Event::End(e.to_owned()))?;
                }
            }
            Ok(event @ Event::Text(_))
            | Ok(event @ Event::CData(_))
            | Ok(event @ Event::Comment(_))
            | Ok(event @ Event::Decl(_))
            | Ok(event @ Event::PI(_))
            | Ok(event @ Event::DocType(_)) => {
                if !skip_until_cell_end {
                    write_event(&mut writer, event.into_owned())?;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML parse error: {e}")),
        }
        buf.clear();
    }

    let out = writer.into_inner();
    String::from_utf8(out).map_err(|e| format!("Output not UTF-8: {e}"))
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn write_event<W: Write>(writer: &mut XmlWriter<W>, event: Event<'_>) -> Result<(), String> {
    writer
        .write_event(event)
        .map_err(|e| format!("XML write error: {e}"))
}

/// Write a patched `<c ...>` start tag for a style-only patch, preserving original children.
fn write_style_only_cell_start<W: Write>(
    writer: &mut XmlWriter<W>,
    cell_ref: &str,
    original: &BytesStart<'_>,
    patch: &CellPatch,
    prefix: Option<&str>,
) -> Result<(), String> {
    let cell_name = qname_for_original(original.name().as_ref(), "c", prefix);
    let mut elem = BytesStart::new(cell_name.as_str());

    // Copy all original attributes except r/s. We'll re-add r and (patched) s.
    for a in original.attributes() {
        let a = a.map_err(|e| format!("XML attr error: {e}"))?;
        let key = a.key.as_ref();
        if key == b"r" || key == b"s" {
            continue;
        }
        elem.push_attribute((key, a.value.as_ref()));
    }

    // Always set r (cell reference)
    elem.push_attribute((b"r".as_slice(), cell_ref.as_bytes()));

    // Apply (or preserve) style index
    let style = if let Some(s) = patch.style_index {
        Some(s)
    } else {
        attr_value(original, b"s").and_then(|s| s.parse().ok())
    };
    if let Some(s) = style {
        if s > 0 {
            let sval = s.to_string();
            elem.push_attribute((b"s".as_slice(), sval.as_bytes()));
        }
    }

    write_event(writer, Event::Start(elem))
}

/// Write a complete patched cell element.
fn write_patched_cell<W: Write>(
    writer: &mut XmlWriter<W>,
    cell_ref: &str,
    original: &BytesStart<'_>,
    patch: &CellPatch,
    prefix: Option<&str>,
) -> Result<(), String> {
    let original_name = original.name();
    let effective_prefix = prefix_for_original(original_name.as_ref(), b"c").or(prefix);
    let cell_name = qname(effective_prefix, "c");
    let value_name = qname(effective_prefix, "v");
    let formula_name = qname(effective_prefix, "f");
    let inline_string_name = qname(effective_prefix, "is");

    let mut elem = BytesStart::new(cell_name.as_str());
    elem.push_attribute(("r", cell_ref));

    // Style index: use patch value if set, otherwise preserve original
    let style = if let Some(s) = patch.style_index {
        Some(s)
    } else {
        attr_value(original, b"s").and_then(|s| s.parse().ok())
    };
    if let Some(s) = style {
        if s > 0 {
            elem.push_attribute(("s", s.to_string().as_str()));
        }
    }

    match &patch.value {
        Some(CellValue::Blank) | None => {
            if patch.value.is_some() {
                // Explicit blank — write empty cell
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(|e| format!("XML write error: {e}"))?;
            } else {
                // No value change — need to preserve original value
                // For simplicity, write the cell with original type
                // This path means only style changed, copy original attributes
                let orig_type = attr_value(original, b"t");
                if let Some(t) = &orig_type {
                    elem.push_attribute(("t", t.as_str()));
                }
                // Write as start tag, original children will follow via skip logic...
                // Actually, since skip_until_cell_end skips children, we need to
                // read and replay them.  For now, write empty if no value patch.
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(|e| format!("XML write error: {e}"))?;
            }
        }
        Some(CellValue::Number(n)) => {
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            // <v>number</v>
            let v_start = BytesStart::new(value_name.as_str());
            writer
                .write_event(Event::Start(v_start))
                .map_err(|e| format!("XML write error: {e}"))?;
            let text = if *n == (*n as i64) as f64 {
                format!("{}", *n as i64)
            } else {
                format!("{n}")
            };
            writer
                .write_event(Event::Text(BytesText::new(&text)))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::String(s)) => {
            elem.push_attribute(("t", "str"));
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            let v_start = BytesStart::new(value_name.as_str());
            writer
                .write_event(Event::Start(v_start))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::Text(BytesText::new(s)))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::Boolean(b)) => {
            elem.push_attribute(("t", "b"));
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            let v_start = BytesStart::new(value_name.as_str());
            writer
                .write_event(Event::Start(v_start))
                .map_err(|e| format!("XML write error: {e}"))?;
            let val = if *b { "1" } else { "0" };
            writer
                .write_event(Event::Text(BytesText::new(val)))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::Formula(f)) => {
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            // <f>formula</f> — no <v> (force recalc)
            let f_start = BytesStart::new(formula_name.as_str());
            writer
                .write_event(Event::Start(f_start))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::Text(BytesText::new(f)))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(formula_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::ArrayFormula { ref_range, text }) => {
            // RFC-057: <c r="..."><f t="array" ref="A1:A10">B1:B10*2</f></c>
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            let mut f_start = BytesStart::new(formula_name.as_str());
            f_start.push_attribute(("t", "array"));
            f_start.push_attribute(("ref", ref_range.as_str()));
            writer
                .write_event(Event::Start(f_start))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::Text(BytesText::new(text)))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(formula_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::DataTableFormula {
            ref_range,
            ca,
            dt2_d,
            dtr,
            r1,
            r2,
        }) => {
            // RFC-057: <c r="..."><f t="dataTable" ref=".." dt2D="1" r1=".." r2=".."/></c>
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            let mut f_empty = BytesStart::new(formula_name.as_str());
            f_empty.push_attribute(("t", "dataTable"));
            f_empty.push_attribute(("ref", ref_range.as_str()));
            if *ca {
                f_empty.push_attribute(("ca", "1"));
            }
            if *dt2_d {
                f_empty.push_attribute(("dt2D", "1"));
            }
            if *dtr {
                f_empty.push_attribute(("dtr", "1"));
            }
            if let Some(rv) = r1.as_ref() {
                f_empty.push_attribute(("r1", rv.as_str()));
            }
            if let Some(rv) = r2.as_ref() {
                f_empty.push_attribute(("r2", rv.as_str()));
            }
            writer
                .write_event(Event::Empty(f_empty))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::SpillChild) => {
            // RFC-057: bare placeholder `<c r="..."/>`.  Style is
            // already on `elem` (set above when constructing the start tag).
            writer
                .write_event(Event::Empty(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::RichText(runs)) => {
            // Sprint Ι Pod-α: rich-text writes use inline strings so
            // the SST never has to be touched.  This matches openpyxl's
            // own rich-text emit path (see the test fixture in the
            // Sprint Ι Pod-α task brief).
            elem.push_attribute(("t", "inlineStr"));
            // Emit the start tag, then hand-write the `<is>` body so
            // we can leverage the run-emitter from `wolfxl_writer::rich_text`
            // verbatim (xml-escape, xml:space=preserve, etc).
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            let body = wolfxl_writer::rich_text::emit_runs(runs);
            let raw = format!("<{inline_string_name}>{body}</{inline_string_name}>");
            // BytesText would re-escape; we want the run XML emitted
            // verbatim. quick-xml's `Writer::get_mut()` lets us drop in
            // raw bytes between events without breaking the surrounding
            // structure.
            writer
                .get_mut()
                .write_all(raw.as_bytes())
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
    }

    Ok(())
}

/// Write a brand-new cell element (insertion, not replacement).
fn write_new_cell<W: Write>(
    writer: &mut XmlWriter<W>,
    cell_ref: &str,
    patch: &CellPatch,
    prefix: Option<&str>,
) -> Result<(), String> {
    let dummy = BytesStart::new("c");
    write_patched_cell(writer, cell_ref, &dummy, patch, prefix)
}

/// Write a brand-new `<row>` element containing patched cells.
fn write_new_row<W: Write>(
    writer: &mut XmlWriter<W>,
    row_num: u32,
    cells: &BTreeMap<u32, &CellPatch>,
    prefix: Option<&str>,
) -> Result<(), String> {
    let row_name = qname(prefix, "row");
    let mut row_elem = BytesStart::new(row_name.as_str());
    row_elem.push_attribute(("r", row_num.to_string().as_str()));

    writer
        .write_event(Event::Start(row_elem))
        .map_err(|e| format!("XML write error: {e}"))?;

    for (&col, patch) in cells {
        let cell_ref = col_row_to_a1(col, row_num);
        write_new_cell(writer, &cell_ref, patch, prefix)?;
    }

    writer
        .write_event(Event::End(BytesEnd::new(row_name.as_str())))
        .map_err(|e| format!("XML write error: {e}"))?;

    Ok(())
}

fn capture_prefix(prefix: &mut Option<String>, qname: &[u8], local: &[u8]) {
    if prefix.is_some() {
        return;
    }
    if let Some(found) = prefix_for_original(qname, local) {
        *prefix = Some(found.to_string());
    }
}

fn qname(prefix: Option<&str>, local: &str) -> String {
    match prefix {
        Some(prefix) if !prefix.is_empty() => format!("{prefix}:{local}"),
        _ => local.to_string(),
    }
}

fn qname_for_original(qname_bytes: &[u8], local: &str, fallback_prefix: Option<&str>) -> String {
    let local_bytes = local.as_bytes();
    qname(
        prefix_for_original(qname_bytes, local_bytes).or(fallback_prefix),
        local,
    )
}

fn prefix_for_original<'a>(qname: &'a [u8], local: &[u8]) -> Option<&'a str> {
    if !qname.ends_with(local) {
        return None;
    }
    let prefix_len = qname.len().checked_sub(local.len() + 1)?;
    if qname.get(prefix_len) != Some(&b':') {
        return None;
    }
    std::str::from_utf8(&qname[..prefix_len]).ok()
}

/// Parse a cell reference like "B3" into (row=3, col=2) — both 1-based.
fn parse_cell_ref(cell_ref: &str) -> (u32, u32) {
    let mut col: u32 = 0;
    let mut row_str = String::new();

    for ch in cell_ref.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - b'A' as u32 + 1);
        } else if ch.is_ascii_digit() {
            row_str.push(ch);
        }
    }

    let row = row_str.parse::<u32>().unwrap_or(0);
    (row, col)
}

/// Convert 1-based (col, row) to A1-style reference.
fn col_row_to_a1(col: u32, row: u32) -> String {
    let mut letters = String::new();
    let mut c = col;
    while c > 0 {
        c -= 1;
        letters.insert(0, (b'A' + (c % 26) as u8) as char);
        c /= 26;
    }
    format!("{letters}{row}")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), (1, 1));
        assert_eq!(parse_cell_ref("B3"), (3, 2));
        assert_eq!(parse_cell_ref("AA100"), (100, 27));
        assert_eq!(parse_cell_ref("Z1"), (1, 26));
    }

    #[test]
    fn test_col_row_to_a1() {
        assert_eq!(col_row_to_a1(1, 1), "A1");
        assert_eq!(col_row_to_a1(2, 3), "B3");
        assert_eq!(col_row_to_a1(27, 100), "AA100");
        assert_eq!(col_row_to_a1(26, 1), "Z1");
    }

    #[test]
    fn test_patch_replace_value() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet><sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row>
</sheetData></worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 2, // B1
            value: Some(CellValue::Number(99.0)),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("<v>99</v>"));
        // A1 should be unchanged (though type=s is preserved)
        assert!(result.contains("r=\"A1\""));
    }

    #[test]
    fn test_patch_insert_new_cell() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><v>1</v></c></row>
</sheetData></worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 3, // C1 — doesn't exist
            value: Some(CellValue::String("new".to_string())),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("r=\"C1\""));
        assert!(result.contains("t=\"str\""));
        assert!(result.contains("<v>new</v>"));
    }

    #[test]
    fn test_patch_insert_new_cell_preserves_prefixed_namespace() {
        let xml = r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c></x:row></x:sheetData></x:worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 10, // J1 — doesn't exist
            value: Some(CellValue::String("wolfxl_modify_smoke".to_string())),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains(r#"<x:c r="J1" t="str"><x:v>wolfxl_modify_smoke</x:v></x:c>"#));
        assert!(!result.contains(r#"<c r="J1""#));
    }

    #[test]
    fn test_patch_insert_new_row() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><v>1</v></c></row>
<row r="3"><c r="A3"><v>3</v></c></row>
</sheetData></worksheet>"#;

        let patches = vec![CellPatch {
            row: 2,
            col: 1, // A2 — row doesn't exist
            value: Some(CellValue::String("inserted".to_string())),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("r=\"A2\""));
        assert!(result.contains("<v>inserted</v>"));
        // Verify row ordering: row 1 before row 2 before row 3
        let pos_r1 = result.find("r=\"A1\"").unwrap();
        let pos_r2 = result.find("r=\"A2\"").unwrap();
        let pos_r3 = result.find("r=\"A3\"").unwrap();
        assert!(pos_r1 < pos_r2);
        assert!(pos_r2 < pos_r3);
    }

    #[test]
    fn test_patch_formula() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><v>10</v></c></row>
</sheetData></worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 1,
            value: Some(CellValue::Formula("SUM(B1:B10)".to_string())),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("<f>SUM(B1:B10)</f>"));
        // No <v> — forces recalculation
        assert!(!result.contains("<v>10</v>"));
    }

    #[test]
    fn test_patch_with_style() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><v>42</v></c></row>
</sheetData></worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 1,
            value: Some(CellValue::Number(42.0)),
            style_index: Some(5),
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("s=\"5\""));
    }

    #[test]
    fn test_patch_boolean() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><v>0</v></c></row>
</sheetData></worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 1,
            value: Some(CellValue::Boolean(true)),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("t=\"b\""));
        assert!(result.contains("<v>1</v>"));
    }

    #[test]
    fn test_patch_empty_sheet_data() {
        let xml = r#"<worksheet><sheetData/></worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 1,
            value: Some(CellValue::String("hello".to_string())),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("r=\"A1\""));
        assert!(result.contains("<v>hello</v>"));
    }

    #[test]
    fn test_no_patches_returns_unchanged() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><v>42</v></c></row>
</sheetData></worksheet>"#;

        let result = patch_worksheet(xml, &[]).unwrap();
        assert_eq!(result, xml);
    }
}
