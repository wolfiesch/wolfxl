//! Top-level orchestrator that walks every workbook part touched by
//! a structural shift and emits the rewritten bytes.
//!
//! Designed to be called once per `(sheet, axis, idx, n)` op with all
//! the workbook's relevant XML parts loaded into the inputs struct.
//! The orchestrator returns a `WorkbookMutations` map of `path → new
//! bytes` that the patcher can fold into its `file_patches` map.
//!
//! Multi-op sequencing is the caller's responsibility — apply each
//! op individually, re-read the rewritten bytes for the next op, then
//! call again.

use std::collections::BTreeMap;

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

use crate::axis::{Axis, ShiftPlan};
use crate::shift_anchors::shift_anchor;
use crate::shift_cells::shift_sheet_cells;
use crate::shift_formulas::{shift_formula, shift_formula_on_sheet};

fn push_attr<'a>(e: &mut BytesStart<'a>, key: &[u8], val: &str) {
    e.push_attribute((key, val.as_bytes()));
}

/// One queued structural-shift op, as supplied by the patcher's
/// `queued_axis_shifts` queue.
#[derive(Debug, Clone)]
pub struct AxisShiftOp {
    /// Sheet name (NOT path).
    pub sheet: String,
    /// Row or col axis.
    pub axis: Axis,
    /// 1-based index where shifting begins.
    pub idx: u32,
    /// Signed shift count (positive = insert; negative = delete).
    pub n: i32,
}

impl AxisShiftOp {
    fn plan(&self) -> ShiftPlan {
        ShiftPlan {
            axis: self.axis,
            idx: self.idx,
            n: self.n,
        }
    }
}

/// All XML/VML parts the orchestrator may need to rewrite, plus a
/// `sheet_positions` map for defined-name `localSheetId` resolution.
pub struct SheetXmlInputs<'a> {
    /// Sheet name → sheet XML bytes (e.g. `xl/worksheets/sheet1.xml`).
    pub sheets: BTreeMap<String, &'a [u8]>,
    /// Sheet name → sheet XML path (so output keys match source paths).
    pub sheet_paths: BTreeMap<String, String>,
    /// Optional `xl/workbook.xml` bytes (for defined-name shift).
    pub workbook_xml: Option<&'a [u8]>,
    /// Per-sheet table parts: sheet name → vec of (path, bytes).
    pub tables: BTreeMap<String, Vec<(String, &'a [u8])>>,
    /// Per-sheet comments part: sheet name → (path, bytes).
    pub comments: BTreeMap<String, (String, &'a [u8])>,
    /// Per-sheet vmlDrawing part: sheet name → (path, bytes).
    pub vml: BTreeMap<String, (String, &'a [u8])>,
    /// Sheet name → 0-based position (for definedName localSheetId).
    pub sheet_positions: BTreeMap<String, u32>,
}

impl<'a> SheetXmlInputs<'a> {
    /// Empty inputs. Useful as a starting point for callers that
    /// build the map field-by-field.
    pub fn empty() -> Self {
        Self {
            sheets: BTreeMap::new(),
            sheet_paths: BTreeMap::new(),
            workbook_xml: None,
            tables: BTreeMap::new(),
            comments: BTreeMap::new(),
            vml: BTreeMap::new(),
            sheet_positions: BTreeMap::new(),
        }
    }
}

/// Output of `apply_workbook_shift`: `path → new bytes`.
#[derive(Debug, Default)]
pub struct WorkbookMutations {
    /// Path → new bytes.
    pub file_patches: BTreeMap<String, Vec<u8>>,
}

/// Apply one structural shift op across every part in `inputs`.
///
/// No-op invariant: an empty `ops` slice or an op with `n == 0`
/// returns an empty `file_patches` map. Callers MUST short-circuit
/// before calling if they want byte-identical output for an empty
/// queue.
pub fn apply_workbook_shift(inputs: SheetXmlInputs<'_>, ops: &[AxisShiftOp]) -> WorkbookMutations {
    let mut out = WorkbookMutations::default();

    if ops.is_empty() {
        return out;
    }

    // Cache rewritten sheet bytes across ops in this call so a single
    // multi-op call sees its own rewrites. (In practice the patcher
    // calls one op at a time, but this future-proofs the API.)
    let mut sheet_bytes: BTreeMap<String, Vec<u8>> = inputs
        .sheets
        .iter()
        .map(|(k, v)| (k.clone(), v.to_vec()))
        .collect();
    let mut workbook_bytes: Option<Vec<u8>> = inputs.workbook_xml.map(|b| b.to_vec());
    let mut table_bytes: BTreeMap<String, BTreeMap<String, Vec<u8>>> = inputs
        .tables
        .iter()
        .map(|(sheet, parts)| {
            (
                sheet.clone(),
                parts.iter().map(|(p, b)| (p.clone(), b.to_vec())).collect(),
            )
        })
        .collect();
    let mut comments_bytes: BTreeMap<String, (String, Vec<u8>)> = inputs
        .comments
        .iter()
        .map(|(sheet, (path, b))| (sheet.clone(), (path.clone(), b.to_vec())))
        .collect();
    let mut vml_bytes: BTreeMap<String, (String, Vec<u8>)> = inputs
        .vml
        .iter()
        .map(|(sheet, (path, b))| (sheet.clone(), (path.clone(), b.to_vec())))
        .collect();

    for op in ops {
        let plan = op.plan();
        if plan.is_noop() {
            continue;
        }

        // 1. Sheet XML.
        if let Some(bytes) = sheet_bytes.get(&op.sheet) {
            let new_bytes = shift_sheet_cells(bytes, &plan);
            // 1b. RFC-031: rewrite the `<cols>` block on Col-axis shifts.
            let new_bytes = crate::cols::shift_sheet_cols_block(&new_bytes, plan);
            sheet_bytes.insert(op.sheet.clone(), new_bytes);
        }

        // 2. Tables on this sheet.
        if let Some(parts) = table_bytes.get_mut(&op.sheet) {
            let mut updated: BTreeMap<String, Vec<u8>> = BTreeMap::new();
            for (path, bytes) in parts.iter() {
                let new_bytes = shift_table_xml(bytes, &plan);
                updated.insert(path.clone(), new_bytes);
            }
            *parts = updated;
        }

        // 3. Comments on this sheet.
        if let Some((path, bytes)) = comments_bytes.get(&op.sheet) {
            let new_bytes = shift_comments_xml(bytes, &plan);
            comments_bytes.insert(op.sheet.clone(), (path.clone(), new_bytes));
        }

        // 4. VML on this sheet.
        if let Some((path, bytes)) = vml_bytes.get(&op.sheet) {
            let new_bytes = shift_vml_xml(bytes, &plan);
            vml_bytes.insert(op.sheet.clone(), (path.clone(), new_bytes));
        }

        // 5. Workbook defined names.
        if let Some(ref wb) = workbook_bytes {
            let new_bytes = shift_defined_names(
                wb,
                &plan,
                &op.sheet,
                inputs.sheet_positions.get(&op.sheet).copied(),
            );
            workbook_bytes = Some(new_bytes);
        }
    }

    // Emit file_patches.
    for (sheet, bytes) in &sheet_bytes {
        if let Some(orig) = inputs.sheets.get(sheet) {
            if bytes.as_slice() != *orig {
                if let Some(path) = inputs.sheet_paths.get(sheet) {
                    out.file_patches.insert(path.clone(), bytes.clone());
                }
            }
        }
    }
    if let Some(bytes) = &workbook_bytes {
        if let Some(orig) = inputs.workbook_xml {
            if bytes.as_slice() != orig {
                out.file_patches
                    .insert("xl/workbook.xml".to_string(), bytes.clone());
            }
        }
    }
    for (_sheet, parts) in &table_bytes {
        for (path, bytes) in parts {
            // Compare against original.
            let mut matched_orig = false;
            for (_s, parts_orig) in &inputs.tables {
                for (p_o, b_o) in parts_orig {
                    if p_o == path {
                        if bytes.as_slice() != *b_o {
                            out.file_patches.insert(path.clone(), bytes.clone());
                        }
                        matched_orig = true;
                        break;
                    }
                }
                if matched_orig {
                    break;
                }
            }
        }
    }
    for (sheet, (path, bytes)) in &comments_bytes {
        if let Some((_, orig)) = inputs.comments.get(sheet) {
            if bytes.as_slice() != *orig {
                out.file_patches.insert(path.clone(), bytes.clone());
            }
        }
    }
    for (sheet, (path, bytes)) in &vml_bytes {
        if let Some((_, orig)) = inputs.vml.get(sheet) {
            if bytes.as_slice() != *orig {
                out.file_patches.insert(path.clone(), bytes.clone());
            }
        }
    }

    out
}

/// Rewrite a `xl/tables/tableN.xml` part: `<table ref>`, `<autoFilter ref>`,
/// `<calculatedColumnFormula>` text, plus the `<tableColumns>` block on
/// Col-axis shifts (insert spawns new `<tableColumn>` entries, delete
/// removes them; `count=` and `id=` are renumbered — RFC-031 §5.4).
pub fn shift_table_xml(xml: &[u8], plan: &ShiftPlan) -> Vec<u8> {
    if plan.is_noop() {
        return xml.to_vec();
    }
    // Capture pre-shift table column band so we can rewrite
    // `<tableColumns>` after the standard ref/autoFilter shift.
    let pre_band: Option<(u32, u32)> = if plan.axis == Axis::Col {
        extract_table_col_band(xml)
    } else {
        None
    };
    let shifted = shift_table_xml_inner(xml, plan);
    if let Some((t_lo, t_hi)) = pre_band {
        rewrite_table_columns_block(&shifted, plan, t_lo, t_hi)
    } else {
        shifted
    }
}

fn shift_table_xml_inner(xml: &[u8], plan: &ShiftPlan) -> Vec<u8> {
    let xml_str = match std::str::from_utf8(xml) {
        Ok(s) => s,
        Err(_) => return xml.to_vec(),
    };
    let mut reader = XmlReader::from_str(xml_str);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(std::io::Cursor::new(Vec::new()));
    let mut buf: Vec<u8> = Vec::new();
    let mut in_calc_formula = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name().as_ref().to_vec();
                match local.as_slice() {
                    b"table" | b"autoFilter" => {
                        let new_e = rewrite_attr_value(e, b"ref", |v| shift_anchor(v, plan));
                        let _ = writer.write_event(Event::Start(new_e));
                    }
                    b"calculatedColumnFormula" => {
                        in_calc_formula = true;
                        let _ = writer.write_event(Event::Start(e.to_owned()));
                    }
                    _ => {
                        let _ = writer.write_event(Event::Start(e.to_owned()));
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name().as_ref().to_vec();
                match local.as_slice() {
                    b"table" | b"autoFilter" => {
                        let new_e = rewrite_attr_value(e, b"ref", |v| shift_anchor(v, plan));
                        let _ = writer.write_event(Event::Empty(new_e));
                    }
                    _ => {
                        let _ = writer.write_event(Event::Empty(e.to_owned()));
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name().as_ref().to_vec();
                if local.as_slice() == b"calculatedColumnFormula" {
                    in_calc_formula = false;
                }
                let _ = writer.write_event(Event::End(BytesEnd::new(
                    String::from_utf8_lossy(local.as_slice()).into_owned(),
                )));
            }
            Ok(Event::Text(ref t)) => {
                if in_calc_formula {
                    let s = match t.unescape() {
                        Ok(c) => c.into_owned(),
                        Err(_) => String::from_utf8_lossy(t.as_ref()).into_owned(),
                    };
                    let new_s = shift_formula(&s, plan);
                    let new_t = BytesText::new(&new_s);
                    let _ = writer.write_event(Event::Text(new_t));
                } else {
                    let _ = writer.write_event(Event::Text(t.to_owned()));
                }
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                let _ = writer.write_event(other);
            }
            Err(_) => break,
        }
        buf.clear();
    }

    writer.into_inner().into_inner()
}

/// Parse the `<table ref="A1:E5">` attribute and return the
/// 1-based [col_lo, col_hi] band. Returns None if no `<table>` ref
/// or the ref is malformed.
fn extract_table_col_band(xml: &[u8]) -> Option<(u32, u32)> {
    let s = std::str::from_utf8(xml).ok()?;
    // Find `<table ` (note trailing space — distinguishes from `<tableColumn`
    // and `<tableColumns`).
    let i = s.find("<table ")?;
    let close = s[i..].find('>')?;
    let elt = &s[i..i + close];
    let r_idx = elt.find(" ref=\"")?;
    let v_start = r_idx + " ref=\"".len();
    let v_end = elt[v_start..].find('"')?;
    let r = &elt[v_start..v_start + v_end];
    parse_ref_col_band(r)
}

/// Parse "A1:E5" → (1, 5). Single cell "B3" → (2, 2).
fn parse_ref_col_band(r: &str) -> Option<(u32, u32)> {
    let (lo, hi) = match r.find(':') {
        Some(c) => (&r[..c], &r[c + 1..]),
        None => (r, r),
    };
    let lo_col = parse_col_letters(lo)?;
    let hi_col = parse_col_letters(hi)?;
    Some((lo_col, hi_col))
}

fn parse_col_letters(cell: &str) -> Option<u32> {
    let mut n = 0u32;
    let bytes = cell.as_bytes();
    let mut i = 0;
    if bytes.get(i)? == &b'$' {
        i += 1;
    }
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        n = n
            .checked_mul(26)?
            .checked_add((bytes[i].to_ascii_uppercase() - b'A' + 1) as u32)?;
        i += 1;
    }
    if n == 0 {
        None
    } else {
        Some(n)
    }
}

/// Rewrite the `<tableColumns count="N">...</tableColumns>` block to
/// reflect a Col-axis insert/delete that overlaps the table band.
///
/// Inputs:
/// - `shifted` — XML *after* `shift_table_xml_inner` has rewritten refs.
/// - `plan` — the col-axis ShiftPlan.
/// - `t_lo`, `t_hi` — the table's PRE-shift 1-based col band.
fn rewrite_table_columns_block(shifted: &[u8], plan: &ShiftPlan, t_lo: u32, t_hi: u32) -> Vec<u8> {
    let s = match std::str::from_utf8(shifted) {
        Ok(s) => s.to_owned(),
        Err(_) => return shifted.to_vec(),
    };
    let block_start = match s.find("<tableColumns") {
        Some(i) => i,
        None => return shifted.to_vec(),
    };
    let block_end = match s[block_start..].find("</tableColumns>") {
        Some(i) => block_start + i + "</tableColumns>".len(),
        None => return shifted.to_vec(),
    };
    let block = &s[block_start..block_end];

    // Parse out each `<tableColumn .../>` element in source order.
    let mut entries: Vec<String> = Vec::new();
    let mut i = 0usize;
    while let Some(start_rel) = block[i..].find("<tableColumn ") {
        let abs_start = i + start_rel;
        // Self-closing; find next "/>" or "</tableColumn>" tag end.
        let end = match block[abs_start..].find("/>") {
            Some(e) => abs_start + e + 2,
            None => break,
        };
        entries.push(block[abs_start..end].to_string());
        i = end;
    }

    // Apply the plan to the per-table positional list.
    let new_entries: Vec<String> = if plan.is_insert() {
        let n = plan.n as u32;
        // Insert n empty cols at position (plan.idx - t_lo + 1) within the
        // table, but ONLY if the insert pivot lands STRICTLY INSIDE the
        // table's pre-shift band (t_lo .. t_hi). Insert at exactly t_lo+1
        // counts; insert at t_hi+1 is "after the table" and adds no cols.
        if plan.idx > t_hi || plan.idx <= t_lo {
            // Outside band → no col change; just renumber ids in case
            // (no-op for renumbering, but keeps the contract clean).
            renumber_ids(entries)
        } else {
            let insert_pos = (plan.idx - t_lo) as usize; // 0-based slot
            let mut out: Vec<String> = Vec::with_capacity(entries.len() + n as usize);
            out.extend_from_slice(&entries[..insert_pos]);
            // Compose new placeholder entries. Names must be unique within
            // the table — use a fresh "ColumnNNN" pattern. id is renumbered
            // below, so emit a placeholder.
            // Find max existing id-suffix to avoid name clashes.
            let mut max_n = 0u32;
            for e in &entries {
                if let Some(name) = extract_attr(e, "name") {
                    if let Some(rest) = name.strip_prefix("Column") {
                        if let Ok(k) = rest.parse::<u32>() {
                            if k > max_n {
                                max_n = k;
                            }
                        }
                    }
                }
            }
            for k in 0..n {
                let new_name = format!("Column{}", max_n + 1 + k);
                out.push(format!(r#"<tableColumn id="0" name="{new_name}"/>"#));
            }
            out.extend_from_slice(&entries[insert_pos..]);
            renumber_ids(out)
        }
    } else if plan.is_delete() {
        let n = plan.abs_n();
        // Band [plan.idx, plan.idx + n - 1]. Drop entries whose 1-based
        // table-position lies inside the band.
        let band_lo = plan.idx;
        let band_hi = plan.idx + n - 1;
        let kept: Vec<String> = entries
            .into_iter()
            .enumerate()
            .filter_map(|(i, e)| {
                let col = t_lo + i as u32; // 1-based workbook col of this table column
                if col >= band_lo && col <= band_hi {
                    None
                } else {
                    Some(e)
                }
            })
            .collect();
        renumber_ids(kept)
    } else {
        renumber_ids(entries)
    };

    let new_count = new_entries.len();
    let mut new_block = format!(r#"<tableColumns count="{new_count}">"#);
    for e in &new_entries {
        new_block.push_str(e);
    }
    new_block.push_str("</tableColumns>");

    let mut out = Vec::with_capacity(shifted.len());
    out.extend_from_slice(s[..block_start].as_bytes());
    out.extend_from_slice(new_block.as_bytes());
    out.extend_from_slice(s[block_end..].as_bytes());
    out
}

/// Renumber `id="N"` on every `<tableColumn .../>` element so they're
/// 1, 2, 3, ... in source order.
fn renumber_ids(entries: Vec<String>) -> Vec<String> {
    entries
        .into_iter()
        .enumerate()
        .map(|(i, e)| {
            let id = (i + 1) as u32;
            replace_attr_value(&e, "id", &id.to_string())
        })
        .collect()
}

/// Replace the value of attribute `key` in `elt` (or insert if absent).
fn replace_attr_value(elt: &str, key: &str, new_val: &str) -> String {
    let pat = format!(" {key}=\"");
    if let Some(start) = elt.find(&pat) {
        let v_start = start + pat.len();
        if let Some(rel_end) = elt[v_start..].find('"') {
            let v_end = v_start + rel_end;
            let mut out = String::with_capacity(elt.len() + new_val.len());
            out.push_str(&elt[..v_start]);
            out.push_str(new_val);
            out.push_str(&elt[v_end..]);
            return out;
        }
    }
    // Insert before "/>" if absent.
    if let Some(close) = elt.rfind("/>") {
        let mut out = String::with_capacity(elt.len() + key.len() + new_val.len() + 4);
        out.push_str(&elt[..close]);
        out.push_str(&format!(" {key}=\"{new_val}\""));
        out.push_str(&elt[close..]);
        return out;
    }
    elt.to_string()
}

fn extract_attr(elt: &str, key: &str) -> Option<String> {
    let pat = format!(" {key}=\"");
    let start = elt.find(&pat)? + pat.len();
    let rel_end = elt[start..].find('"')?;
    Some(elt[start..start + rel_end].to_string())
}

/// Rewrite a `xl/comments*.xml` part: `<comment ref>`.
pub fn shift_comments_xml(xml: &[u8], plan: &ShiftPlan) -> Vec<u8> {
    if plan.is_noop() {
        return xml.to_vec();
    }
    let xml_str = match std::str::from_utf8(xml) {
        Ok(s) => s,
        Err(_) => return xml.to_vec(),
    };
    let mut reader = XmlReader::from_str(xml_str);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(std::io::Cursor::new(Vec::new()));
    let mut buf: Vec<u8> = Vec::new();
    let mut skip_depth: u32 = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name().as_ref().to_vec();
                if skip_depth > 0 {
                    skip_depth += 1;
                    buf.clear();
                    continue;
                }
                if local.as_slice() == b"comment" {
                    let mut keep = true;
                    let mut new_e =
                        BytesStart::new(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                    for attr_res in e.attributes().with_checks(false) {
                        let Ok(attr) = attr_res else { continue };
                        let key = attr.key.as_ref();
                        let val = match attr.unescape_value() {
                            Ok(v) => v.into_owned(),
                            Err(_) => continue,
                        };
                        if key == b"ref" {
                            let new_val = shift_anchor(&val, plan);
                            if new_val == "#REF!" {
                                keep = false;
                            }
                            push_attr(&mut new_e, key, &new_val);
                        } else {
                            push_attr(&mut new_e, key, &val);
                        }
                    }
                    if keep {
                        let _ = writer.write_event(Event::Start(new_e));
                    } else {
                        skip_depth = 1;
                    }
                } else {
                    let _ = writer.write_event(Event::Start(e.to_owned()));
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name().as_ref().to_vec();
                if skip_depth > 0 {
                    buf.clear();
                    continue;
                }
                if local.as_slice() == b"comment" {
                    let mut keep = true;
                    let mut new_e =
                        BytesStart::new(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                    for attr_res in e.attributes().with_checks(false) {
                        let Ok(attr) = attr_res else { continue };
                        let key = attr.key.as_ref();
                        let val = match attr.unescape_value() {
                            Ok(v) => v.into_owned(),
                            Err(_) => continue,
                        };
                        if key == b"ref" {
                            let new_val = shift_anchor(&val, plan);
                            if new_val == "#REF!" {
                                keep = false;
                            }
                            push_attr(&mut new_e, key, &new_val);
                        } else {
                            push_attr(&mut new_e, key, &val);
                        }
                    }
                    if keep {
                        let _ = writer.write_event(Event::Empty(new_e));
                    }
                } else {
                    let _ = writer.write_event(Event::Empty(e.to_owned()));
                }
            }
            Ok(Event::End(ref e)) => {
                if skip_depth > 0 {
                    skip_depth -= 1;
                    buf.clear();
                    continue;
                }
                let local = e.local_name().as_ref().to_vec();
                let _ = writer.write_event(Event::End(BytesEnd::new(
                    String::from_utf8_lossy(local.as_slice()).into_owned(),
                )));
            }
            Ok(Event::Text(ref t)) => {
                if skip_depth > 0 {
                    buf.clear();
                    continue;
                }
                let _ = writer.write_event(Event::Text(t.to_owned()));
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                if skip_depth > 0 {
                    buf.clear();
                    continue;
                }
                let _ = writer.write_event(other);
            }
            Err(_) => break,
        }
        buf.clear();
    }

    writer.into_inner().into_inner()
}

/// Rewrite a VML drawing part. Looks for `<x:Anchor>` elements (or
/// `<Anchor>` without prefix) and shifts the row component (axis 1
/// in 0-based, i.e. the second integer in the comma-separated list).
///
/// VML structure:
///
/// ```vml
/// <x:Anchor>1, 0, 5, 0, 3, 0, 7, 0</x:Anchor>
///                ^row(0-based)        ^row-bottom(0-based)
/// ```
///
/// Eight ints: col-left, col-left-offset, row-top, row-top-offset,
/// col-right, col-right-offset, row-bottom, row-bottom-offset.
/// (Reference: Microsoft `[MS-OI29500]` VML anchor format.)
///
/// On a row shift we mutate ints #2 (row-top) and #6 (row-bottom).
/// On a col shift we mutate ints #0 (col-left) and #4 (col-right).
/// Tombstoned anchors collapse the entire VML shape (drop it).
pub fn shift_vml_xml(xml: &[u8], plan: &ShiftPlan) -> Vec<u8> {
    if plan.is_noop() {
        return xml.to_vec();
    }
    // VML is a different XML namespace and quick-xml parsing it
    // robustly would require special handling. Use a string-rewrite
    // pass: find `<x:Anchor>...</x:Anchor>` (or `<Anchor>...</Anchor>`),
    // mutate the int list. Track whether any anchor goes #REF! — if
    // so, drop the enclosing `<v:shape>` element.
    let s = match std::str::from_utf8(xml) {
        Ok(s) => s,
        Err(_) => return xml.to_vec(),
    };

    // Walk shapes in order. We assume the file is small (typical
    // commentsN VML is <10KB).
    let mut out = String::with_capacity(s.len());
    let mut cursor = 0;
    while cursor < s.len() {
        let rest = &s[cursor..];
        // Find next <v:shape ... or </v:shape>
        if let Some(shape_start_rel) = rest.find("<v:shape") {
            out.push_str(&rest[..shape_start_rel]);
            let shape_start = cursor + shape_start_rel;
            // Find matching </v:shape>
            let after_start = &s[shape_start..];
            let shape_end_rel = after_start.find("</v:shape>");
            let shape_end = match shape_end_rel {
                Some(e) => shape_start + e + "</v:shape>".len(),
                None => {
                    out.push_str(after_start);
                    break;
                }
            };
            let shape_block = &s[shape_start..shape_end];
            // Find anchor inside this shape.
            let new_block = match shift_vml_anchor_in_shape(shape_block, plan) {
                Some(b) => b,
                None => {
                    // Tombstone: drop the entire shape.
                    cursor = shape_end;
                    continue;
                }
            };
            out.push_str(&new_block);
            cursor = shape_end;
        } else {
            out.push_str(rest);
            break;
        }
    }
    out.into_bytes()
}

fn shift_vml_anchor_in_shape(shape: &str, plan: &ShiftPlan) -> Option<String> {
    // Find <x:Anchor>...</x:Anchor>
    let open = shape.find("<x:Anchor>")?;
    let close = shape.find("</x:Anchor>")?;
    if close <= open {
        return Some(shape.to_string());
    }
    let payload = &shape[open + "<x:Anchor>".len()..close];
    let new_payload = shift_vml_anchor_payload(payload, plan)?;
    let mut out = String::with_capacity(shape.len());
    out.push_str(&shape[..open]);
    out.push_str("<x:Anchor>");
    out.push_str(&new_payload);
    out.push_str(&shape[close..]);
    Some(out)
}

fn shift_vml_anchor_payload(payload: &str, plan: &ShiftPlan) -> Option<String> {
    let parts: Vec<&str> = payload.split(',').map(|p| p.trim()).collect();
    if parts.len() != 8 {
        return Some(payload.to_string());
    }
    let mut nums: Vec<i64> = Vec::with_capacity(8);
    for p in &parts {
        nums.push(p.parse::<i64>().ok()?);
    }
    let abs = plan.abs_n() as i64;
    match plan.axis {
        Axis::Row => {
            // Indices 2 and 6 are the row anchors (0-based).
            // Convert to 1-based row index for compare.
            for &i in &[2usize, 6usize] {
                let row1b = nums[i] + 1;
                if plan.is_insert() {
                    if row1b as u32 >= plan.idx {
                        nums[i] = nums[i] + plan.n as i64;
                    }
                } else {
                    if row1b as u32 >= plan.idx && (row1b as u32) < plan.idx + abs as u32 {
                        return None; // tombstone the shape
                    }
                    if row1b as u32 >= plan.idx + abs as u32 {
                        nums[i] = nums[i] + plan.n as i64;
                    }
                }
                if nums[i] < 0 {
                    return None;
                }
                if nums[i] >= crate::MAX_ROW as i64 {
                    return None;
                }
            }
        }
        Axis::Col => {
            for &i in &[0usize, 4usize] {
                let col1b = nums[i] + 1;
                if plan.is_insert() {
                    if col1b as u32 >= plan.idx {
                        nums[i] = nums[i] + plan.n as i64;
                    }
                } else {
                    if col1b as u32 >= plan.idx && (col1b as u32) < plan.idx + abs as u32 {
                        return None;
                    }
                    if col1b as u32 >= plan.idx + abs as u32 {
                        nums[i] = nums[i] + plan.n as i64;
                    }
                }
                if nums[i] < 0 {
                    return None;
                }
                if nums[i] >= crate::MAX_COL as i64 {
                    return None;
                }
            }
        }
    }
    Some(
        nums.iter()
            .map(|n| n.to_string())
            .collect::<Vec<_>>()
            .join(", "),
    )
}

/// Rewrite `<definedName>` text content inside `xl/workbook.xml`.
///
/// Workbook-scope names (no `localSheetId` attr) shift only refs that
/// target `op_sheet`. Per-sheet names (with `localSheetId="N"`) shift
/// only when `N == sheet_position[op_sheet]`.
pub fn shift_defined_names(
    xml: &[u8],
    plan: &ShiftPlan,
    op_sheet: &str,
    op_sheet_position: Option<u32>,
) -> Vec<u8> {
    if plan.is_noop() {
        return xml.to_vec();
    }
    let xml_str = match std::str::from_utf8(xml) {
        Ok(s) => s,
        Err(_) => return xml.to_vec(),
    };
    let mut reader = XmlReader::from_str(xml_str);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(std::io::Cursor::new(Vec::new()));
    let mut buf: Vec<u8> = Vec::new();

    let mut in_dn = false;
    let mut current_local_sheet_id: Option<u32> = None;
    // Buffer the start tag so we can decide to apply or skip when we
    // hit the text event.

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name().as_ref().to_vec();
                if local.as_slice() == b"definedName" {
                    in_dn = true;
                    current_local_sheet_id = None;
                    for attr_res in e.attributes().with_checks(false) {
                        let Ok(attr) = attr_res else { continue };
                        if attr.key.as_ref() == b"localSheetId" {
                            if let Ok(v) = attr.unescape_value() {
                                if let Ok(n) = v.parse::<u32>() {
                                    current_local_sheet_id = Some(n);
                                }
                            }
                        }
                    }
                }
                let _ = writer.write_event(Event::Start(e.to_owned()));
            }
            Ok(Event::Empty(ref e)) => {
                let _ = writer.write_event(Event::Empty(e.to_owned()));
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name().as_ref().to_vec();
                if local.as_slice() == b"definedName" {
                    in_dn = false;
                    current_local_sheet_id = None;
                }
                let _ = writer.write_event(Event::End(BytesEnd::new(
                    String::from_utf8_lossy(local.as_slice()).into_owned(),
                )));
            }
            Ok(Event::Text(ref t)) => {
                if in_dn {
                    // Decide if this DN's scope matches our op.
                    let applies = match (current_local_sheet_id, op_sheet_position) {
                        // Per-sheet name: only shift if the localSheetId
                        // matches the op sheet's position.
                        (Some(sid), Some(pos)) => sid == pos,
                        // Workbook-scope name: always thread through
                        // shift_formula_on_sheet so refs to other sheets
                        // are left alone.
                        (None, _) => true,
                        // Per-sheet name but no sheet_position info:
                        // be conservative and don't shift.
                        (Some(_), None) => false,
                    };
                    if applies {
                        let s = match t.unescape() {
                            Ok(c) => c.into_owned(),
                            Err(_) => String::from_utf8_lossy(t.as_ref()).into_owned(),
                        };
                        let new_s = if current_local_sheet_id.is_some() {
                            // For per-sheet names, the formula's
                            // unqualified refs implicitly belong to
                            // the local sheet; treat them as on op_sheet.
                            shift_formula_on_sheet(&s, plan, op_sheet)
                        } else {
                            // Workbook-scope: refs MUST be qualified; the
                            // formula translator handles sheet matching
                            // via sheet_renames-empty + formula_sheet=None
                            // semantics — but to scope by sheet name we
                            // route through shift_formula_on_sheet too.
                            shift_formula_on_sheet(&s, plan, op_sheet)
                        };
                        let new_t = BytesText::new(&new_s);
                        let _ = writer.write_event(Event::Text(new_t));
                    } else {
                        let _ = writer.write_event(Event::Text(t.to_owned()));
                    }
                } else {
                    let _ = writer.write_event(Event::Text(t.to_owned()));
                }
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                let _ = writer.write_event(other);
            }
            Err(_) => break,
        }
        buf.clear();
    }

    writer.into_inner().into_inner()
}

/// Helper: rewrite a single attribute on a `BytesStart` element by
/// applying a function to its value.
fn rewrite_attr_value<'a, F: Fn(&str) -> String>(
    e: &BytesStart<'a>,
    attr_name: &[u8],
    f: F,
) -> BytesStart<'a> {
    let mut new_e = BytesStart::new(String::from_utf8_lossy(e.name().as_ref()).into_owned());
    for attr_res in e.attributes().with_checks(false) {
        let Ok(attr) = attr_res else { continue };
        let key = attr.key.as_ref();
        let val = match attr.unescape_value() {
            Ok(v) => v.into_owned(),
            Err(_) => continue,
        };
        if key == attr_name {
            let new_val = f(&val);
            push_attr(&mut new_e, key, &new_val);
        } else {
            push_attr(&mut new_e, key, &val);
        }
    }
    new_e
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_ops_returns_empty_mutations() {
        let inputs = SheetXmlInputs::empty();
        let out = apply_workbook_shift(inputs, &[]);
        assert!(out.file_patches.is_empty());
    }

    #[test]
    fn shifts_sheet_xml_via_orchestrator() {
        let sheet_xml = r#"<sheetData><row r="5"><c r="A5"><v>1</v></c></row></sheetData>"#;
        let mut inputs = SheetXmlInputs::empty();
        inputs
            .sheets
            .insert("Sheet1".to_string(), sheet_xml.as_bytes());
        inputs
            .sheet_paths
            .insert("Sheet1".to_string(), "xl/worksheets/sheet1.xml".to_string());
        let ops = vec![AxisShiftOp {
            sheet: "Sheet1".to_string(),
            axis: Axis::Row,
            idx: 5,
            n: 3,
        }];
        let out = apply_workbook_shift(inputs, &ops);
        let new_sheet = out
            .file_patches
            .get("xl/worksheets/sheet1.xml")
            .expect("sheet was rewritten");
        let s = String::from_utf8_lossy(new_sheet);
        assert!(s.contains(r#"<row r="8">"#));
    }

    #[test]
    fn shifts_table_ref_and_autofilter() {
        let table_xml = r#"<?xml version="1.0"?><table xmlns="..." id="1" name="T" displayName="T" ref="A1:E10"><autoFilter ref="A1:E10"/><tableColumns count="1"><tableColumn id="1" name="X"><calculatedColumnFormula>A5</calculatedColumnFormula></tableColumn></tableColumns></table>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        let out = shift_table_xml(table_xml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(r#"ref="A1:E13""#));
        assert!(s.contains(r#"<autoFilter ref="A1:E13""#));
        assert!(s.contains("<calculatedColumnFormula>A8</calculatedColumnFormula>"));
    }

    #[test]
    fn shifts_comments_ref() {
        let xml = r#"<comments><commentList><comment ref="A5" authorId="0"><text><t>hi</t></text></comment></commentList></comments>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        let out = shift_comments_xml(xml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(r#"ref="A8""#));
    }

    #[test]
    fn drops_tombstoned_comment() {
        let xml = r#"<comments><commentList><comment ref="A5" authorId="0"><text><t>hi</t></text></comment></commentList></comments>"#;
        let p = ShiftPlan::delete(Axis::Row, 5, 1);
        let out = shift_comments_xml(xml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(!s.contains("<comment "));
    }

    #[test]
    fn shifts_vml_anchor_row() {
        // 8 ints: 0, 0, 4, 0, 2, 0, 6, 0
        // (col-left=0, col-left-off=0, row-top=4, row-top-off=0,
        //  col-right=2, col-right-off=0, row-bottom=6, row-bottom-off=0)
        // Insert 3 rows at idx=5 (0-based row=4 is 1-based row=5).
        // Both 2 (row-top) and 6 (row-bottom) are >= 5 in 1-based, so
        // both shift by +3 → row-top=7, row-bottom=9.
        let vml = r#"<v:shape><x:ClientData><x:Anchor>0, 0, 4, 0, 2, 0, 6, 0</x:Anchor></x:ClientData></v:shape>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        let out = shift_vml_xml(vml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains("0, 0, 7, 0, 2, 0, 9, 0"));
    }

    #[test]
    fn drops_vml_shape_when_anchor_tombstoned() {
        let vml = r#"<v:shape><x:ClientData><x:Anchor>0, 0, 4, 0, 2, 0, 4, 0</x:Anchor></x:ClientData></v:shape>"#;
        let p = ShiftPlan::delete(Axis::Row, 5, 1);
        let out = shift_vml_xml(vml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(!s.contains("v:shape"));
    }

    #[test]
    fn shifts_workbook_scope_defined_name() {
        let xml = r#"<workbook><definedNames><definedName name="Total">Sheet1!$A$5</definedName></definedNames></workbook>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        let out = shift_defined_names(xml.as_bytes(), &p, "Sheet1", Some(0));
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains("Sheet1!$A$8"));
    }

    #[test]
    fn skips_per_sheet_dn_for_other_sheet() {
        let xml = r#"<workbook><definedNames><definedName name="Total" localSheetId="1">A5</definedName></definedNames></workbook>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        // op_sheet is at position 0; DN's localSheetId is 1 → skip.
        let out = shift_defined_names(xml.as_bytes(), &p, "Sheet1", Some(0));
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(">A5<"));
    }

    #[test]
    fn shifts_per_sheet_dn_for_matching_sheet() {
        let xml = r#"<workbook><definedNames><definedName name="Total" localSheetId="0">A5</definedName></definedNames></workbook>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        let out = shift_defined_names(xml.as_bytes(), &p, "Sheet1", Some(0));
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(">A8<"));
    }

    #[test]
    fn passes_through_unrelated_workbook_xml() {
        let xml = r#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        let out = shift_defined_names(xml.as_bytes(), &p, "Sheet1", Some(0));
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(r#"sheet name="Sheet1""#));
    }

    // --- RFC-031 §5.4: <tableColumns> insert/delete on Col axis -----------

    const TBL_5COL: &str = r#"<?xml version="1.0"?><table id="1" name="T" displayName="T" ref="A1:E5"><autoFilter ref="A1:E5"/><tableColumns count="5"><tableColumn id="1" name="H1"/><tableColumn id="2" name="H2"/><tableColumn id="3" name="H3"/><tableColumn id="4" name="H4"/><tableColumn id="5" name="H5"/></tableColumns></table>"#;

    #[test]
    fn table_insert_inside_band_adds_columns_and_renumbers() {
        // insert_cols(3, 2) on table A1:E5 → ref A1:G5, count=7,
        // ids 1..7, names H1, H2, Column<NEW>, Column<NEW>, H3, H4, H5.
        let plan = ShiftPlan::insert(Axis::Col, 3, 2);
        let out = shift_table_xml(TBL_5COL.as_bytes(), &plan);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(r#"ref="A1:G5""#), "got: {s}");
        assert!(s.contains(r#"<tableColumns count="7">"#), "got: {s}");
        assert!(s.contains(r#"<tableColumn id="1" name="H1"/>"#), "got: {s}");
        assert!(s.contains(r#"<tableColumn id="2" name="H2"/>"#), "got: {s}");
        assert!(
            s.contains(r#"<tableColumn id="3" name="Column"#),
            "got: {s}"
        );
        assert!(
            s.contains(r#"<tableColumn id="4" name="Column"#),
            "got: {s}"
        );
        assert!(s.contains(r#"<tableColumn id="5" name="H3"/>"#), "got: {s}");
        assert!(s.contains(r#"<tableColumn id="7" name="H5"/>"#), "got: {s}");
    }

    #[test]
    fn table_delete_inside_band_removes_columns_and_renumbers() {
        // delete_cols(3, 2) on A1:E5 → ref A1:C5, count=3, ids 1,2,3
        // names H1, H2, H5 (H3, H4 dropped).
        let plan = ShiftPlan::delete(Axis::Col, 3, 2);
        let out = shift_table_xml(TBL_5COL.as_bytes(), &plan);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(r#"ref="A1:C5""#), "got: {s}");
        assert!(s.contains(r#"<tableColumns count="3">"#), "got: {s}");
        assert!(s.contains(r#"<tableColumn id="1" name="H1"/>"#), "got: {s}");
        assert!(s.contains(r#"<tableColumn id="2" name="H2"/>"#), "got: {s}");
        assert!(s.contains(r#"<tableColumn id="3" name="H5"/>"#), "got: {s}");
        assert!(!s.contains(r#"name="H3""#), "got: {s}");
        assert!(!s.contains(r#"name="H4""#), "got: {s}");
    }

    #[test]
    fn table_insert_after_band_does_not_change_columns() {
        // insert_cols(7, 2) on A1:E5 — entirely outside table → ref unchanged,
        // tableColumns unchanged (count stays 5).
        let plan = ShiftPlan::insert(Axis::Col, 7, 2);
        let out = shift_table_xml(TBL_5COL.as_bytes(), &plan);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(r#"ref="A1:E5""#), "got: {s}");
        assert!(s.contains(r#"<tableColumns count="5">"#), "got: {s}");
    }

    #[test]
    fn table_insert_before_band_shifts_ref_but_keeps_columns() {
        // insert_cols(1, 2) before the table — ref shifts to C1:G5, but
        // the table still has 5 cols (cols just renumbered in the workbook).
        let plan = ShiftPlan::insert(Axis::Col, 1, 2);
        let out = shift_table_xml(TBL_5COL.as_bytes(), &plan);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(r#"ref="C1:G5""#), "got: {s}");
        assert!(s.contains(r#"<tableColumns count="5">"#), "got: {s}");
    }

    #[test]
    fn row_axis_shift_does_not_touch_table_columns() {
        // insert_rows(3, 2) — table_columns block must be untouched
        // because it's a column-axis quirk.
        let plan = ShiftPlan::insert(Axis::Row, 3, 2);
        let out = shift_table_xml(TBL_5COL.as_bytes(), &plan);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(r#"<tableColumns count="5">"#), "got: {s}");
        assert!(s.contains(r#"<tableColumn id="3" name="H3"/>"#), "got: {s}");
        assert!(s.contains(r#"ref="A1:E7""#), "got: {s}");
    }

    #[test]
    fn parse_ref_col_band_handles_dollar_and_single_cell() {
        assert_eq!(parse_ref_col_band("A1:E5"), Some((1, 5)));
        assert_eq!(parse_ref_col_band("$A$1:$E$5"), Some((1, 5)));
        assert_eq!(parse_ref_col_band("B3"), Some((2, 2)));
        assert_eq!(parse_ref_col_band("AA1:AB1"), Some((27, 28)));
    }
}
