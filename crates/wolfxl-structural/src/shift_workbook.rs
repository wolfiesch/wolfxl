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
        ShiftPlan { axis: self.axis, idx: self.idx, n: self.n }
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
pub fn apply_workbook_shift(
    inputs: SheetXmlInputs<'_>,
    ops: &[AxisShiftOp],
) -> WorkbookMutations {
    let mut out = WorkbookMutations::default();

    if ops.is_empty() {
        return out;
    }

    // Cache rewritten sheet bytes across ops in this call so a single
    // multi-op call sees its own rewrites. (In practice the patcher
    // calls one op at a time, but this future-proofs the API.)
    let mut sheet_bytes: BTreeMap<String, Vec<u8>> =
        inputs.sheets.iter().map(|(k, v)| (k.clone(), v.to_vec())).collect();
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
                out.file_patches.insert("xl/workbook.xml".to_string(), bytes.clone());
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
/// and `<calculatedColumnFormula>` text.
pub fn shift_table_xml(xml: &[u8], plan: &ShiftPlan) -> Vec<u8> {
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
                    let mut new_e = BytesStart::new(
                        String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                    );
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
                    let mut new_e = BytesStart::new(
                        String::from_utf8_lossy(e.name().as_ref()).into_owned(),
                    );
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
        inputs.sheets.insert("Sheet1".to_string(), sheet_xml.as_bytes());
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
}
