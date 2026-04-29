//! Table XML rewrites for worksheet structural shifts.

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

use crate::axis::{Axis, ShiftPlan};
use crate::shift_anchors::shift_anchor;
use crate::shift_formulas::shift_formula;

fn push_attr<'a>(e: &mut BytesStart<'a>, key: &[u8], val: &str) {
    e.push_attribute((key, val.as_bytes()));
}

/// Rewrite a `xl/tables/tableN.xml` part: `<table ref>`, `<autoFilter ref>`,
/// `<calculatedColumnFormula>` text, plus the `<tableColumns>` block on
/// Col-axis shifts. Inserts spawn new `<tableColumn>` entries, deletes remove
/// them, and `count=` / `id=` are renumbered.
pub(crate) fn shift_table_xml(xml: &[u8], plan: &ShiftPlan) -> Vec<u8> {
    if plan.is_noop() {
        return xml.to_vec();
    }
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

/// Parse the `<table ref="A1:E5">` attribute and return the 1-based
/// `[col_lo, col_hi]` band.
fn extract_table_col_band(xml: &[u8]) -> Option<(u32, u32)> {
    let s = std::str::from_utf8(xml).ok()?;
    let i = s.find("<table ")?;
    let close = s[i..].find('>')?;
    let elt = &s[i..i + close];
    let r_idx = elt.find(" ref=\"")?;
    let v_start = r_idx + " ref=\"".len();
    let v_end = elt[v_start..].find('"')?;
    let r = &elt[v_start..v_start + v_end];
    parse_ref_col_band(r)
}

/// Parse "A1:E5" into `(1, 5)`. Single-cell refs return the same column twice.
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

    let mut entries: Vec<String> = Vec::new();
    let mut i = 0usize;
    while let Some(start_rel) = block[i..].find("<tableColumn ") {
        let abs_start = i + start_rel;
        let end = match block[abs_start..].find("/>") {
            Some(e) => abs_start + e + 2,
            None => break,
        };
        entries.push(block[abs_start..end].to_string());
        i = end;
    }

    let new_entries: Vec<String> = if plan.is_insert() {
        let n = plan.n as u32;
        if plan.idx > t_hi || plan.idx <= t_lo {
            renumber_ids(entries)
        } else {
            let insert_pos = (plan.idx - t_lo) as usize;
            let mut out: Vec<String> = Vec::with_capacity(entries.len() + n as usize);
            out.extend_from_slice(&entries[..insert_pos]);
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
        let band_lo = plan.idx;
        let band_hi = plan.idx + n - 1;
        let kept: Vec<String> = entries
            .into_iter()
            .enumerate()
            .filter_map(|(i, e)| {
                let col = t_lo + i as u32;
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
    fn shifts_table_ref_and_autofilter() {
        let table_xml = r#"<?xml version="1.0"?><table xmlns="..." id="1" name="T" displayName="T" ref="A1:E10"><autoFilter ref="A1:E10"/><tableColumns count="1"><tableColumn id="1" name="X"><calculatedColumnFormula>A5</calculatedColumnFormula></tableColumn></tableColumns></table>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        let out = shift_table_xml(table_xml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(r#"ref="A1:E13""#));
        assert!(s.contains(r#"<autoFilter ref="A1:E13""#));
        assert!(s.contains("<calculatedColumnFormula>A8</calculatedColumnFormula>"));
    }

    const TBL_5COL: &str = r#"<?xml version="1.0"?><table id="1" name="T" displayName="T" ref="A1:E5"><autoFilter ref="A1:E5"/><tableColumns count="5"><tableColumn id="1" name="H1"/><tableColumn id="2" name="H2"/><tableColumn id="3" name="H3"/><tableColumn id="4" name="H4"/><tableColumn id="5" name="H5"/></tableColumns></table>"#;

    #[test]
    fn table_insert_inside_band_adds_columns_and_renumbers() {
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
        let plan = ShiftPlan::insert(Axis::Col, 7, 2);
        let out = shift_table_xml(TBL_5COL.as_bytes(), &plan);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(r#"ref="A1:E5""#), "got: {s}");
        assert!(s.contains(r#"<tableColumns count="5">"#), "got: {s}");
    }

    #[test]
    fn table_insert_before_band_shifts_ref_but_keeps_columns() {
        let plan = ShiftPlan::insert(Axis::Col, 1, 2);
        let out = shift_table_xml(TBL_5COL.as_bytes(), &plan);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(r#"ref="C1:G5""#), "got: {s}");
        assert!(s.contains(r#"<tableColumns count="5">"#), "got: {s}");
    }

    #[test]
    fn row_axis_shift_does_not_touch_table_columns() {
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
