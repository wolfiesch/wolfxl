//! VML drawing anchor shifts for comment shapes.

use crate::axis::{Axis, ShiftPlan};

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
        if let Some(shape_start_rel) = rest.find("<v:shape") {
            out.push_str(&rest[..shape_start_rel]);
            let shape_start = cursor + shape_start_rel;
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
            let new_block = match shift_vml_anchor_in_shape(shape_block, plan) {
                Some(b) => b,
                None => {
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
            for &i in &[2usize, 6usize] {
                let row1b = nums[i] + 1;
                if plan.is_insert() {
                    if row1b as u32 >= plan.idx {
                        nums[i] += plan.n as i64;
                    }
                } else {
                    if row1b as u32 >= plan.idx && (row1b as u32) < plan.idx + abs as u32 {
                        return None;
                    }
                    if row1b as u32 >= plan.idx + abs as u32 {
                        nums[i] += plan.n as i64;
                    }
                }
                if nums[i] < 0 || nums[i] >= crate::MAX_ROW as i64 {
                    return None;
                }
            }
        }
        Axis::Col => {
            for &i in &[0usize, 4usize] {
                let col1b = nums[i] + 1;
                if plan.is_insert() {
                    if col1b as u32 >= plan.idx {
                        nums[i] += plan.n as i64;
                    }
                } else {
                    if col1b as u32 >= plan.idx && (col1b as u32) < plan.idx + abs as u32 {
                        return None;
                    }
                    if col1b as u32 >= plan.idx + abs as u32 {
                        nums[i] += plan.n as i64;
                    }
                }
                if nums[i] < 0 || nums[i] >= crate::MAX_COL as i64 {
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
