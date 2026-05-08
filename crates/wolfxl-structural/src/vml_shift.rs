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
        if let Some(shape_start_rel) = find_shape_start(rest) {
            out.push_str(&rest[..shape_start_rel]);
            let shape_start = cursor + shape_start_rel;
            let after_start = &s[shape_start..];
            let close_tag = if after_start.starts_with("<v:shape") {
                "</v:shape>"
            } else {
                "</shape>"
            };
            let shape_end_rel = after_start.find(close_tag);
            let shape_end = match shape_end_rel {
                Some(e) => shape_start + e + close_tag.len(),
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

fn find_shape_start(s: &str) -> Option<usize> {
    let mut cursor = 0;
    while cursor < s.len() {
        let rest = &s[cursor..];
        let prefixed = rest.find("<v:shape");
        let unprefixed = rest.find("<shape");
        let next = match (prefixed, unprefixed) {
            (Some(a), Some(b)) => a.min(b),
            (Some(a), None) | (None, Some(a)) => a,
            (None, None) => return None,
        };
        let absolute = cursor + next;
        let candidate = &s[absolute..];
        let tag_len = if candidate.starts_with("<v:shape") {
            "<v:shape".len()
        } else {
            "<shape".len()
        };
        let next_char = candidate[tag_len..].chars().next();
        if matches!(next_char, Some(' ' | '\t' | '\r' | '\n' | '>' | '/')) {
            return Some(absolute);
        }
        cursor = absolute + tag_len;
    }
    None
}

fn shift_vml_anchor_in_shape(shape: &str, plan: &ShiftPlan) -> Option<String> {
    let (open_tag, close_tag) = if shape.contains("<x:Anchor>") {
        ("<x:Anchor>", "</x:Anchor>")
    } else if shape.contains("<Anchor>") {
        ("<Anchor>", "</Anchor>")
    } else {
        return Some(shape.to_string());
    };
    let open = shape.find(open_tag)?;
    let close = shape.find(close_tag)?;
    if close <= open {
        return Some(shape.to_string());
    }
    let payload = &shape[open + open_tag.len()..close];
    let new_payload = shift_vml_anchor_payload(payload, plan)?;
    let mut out = String::with_capacity(shape.len());
    out.push_str(&shape[..open]);
    out.push_str(open_tag);
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
    match plan.axis {
        Axis::Row => {
            if plan.is_insert() {
                for &i in &[2usize, 6usize] {
                    if nums[i] + 1 >= plan.idx as i64 {
                        nums[i] += plan.n as i64;
                    }
                }
            } else if !shift_deleted_anchor_span(&mut nums, 2, 6, plan) {
                return None;
            }
            if nums[2] < 0
                || nums[6] < 0
                || nums[2] >= crate::MAX_ROW as i64
                || nums[6] >= crate::MAX_ROW as i64
                || nums[6] < nums[2]
            {
                return None;
            }
        }
        Axis::Col => {
            if plan.is_insert() {
                for &i in &[0usize, 4usize] {
                    if nums[i] + 1 >= plan.idx as i64 {
                        nums[i] += plan.n as i64;
                    }
                }
            } else if !shift_deleted_anchor_span(&mut nums, 0, 4, plan) {
                return None;
            }
            if nums[0] < 0
                || nums[4] < 0
                || nums[0] >= crate::MAX_COL as i64
                || nums[4] >= crate::MAX_COL as i64
                || nums[4] < nums[0]
            {
                return None;
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

fn shift_deleted_anchor_span(
    nums: &mut [i64],
    start_idx: usize,
    end_idx: usize,
    plan: &ShiftPlan,
) -> bool {
    let delete_start = plan.idx as i64 - 1;
    let delete_end = delete_start + plan.abs_n() as i64;
    let start = nums[start_idx];
    let end = nums[end_idx];
    let start_deleted = start >= delete_start && start < delete_end;
    let end_deleted = end >= delete_start && end < delete_end;
    if start_deleted && end_deleted {
        return false;
    }

    if start_deleted {
        nums[start_idx] = delete_start;
    } else if nums[start_idx] >= delete_end {
        nums[start_idx] += plan.n as i64;
    }
    if end_deleted {
        nums[end_idx] = delete_start.saturating_sub(1);
    } else if nums[end_idx] >= delete_end {
        nums[end_idx] += plan.n as i64;
    }
    true
}
