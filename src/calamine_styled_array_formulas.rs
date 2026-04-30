//! Array/data-table formula parsing for the styled calamine backend.

use std::collections::HashMap;

use pyo3::exceptions::PyIOError;
use pyo3::prelude::*;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;

use crate::ooxml_util;
use crate::util::a1_to_row_col;

/// RFC-057 (Sprint Ο Pod 1C) — array / data-table formula metadata for one
/// cell.
#[derive(Clone, Debug)]
pub(crate) enum ArrayFormulaInfo {
    /// Master cell of an array formula (`<f t="array" ref="...">`).
    Array { ref_range: String, text: String },
    /// Master cell of a data-table formula (`<f t="dataTable" ref="..."/>`).
    DataTable {
        ref_range: String,
        ca: bool,
        dt2_d: bool,
        dtr: bool,
        r1: Option<String>,
        r2: Option<String>,
    },
    /// Cell inside the spill range of an array formula but not the master.
    SpillChild,
}

/// Parse array-formula / data-table-formula metadata from worksheet XML.
///
/// Walks every ``<c>`` looking for a child ``<f>`` whose ``t`` attribute is
/// ``"array"`` or ``"dataTable"``. When found, every cell inside the ``ref``
/// range becomes a ``SpillChild`` entry except the master cell itself, which
/// gets the typed ``Array`` / ``DataTable`` payload.
pub(crate) fn parse_array_formulas_from_sheet_xml(
    xml: &str,
) -> PyResult<HashMap<(u32, u32), ArrayFormulaInfo>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf: Vec<u8> = Vec::new();
    let mut out: HashMap<(u32, u32), ArrayFormulaInfo> = HashMap::new();

    // Track state per-cell so we can correlate the parent <c r="...">
    // with the child <f t="array" ref="...">...</f> body.
    let mut current_cell: Option<(u32, u32)> = None;
    let mut in_array_formula: bool = false;
    let mut array_text: String = String::new();
    let mut pending_master: Option<((u32, u32), String)> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"c" {
                    let a1 = ooxml_util::attr_value(e, b"r").unwrap_or_default();
                    current_cell = if a1.is_empty() {
                        None
                    } else {
                        a1_to_row_col(&a1).ok()
                    };
                } else if name.as_ref() == b"f" {
                    let t_attr = ooxml_util::attr_value(e, b"t").unwrap_or_default();
                    if t_attr == "array" {
                        let ref_range = ooxml_util::attr_value(e, b"ref").unwrap_or_default();
                        in_array_formula = true;
                        array_text.clear();
                        if let Some(pos) = current_cell {
                            pending_master = Some((pos, ref_range));
                        }
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"f" {
                    let t_attr = ooxml_util::attr_value(e, b"t").unwrap_or_default();
                    if t_attr == "dataTable" {
                        // Self-closing `<f t="dataTable" ref="..." dt2D="..." r1="..." r2="..."/>`.
                        let ref_range = ooxml_util::attr_value(e, b"ref").unwrap_or_default();
                        let ca = parse_attr_bool(e, b"ca");
                        let dt2_d = parse_attr_bool(e, b"dt2D");
                        let dtr = parse_attr_bool(e, b"dtr");
                        let r1 = ooxml_util::attr_value(e, b"r1");
                        let r2 = ooxml_util::attr_value(e, b"r2");
                        if let Some(pos) = current_cell {
                            let info = ArrayFormulaInfo::DataTable {
                                ref_range: ref_range.clone(),
                                ca,
                                dt2_d,
                                dtr,
                                r1,
                                r2,
                            };
                            out.insert(pos, info);
                            mark_spill_children(&mut out, &ref_range, pos);
                        }
                    } else if t_attr == "array" {
                        // Self-closing `<f t="array" ref="..."/>` happens
                        // when the formula body is empty.
                        let ref_range = ooxml_util::attr_value(e, b"ref").unwrap_or_default();
                        if let Some(pos) = current_cell {
                            out.insert(
                                pos,
                                ArrayFormulaInfo::Array {
                                    ref_range: ref_range.clone(),
                                    text: String::new(),
                                },
                            );
                            mark_spill_children(&mut out, &ref_range, pos);
                        }
                    }
                } else if name.as_ref() == b"c" {
                    current_cell = None;
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"f" && in_array_formula {
                    in_array_formula = false;
                    if let Some((pos, ref_range)) = pending_master.take() {
                        out.insert(
                            pos,
                            ArrayFormulaInfo::Array {
                                ref_range: ref_range.clone(),
                                text: array_text.clone(),
                            },
                        );
                        mark_spill_children(&mut out, &ref_range, pos);
                    }
                } else if name.as_ref() == b"c" {
                    current_cell = None;
                }
            }
            Ok(Event::Text(ref t)) if in_array_formula => {
                if let Ok(text) = t.unescape() {
                    array_text.push_str(&text);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(PyErr::new::<PyIOError, _>(format!(
                    "Failed to parse worksheet XML for array formulas: {e}"
                )))
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

/// Parse an XML attribute as an Excel-style boolean.
///
/// OOXML uses `"1"`/`"0"` and `"true"`/`"false"` interchangeably for boolean
/// attributes; openpyxl also accepts the latter.
fn parse_attr_bool(e: &BytesStart<'_>, key: &[u8]) -> bool {
    let v = ooxml_util::attr_value(e, key).unwrap_or_default();
    matches!(v.as_str(), "1" | "true" | "True")
}

/// Tag every cell inside ``ref_range`` except ``master`` as ``SpillChild``.
fn mark_spill_children(
    out: &mut HashMap<(u32, u32), ArrayFormulaInfo>,
    ref_range: &str,
    master: (u32, u32),
) {
    let parts: Vec<&str> = ref_range.split(':').collect();
    if parts.is_empty() {
        return;
    }
    let top_left = parts[0];
    let bottom_right = if parts.len() > 1 { parts[1] } else { top_left };
    let Ok((r1, c1)) = a1_to_row_col(top_left) else {
        return;
    };
    let Ok((r2, c2)) = a1_to_row_col(bottom_right) else {
        return;
    };
    let (top, bottom) = if r1 < r2 { (r1, r2) } else { (r2, r1) };
    let (left, right) = if c1 < c2 { (c1, c2) } else { (c2, c1) };
    for r in top..=bottom {
        for c in left..=right {
            if (r, c) == master {
                continue;
            }
            out.entry((r, c)).or_insert(ArrayFormulaInfo::SpillChild);
        }
    }
}
