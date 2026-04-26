//! wolfxl-writer — native xlsx writer for wolfxl.
//!
//! # Design
//!
//! A pure-Rust OOXML emitter that replaces `rust_xlsxwriter`. The writer is
//! split into three layers:
//!
//! 1. [`model`] — pure data (Workbook, Worksheet, Row, WriteCell, format specs).
//!    No I/O, no XML. You build a model, then hand it to the emitter.
//! 2. [`emit`] — one module per OOXML part (styles.xml, sheet1.xml, workbook.xml, ...).
//!    Each module takes the relevant slice of the model and returns UTF-8 bytes.
//! 3. [`zip`] — deterministic ZIP packager that assembles emitted parts into a
//!    valid xlsx container.
//!
//! The top-level [`Workbook`] facade orchestrates all three.
//!
//! # Determinism
//!
//! Byte-identical output is a non-goal for shipping but a gold-star target for
//! the differential test harness. To get close:
//!
//! - `WOLFXL_TEST_EPOCH=0` env var → ZIP entry mtimes are forced to the Unix
//!   epoch (1970-01-01) so two runs produce identical bytes.
//! - `BTreeMap` for row/cell collections → emission order matches the OOXML
//!   spec's "sorted ascending by `r` attribute" rule without an extra sort pass.
//! - `IndexMap` for author lists → preserves insertion order (fixes the
//!   `rust_xlsxwriter` BTreeMap bug that corrupted mixed-author comment files).
//!
//! # Status
//!
//! Skeleton — public surface is stubbed while Wave 1 subagents fill in
//! `refs` and `zip`. Full usage docs arrive at Wave 4 integration.

pub mod emit;
pub mod intern;
pub mod model;
pub mod refs;
pub mod rich_text;
pub mod xml_escape;
pub mod zip;

#[cfg(test)]
mod test_utils;

pub use model::workbook::Workbook;
pub use model::worksheet::Worksheet;

/// Render a complete `.xlsx` archive from the in-memory workbook.
///
/// This is the W4 entry point used by the `NativeWorkbook` pyclass and by
/// the `tests/roundtrip_minimal.rs` integration test. It emits every Wave
/// 1+2+3 part, packages them in canonical ZIP order, and returns the raw
/// bytes — the caller writes them to disk (or hands them to a diff tool).
///
/// Mutates `wb` because sheet emission interns strings into the workbook's
/// shared-string table; SST emission has to run AFTER all sheets to see
/// the final intern set. The mutation is monotonic (only string indices
/// are added) — calling `emit_xlsx` twice on the same workbook produces
/// the same archive both times.
pub fn emit_xlsx(wb: &mut Workbook) -> Vec<u8> {
    use crate::emit::{
        calc_chain_xml, comments_xml, content_types, doc_props, drawings_vml, rels,
        shared_strings_xml, sheet_xml, styles_xml, tables_xml, workbook_xml,
    };
    use crate::zip::{package, ZipEntry};

    // Sheet emission mutates the SST — must run before the SST emitter.
    let mut sheet_parts: Vec<(String, Vec<u8>)> = Vec::new();
    for (idx, sheet) in wb.sheets.iter().enumerate() {
        let bytes = sheet_xml::emit(sheet, idx as u32, &mut wb.sst, &wb.styles);
        sheet_parts.push((format!("xl/worksheets/sheet{}.xml", idx + 1), bytes));
    }

    let mut entries: Vec<ZipEntry> = vec![
        ZipEntry {
            path: "[Content_Types].xml".to_string(),
            bytes: content_types::emit(wb),
        },
        ZipEntry {
            path: "_rels/.rels".to_string(),
            bytes: rels::emit_root(wb),
        },
        ZipEntry {
            path: "xl/workbook.xml".to_string(),
            bytes: workbook_xml::emit(wb),
        },
        ZipEntry {
            path: "xl/_rels/workbook.xml.rels".to_string(),
            bytes: rels::emit_workbook(wb),
        },
    ];
    for (path, bytes) in sheet_parts {
        entries.push(ZipEntry { path, bytes });
    }
    for idx in 0..wb.sheets.len() {
        let sheet_rels = rels::emit_sheet(wb, idx);
        if !sheet_rels.is_empty() {
            entries.push(ZipEntry {
                path: format!("xl/worksheets/_rels/sheet{}.xml.rels", idx + 1),
                bytes: sheet_rels,
            });
        }
    }

    // Wave 3 rich-feature parts: emit only for sheets that actually use them.
    // Tables are globally numbered (table1.xml, table2.xml, ...) across all
    // sheets — this counter mirrors the rels allocation logic in emit_sheet.
    let mut global_table_idx: usize = 1;
    for (idx, sheet) in wb.sheets.iter().enumerate() {
        if !sheet.comments.is_empty() {
            entries.push(ZipEntry {
                path: format!("xl/comments/comments{}.xml", idx + 1),
                bytes: comments_xml::emit(sheet, &wb.comment_authors),
            });
            entries.push(ZipEntry {
                path: format!("xl/drawings/vmlDrawing{}.vml", idx + 1),
                bytes: drawings_vml::emit(sheet),
            });
        }
        for table in &sheet.tables {
            // The path counter is 1-based (table1.xml, table2.xml, …);
            // `tables_xml::emit` interprets its third argument as a
            // 0-based `table_idx` and emits `id="(idx+1)"`. Subtract
            // one so the emitted `id` matches the part filename
            // (RFC-024 cross-mode parity gate; modify-mode allocates
            // workbook-unique ids starting at 1, which matches the
            // emit fn's contract from
            // `tables_xml::tests::table_id_uses_idx_plus_one`).
            entries.push(ZipEntry {
                path: format!("xl/tables/table{}.xml", global_table_idx),
                bytes: tables_xml::emit(table, idx, global_table_idx - 1),
            });
            global_table_idx += 1;
        }
    }

    entries.extend([
        ZipEntry {
            path: "xl/styles.xml".to_string(),
            bytes: styles_xml::emit(&wb.styles),
        },
        ZipEntry {
            path: "xl/sharedStrings.xml".to_string(),
            bytes: shared_strings_xml::emit(&wb.sst),
        },
        ZipEntry {
            path: "docProps/core.xml".to_string(),
            bytes: doc_props::emit_core(wb),
        },
        ZipEntry {
            path: "docProps/app.xml".to_string(),
            bytes: doc_props::emit_app(wb),
        },
    ]);

    // Sprint Θ Pod-C3: write-mode calcChain. Only emit when the
    // workbook has at least one formula cell — Excel transparently
    // works without it, so an empty workbook should ship without the
    // part (matching openpyxl's behaviour).
    if let Some(cc_bytes) = calc_chain_xml::emit(wb) {
        entries.push(ZipEntry {
            path: "xl/calcChain.xml".to_string(),
            bytes: cc_bytes,
        });
    }

    package(&entries).expect("zip package")
}
