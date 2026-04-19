//! `wolfxl agent <file> --max-tokens N` — token-budgeted workbook briefing.
//!
//! The brief is composed greedily for an LLM context window. Priority order:
//! 1. Workbook overview (every sheet, dims + class + first-column header)
//! 2. Focused sheet header (chosen sheet, dims, class)
//! 3. Column headers for the focused sheet
//! 4. Head rows (first 3) — ordering tends to be authored, so head matters
//! 5. Tail rows (last 2) — totals/EPS/footnotes typically live at the bottom
//! 6. Stratified middle samples — fills remaining budget with diverse rows
//! 7. Footer line: `# wolfxl agent: USED/LIMIT tokens (cl100k_base)`
//!
//! Block 1+2+3 is the orientation core and is emitted even if the row blocks
//! get dropped. If even the orientation core overflows, we still emit it
//! (truncating the budget would defeat the purpose of telling the agent what
//! sheets exist) and the footer reports the overage so the caller knows.
//!
//! Token counts use `tiktoken-rs::cl100k_base` to match the GPT-4 family
//! tokenizer used by `spreadsheet-peek/benchmarks/measure_tokens.py`.

use std::io::{self, Write};
use std::path::PathBuf;

use anyhow::{Context, Result};
use tiktoken_rs::{cl100k_base_singleton, CoreBPE};
use wolfxl_core::{Cell, CellValue, Sheet, SheetClass, Workbook, WorkbookMap};

pub fn run(file: PathBuf, max_tokens: usize, target_sheet: Option<String>) -> Result<()> {
    let mut wb = Workbook::open(&file)
        .with_context(|| format!("opening workbook: {}", file.display()))?;
    let map = wb.map().context("building workbook map")?;
    let target = pick_target(&map, target_sheet.as_deref())?;
    let sheet = wb
        .sheet(&target)
        .with_context(|| format!("loading sheet {target:?}"))?;

    let bpe = cl100k_base_singleton();
    let budget = Budget::new(bpe, max_tokens);
    let mut buf = String::new();

    // Orientation core (overview + sheet header + columns) always lands,
    // even if it overflows — we'd rather report overage in the footer than
    // hide the workbook structure from the agent.
    write_overview(&mut buf, &map);
    write_sheet_header(&mut buf, &target, &sheet, &map);

    write_rows(&mut buf, &sheet, &budget);

    let footer = format!(
        "\n# wolfxl agent: {used}/{limit} tokens (cl100k_base)\n",
        used = budget.used(&buf),
        limit = max_tokens
    );
    buf.push_str(&footer);

    io::stdout().lock().write_all(buf.as_bytes())?;
    Ok(())
}

/// Pick the focus sheet. Explicit `--sheet` wins; otherwise prefer the
/// largest `Data` sheet by row count, falling back to the first sheet so
/// workbooks of all-summary or all-empty shape still produce output.
fn pick_target(map: &WorkbookMap, requested: Option<&str>) -> Result<String> {
    if let Some(name) = requested {
        if !map.sheets.iter().any(|s| s.name == name) {
            let names: Vec<&str> = map.sheets.iter().map(|s| s.name.as_str()).collect();
            anyhow::bail!("sheet {name:?} not found; available: {}", names.join(", "));
        }
        return Ok(name.to_string());
    }
    let best_data = map
        .sheets
        .iter()
        .filter(|s| s.class == SheetClass::Data)
        .max_by_key(|s| s.rows);
    if let Some(s) = best_data {
        return Ok(s.name.clone());
    }
    map.sheets
        .first()
        .map(|s| s.name.clone())
        .ok_or_else(|| anyhow::anyhow!("workbook has no sheets"))
}

fn write_overview(buf: &mut String, map: &WorkbookMap) {
    let path = std::path::Path::new(&map.path)
        .file_name()
        .and_then(|s| s.to_str())
        .unwrap_or(&map.path);
    buf.push_str(&format!("WORKBOOK: {path}\n"));
    buf.push_str("SHEETS:\n");
    for s in &map.sheets {
        let first_col = s.headers.first().map(String::as_str).unwrap_or("");
        if first_col.is_empty() {
            buf.push_str(&format!(
                "  - {name} ({rows}r×{cols}c) [{class}]\n",
                name = s.name,
                rows = s.rows,
                cols = s.cols,
                class = s.class.as_str()
            ));
        } else {
            buf.push_str(&format!(
                "  - {name} ({rows}r×{cols}c) [{class}] col1={first_col:?}\n",
                name = s.name,
                rows = s.rows,
                cols = s.cols,
                class = s.class.as_str()
            ));
        }
    }
    if !map.named_ranges.is_empty() {
        buf.push_str("NAMED_RANGES:\n");
        for (name, formula) in &map.named_ranges {
            buf.push_str(&format!("  - {name} = {formula}\n"));
        }
    }
}

fn write_sheet_header(buf: &mut String, target: &str, sheet: &Sheet, map: &WorkbookMap) {
    let (rows, cols) = sheet.dimensions();
    let class = map
        .sheets
        .iter()
        .find(|s| s.name == target)
        .map(|s| s.class.as_str())
        .unwrap_or("data");
    buf.push('\n');
    buf.push_str(&format!(
        "SHEET: {target}  [{class}]  {rows} rows × {cols} cols\n"
    ));
    let headers = sheet.headers();
    if !headers.is_empty() {
        buf.push_str("HEADERS: ");
        buf.push_str(&headers.join("\t"));
        buf.push('\n');
    }
}

/// Emit head + tail + middle stratified samples within the remaining budget.
/// The body is `sheet.rows()[1..]` (header is row 0).
fn write_rows(buf: &mut String, sheet: &Sheet, budget: &Budget) {
    let body_start = 1usize;
    let total = sheet.rows().len();
    if total <= body_start {
        return;
    }
    let body_count = total - body_start;
    let head_n = 3.min(body_count);
    let tail_n = if body_count > head_n { 2.min(body_count - head_n) } else { 0 };

    // Head: emit all-or-nothing as a labelled block. If the section overflows
    // we'd rather skip than half-emit (truncated rows lie about row count).
    let head_section = render_section(
        sheet,
        body_start,
        body_start + head_n,
        &format!("ROWS (head {head_n} of {body_count}):"),
    );
    try_append(buf, budget, &head_section);

    if tail_n > 0 {
        let start = total - tail_n;
        let tail_section = render_section(
            sheet,
            start,
            total,
            &format!("ROWS (tail {tail_n} of {body_count}):"),
        );
        try_append(buf, budget, &tail_section);
    }

    let middle_lo = body_start + head_n;
    let middle_hi = total.saturating_sub(tail_n);
    if middle_hi > middle_lo {
        let middle_count = middle_hi - middle_lo;
        emit_stratified_middle(buf, sheet, budget, middle_lo, middle_count);
    }
}

/// Stratification = uniform stride. For a 10-row middle and 4 picks we'd
/// take indices [0, 3, 6, 9]. Caller supplies `lo` (absolute row index) and
/// `middle_count` (the size of the middle window).
///
/// We try up to 8 picks (cap on visual noise + budget reasonableness) and
/// emit them one at a time. The header line lands first; if even the
/// header overflows the row block is skipped entirely.
fn emit_stratified_middle(
    buf: &mut String,
    sheet: &Sheet,
    budget: &Budget,
    lo: usize,
    middle_count: usize,
) {
    let target_picks = 8usize.min(middle_count);
    if target_picks == 0 {
        return;
    }

    let mut picks: Vec<usize> = Vec::with_capacity(target_picks);
    if target_picks == 1 {
        picks.push(lo + middle_count / 2);
    } else {
        for i in 0..target_picks {
            let off = (i as f64) * ((middle_count - 1) as f64) / ((target_picks - 1) as f64);
            let row_idx = lo + (off.round() as usize);
            if !picks.contains(&row_idx) {
                picks.push(row_idx);
            }
        }
    }

    let header = format!(
        "ROWS (middle stratified, up to {} of {}):\n",
        picks.len(),
        middle_count
    );
    if !try_append(buf, budget, &header) {
        return;
    }
    for idx in picks {
        let line = format!("  {}\n", row_as_tsv(&sheet.rows()[idx]));
        if !try_append(buf, budget, &line) {
            break;
        }
    }
}

fn render_section(sheet: &Sheet, lo: usize, hi: usize, label: &str) -> String {
    let mut out = String::new();
    out.push_str(label);
    out.push('\n');
    for row in &sheet.rows()[lo..hi] {
        out.push_str("  ");
        out.push_str(&row_as_tsv(row));
        out.push('\n');
    }
    out
}

/// Append `section` to `buf` only if doing so keeps the running token count
/// under budget. Returns whether the section landed.
///
/// We probe by counting `buf + section` rather than `used(buf) + count(section)`
/// because cl100k_base BPE merges across boundaries — a token boundary that
/// sits at the seam can collapse two pieces into one. Sub-additive merge
/// would let an additive check reject sections that would actually fit.
fn try_append(buf: &mut String, budget: &Budget, section: &str) -> bool {
    let candidate_used = budget.used_with(buf, section);
    if candidate_used > budget.limit {
        return false;
    }
    buf.push_str(section);
    true
}

fn row_as_tsv(row: &[Cell]) -> String {
    row.iter()
        .map(|c| display_cell(c))
        .collect::<Vec<_>>()
        .join("\t")
}

/// Compact text rendering for the agent brief. We deliberately do NOT
/// thousand-group integers here: every comma is a token boundary in
/// cl100k_base, so `1234567` is two tokens but `1,234,567` is five. The
/// agent doesn't need pretty output; it needs cheap output.
fn display_cell(cell: &Cell) -> String {
    match &cell.value {
        CellValue::Empty => String::new(),
        CellValue::String(s) => s.clone(),
        CellValue::Bool(b) => if *b { "true" } else { "false" }.to_string(),
        CellValue::Int(n) => n.to_string(),
        CellValue::Float(n) => {
            if n.is_finite() && n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{n:.0}")
            } else if n.is_finite() {
                format!("{n:.2}")
            } else {
                n.to_string()
            }
        }
        CellValue::Date(d) => d.format("%Y-%m-%d").to_string(),
        CellValue::DateTime(dt) => dt.format("%Y-%m-%dT%H:%M:%S").to_string(),
        CellValue::Time(t) => t.format("%H:%M:%S").to_string(),
        CellValue::Error(e) => format!("ERR:{e}"),
    }
}

/// Token budget tracker. Wraps `cl100k_base` and tracks the running total
/// of tokens emitted so far. `consume(s)` returns the token count of `s`
/// (so callers can decide whether to commit it); `remaining()` returns
/// what's left given the latest `used()` recompute.
///
/// Implementation note: we recount the whole accumulated buffer in
/// `used()` rather than incrementing per-add. BPE is *not* additive —
/// `tokens(a + b) <= tokens(a) + tokens(b)` because adjacent pieces can
/// merge into a single token across the boundary. Recomputing on the
/// final buffer is the only way to get a count that matches what
/// `tiktoken.cl100k_base.encode(full_output)` will produce.
struct Budget<'a> {
    bpe: &'a CoreBPE,
    limit: usize,
}

impl<'a> Budget<'a> {
    fn new(bpe: &'a CoreBPE, limit: usize) -> Self {
        Self { bpe, limit }
    }

    /// Total tokens an accumulated buffer would cost (matches what
    /// `tiktoken.cl100k_base.encode(buf)` returns in Python).
    fn used(&self, buf: &str) -> usize {
        self.bpe.encode_ordinary(buf).len()
    }

    /// Tokens of `buf + section`, encoded as one piece so cross-boundary
    /// BPE merges are counted correctly. Cheaper allocations would split
    /// the encode but lose merge accuracy at the seam.
    fn used_with(&self, buf: &str, section: &str) -> usize {
        // Re-encoding the full concatenation is the only way to get the
        // exact final count; BPE is non-additive at piece boundaries.
        let mut probe = String::with_capacity(buf.len() + section.len());
        probe.push_str(buf);
        probe.push_str(section);
        self.bpe.encode_ordinary(&probe).len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use wolfxl_core::SheetMap;

    #[test]
    fn budget_recomputes_against_full_buffer() {
        // BPE is non-additive across boundaries; `used_with` re-encodes
        // the full concatenation so cross-seam merges are counted. An
        // additive `count(a) + count(b)` would over-count and reject
        // sections that actually fit.
        let bpe = cl100k_base_singleton();
        let b = Budget::new(bpe, 100);
        let combined = b.used("hello world");
        let with = b.used_with("hello", " world");
        assert_eq!(combined, with);
    }

    #[test]
    fn pick_target_prefers_largest_data_sheet() {
        // Build a synthetic WorkbookMap with three sheets of different
        // sizes and classes; the largest-by-rows Data sheet wins even if
        // a Summary sheet has more rows.
        let map = WorkbookMap {
            path: "test.xlsx".to_string(),
            sheets: vec![
                SheetMap {
                    name: "Notes".to_string(),
                    rows: 50,
                    cols: 1,
                    class: SheetClass::Readme,
                    headers: vec!["note".to_string()],
                    tables: vec![],
                },
                SheetMap {
                    name: "P&L".to_string(),
                    rows: 21,
                    cols: 7,
                    class: SheetClass::Data,
                    headers: vec![],
                    tables: vec![],
                },
                SheetMap {
                    name: "Detail".to_string(),
                    rows: 200,
                    cols: 12,
                    class: SheetClass::Data,
                    headers: vec![],
                    tables: vec![],
                },
            ],
            named_ranges: vec![],
        };
        assert_eq!(pick_target(&map, None).unwrap(), "Detail");
    }

    #[test]
    fn pick_target_falls_back_to_first_sheet_when_no_data() {
        let map = WorkbookMap {
            path: "test.xlsx".to_string(),
            sheets: vec![
                SheetMap {
                    name: "Cover".to_string(),
                    rows: 5,
                    cols: 1,
                    class: SheetClass::Readme,
                    headers: vec![],
                    tables: vec![],
                },
                SheetMap {
                    name: "Summary".to_string(),
                    rows: 10,
                    cols: 4,
                    class: SheetClass::Summary,
                    headers: vec![],
                    tables: vec![],
                },
            ],
            named_ranges: vec![],
        };
        assert_eq!(pick_target(&map, None).unwrap(), "Cover");
    }

    #[test]
    fn pick_target_explicit_wins_over_heuristic() {
        let map = WorkbookMap {
            path: "test.xlsx".to_string(),
            sheets: vec![SheetMap {
                name: "P&L".to_string(),
                rows: 21,
                cols: 7,
                class: SheetClass::Data,
                headers: vec![],
                tables: vec![],
            }],
            named_ranges: vec![],
        };
        assert_eq!(pick_target(&map, Some("P&L")).unwrap(), "P&L");
        assert!(pick_target(&map, Some("Nope")).is_err());
    }
}
