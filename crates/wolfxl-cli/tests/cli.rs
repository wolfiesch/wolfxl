//! Integration tests for `wolfxl peek`.
//!
//! Goldens lock in the rendering contract for `text`, `csv`, and `json`
//! exporters: integer thousand-grouping, two-decimal float formatting,
//! ISO date/time, RFC 4180 CSV quoting, and the JSON `{sheet, rows,
//! columns, headers, data}` shape. If you change a renderer and a golden
//! moves, either regenerate the golden intentionally and update the token
//! tables in `spreadsheet-peek`, or fix the bug.
//!
//! The boxed renderer isn't goldened - its banner is wolfxl-branded and
//! column widths can drift on minor rendering tweaks. Smoke-tested
//! instead via `boxed_smoke`.

use std::path::{Path, PathBuf};

use assert_cmd::Command;

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

fn golden(name: &str) -> String {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/golden")
        .join(name);
    std::fs::read_to_string(&path)
        .unwrap_or_else(|e| panic!("read golden {}: {e}", path.display()))
}

fn run(args: &[&str]) -> String {
    let out = Command::cargo_bin("wolfxl")
        .expect("wolfxl bin built")
        .args(args)
        .output()
        .expect("run wolfxl");
    assert!(
        out.status.success(),
        "wolfxl {:?} failed: stderr={}",
        args,
        String::from_utf8_lossy(&out.stderr)
    );
    String::from_utf8(out.stdout).expect("utf-8 stdout")
}

#[test]
fn text_export_matches_golden_sample() {
    let path = fixture("sample-financials.xlsx");
    let actual = run(&["peek", path.to_str().unwrap(), "-e", "text"]);
    assert_eq!(actual, golden("sample-financials.text"));
}

#[test]
fn csv_export_matches_golden_sample() {
    let path = fixture("sample-financials.xlsx");
    let actual = run(&["peek", path.to_str().unwrap(), "-e", "csv"]);
    assert_eq!(actual, golden("sample-financials.csv"));
}

#[test]
fn json_export_matches_golden_sample() {
    let path = fixture("sample-financials.xlsx");
    let actual = run(&["peek", path.to_str().unwrap(), "-e", "json"]);
    assert_eq!(actual, golden("sample-financials.json"));
}

#[test]
fn text_export_matches_golden_wide() {
    let path = fixture("wide-table.xlsx");
    let actual = run(&["peek", path.to_str().unwrap(), "-e", "text"]);
    assert_eq!(actual, golden("wide-table.text"));
}

#[test]
fn csv_export_matches_golden_wide() {
    let path = fixture("wide-table.xlsx");
    let actual = run(&["peek", path.to_str().unwrap(), "-e", "csv"]);
    assert_eq!(actual, golden("wide-table.csv"));
}

#[test]
fn json_export_matches_golden_wide() {
    let path = fixture("wide-table.xlsx");
    let actual = run(&["peek", path.to_str().unwrap(), "-e", "json"]);
    assert_eq!(actual, golden("wide-table.json"));
}

#[test]
fn boxed_smoke() {
    let path = fixture("sample-financials.xlsx");
    let out = run(&["peek", path.to_str().unwrap(), "-n", "5"]);
    assert!(out.contains("wolfxl peek"), "missing banner: {out}");
    assert!(out.contains("Sheet: P&L"), "missing sheet header: {out}");
    assert!(out.contains("Available sheets:"), "missing sheet list");
    assert!(
        out.contains("Showing 5 of 20 rows"),
        "missing truncation footer: {out}"
    );
    assert!(out.contains("│"), "missing box-drawing char");
}

#[test]
fn boxed_named_sheet() {
    let path = fixture("sample-financials.xlsx");
    let out = run(&["peek", path.to_str().unwrap(), "-s", "Revenue Breakdown", "-n", "3"]);
    assert!(out.contains("Sheet: Revenue Breakdown"), "wrong sheet: {out}");
    assert!(out.contains("Showing 3 of 12 rows"), "wrong cap: {out}");
}

#[test]
fn boxed_max_width_truncates() {
    let path = fixture("wide-table.xlsx");
    let out = run(&["peek", path.to_str().unwrap(), "-n", "2", "-w", "10"]);
    assert!(out.contains("…"), "expected ellipsis from -w 10: {out}");
}

#[test]
fn unknown_sheet_errors() {
    let path = fixture("sample-financials.xlsx");
    let out = Command::cargo_bin("wolfxl")
        .unwrap()
        .args(["peek", path.to_str().unwrap(), "-s", "Does Not Exist"])
        .output()
        .unwrap();
    assert!(!out.status.success());
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert!(stderr.contains("not found"), "stderr was: {stderr}");
}

#[test]
fn missing_file_errors() {
    let out = Command::cargo_bin("wolfxl")
        .unwrap()
        .args(["peek", "/nonexistent/wolfxl-test.xlsx"])
        .output()
        .unwrap();
    assert!(!out.status.success());
}
