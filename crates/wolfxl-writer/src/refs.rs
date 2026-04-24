//! A1 / R1C1 / column-letter helpers. Wave 1B subagent fills this in.
//!
//! Public surface (stubs until Wave 1B):
//!
//! - [`col_to_letters`] — 1-based column number → A/B/…/Z/AA/…/XFD
//! - [`letters_to_col`] — inverse
//! - [`parse_a1`] — `"A1"` → (1, 1); `"XFD1048576"` → max valid
//! - [`format_a1`] — (1, 1) → `"A1"`
//! - [`parse_range`] — `"A1:D20"` → ((1,1),(20,4))
//! - [`format_range`] — inverse
//! - [`sanitize_sheet_name`] — strips `/\?*[]:`, trims to 31 chars
//! - [`quote_sheet_name_if_needed`] — wraps in `'…'` when the name
//!   contains spaces or reserved chars, and doubles embedded `'`.
//!
//! See the plan's "Sheet name sanitization" and "A1 coverage" notes.

/// The maximum column Excel accepts: `XFD` = 16384.
pub const MAX_COL: u32 = 16_384;

/// The maximum row Excel accepts: 1,048,576.
pub const MAX_ROW: u32 = 1_048_576;

/// Convert a 1-based column number to its A1 letter form.
///
/// `1` → `"A"`, `27` → `"AA"`, `16384` → `"XFD"`.
///
/// **Stub** — returns a placeholder; fill in with the Wave 1B implementation.
pub fn col_to_letters(col: u32) -> String {
    assert!(col >= 1 && col <= MAX_COL, "column out of range: {col}");
    // Placeholder implementation to keep the crate compiling.
    // The Wave 1B subagent replaces this with the real base-26 encoder
    // (note: Excel's letters are "bijective base-26" — A, B, ..., Z, AA,
    // AB — not ordinary base-26 which would need a "zero" digit).
    let mut n = col;
    let mut out = Vec::new();
    while n > 0 {
        n -= 1;
        out.push(b'A' + (n % 26) as u8);
        n /= 26;
    }
    out.reverse();
    String::from_utf8(out).expect("base-26 A..Z is ASCII")
}

/// Convert an A1 letter sequence back to a 1-based column number.
///
/// **Stub** — returns a placeholder; fill in with the Wave 1B implementation.
pub fn letters_to_col(s: &str) -> Option<u32> {
    if s.is_empty() {
        return None;
    }
    let mut n: u32 = 0;
    for ch in s.chars() {
        let v = match ch {
            'A'..='Z' => (ch as u32) - ('A' as u32) + 1,
            'a'..='z' => (ch as u32) - ('a' as u32) + 1,
            _ => return None,
        };
        n = n.checked_mul(26)?.checked_add(v)?;
    }
    if n > MAX_COL {
        None
    } else {
        Some(n)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn col_letters_round_trip() {
        for &col in &[1u32, 2, 26, 27, 52, 702, 703, MAX_COL] {
            let letters = col_to_letters(col);
            assert_eq!(letters_to_col(&letters), Some(col), "col={col}, letters={letters}");
        }
    }

    #[test]
    fn well_known_anchors() {
        assert_eq!(col_to_letters(1), "A");
        assert_eq!(col_to_letters(26), "Z");
        assert_eq!(col_to_letters(27), "AA");
        assert_eq!(col_to_letters(702), "ZZ");
        assert_eq!(col_to_letters(703), "AAA");
        assert_eq!(col_to_letters(MAX_COL), "XFD");
    }
}
