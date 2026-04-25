//! Defined names (named ranges) scoped to the workbook or a specific sheet.

/// A single defined name. Workbook-scope names live on
/// [`crate::model::workbook::Workbook::defined_names`]; sheet-scope names
/// are also stored there with `scope_sheet_index` set to the owning sheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DefinedName {
    /// The name Excel shows in the Name Manager. Rules: starts with a
    /// letter or `_`, no spaces, ≤ 255 chars, not a cell reference.
    pub name: String,

    /// The formula / range reference the name resolves to. Sheet names
    /// inside the formula are quoted on emission if they contain spaces.
    pub formula: String,

    /// `None` for workbook-scope, `Some(idx)` for a specific sheet.
    pub scope_sheet_index: Option<usize>,

    /// If set, the name is a print area or title — Excel uses magic
    /// well-known names like `_xlnm.Print_Area` and `_xlnm.Print_Titles`.
    /// Callers pass the user-facing name; the emitter adds the `_xlnm.`
    /// prefix where appropriate.
    pub builtin: Option<BuiltinName>,

    /// Whether the name is hidden from the Name Manager (but still usable
    /// in formulas). Used by some apps for bookkeeping ranges.
    pub hidden: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BuiltinName {
    PrintArea,
    PrintTitles,
    /// Referenced by Excel as `_xlnm._FilterDatabase`.
    FilterDatabase,
}
