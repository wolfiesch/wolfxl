//! Cell value and wrapper — what goes into a single spreadsheet cell at write time.

/// The value stored in a single spreadsheet cell, pre-serialization.
///
/// This is distinct from [`wolfxl_core::CellValue`] (the read-path type).
/// The read type carries calamine-parsed dates as `NaiveDate`/`NaiveDateTime`;
/// the write type carries only the f64 Excel serial because that's what
/// OOXML actually stores. Conversion from a chrono date to a serial happens
/// via [`crate::model::date::to_excel_serial`] before the value reaches
/// [`WriteCellValue::DateSerial`].
#[derive(Debug, Clone, PartialEq)]
pub enum WriteCellValue {
    /// A cell with no value. Emitted as `<c r="A1" s="…"/>` when styled,
    /// or omitted entirely when unstyled.
    Blank,

    /// IEEE-754 f64. Excel stores ints and floats in the same slot;
    /// distinguishing them at the OOXML layer is not meaningful.
    Number(f64),

    /// A string. Always routed through the shared string table
    /// (see [`crate::intern`]) — inline strings are never emitted.
    String(String),

    /// Boolean. Serialized as `<c t="b"><v>0</v></c>` or `<v>1</v>`.
    Boolean(bool),

    /// Formula with optional pre-computed result.
    ///
    /// `expr` is the formula body without a leading `=`
    /// (i.e. `"SUM(A1:A10)"`, not `"=SUM(A1:A10)"`).
    /// `result` is the cached value Excel shows before recalculation;
    /// if `None`, the emitter writes `<f>…</f>` without a `<v>` sibling.
    Formula {
        expr: String,
        result: Option<FormulaResult>,
    },

    /// Pre-computed Excel date serial (days since 1899-12-30, with the
    /// 1900-leap-year quirk handled — see [`crate::model::date`]).
    ///
    /// Rendered the same as `Number` at the XML layer; the distinction
    /// exists so higher-level code can keep track of which numbers are
    /// dates for style-application decisions.
    DateSerial(f64),
}

/// The cached result of a formula.
#[derive(Debug, Clone, PartialEq)]
pub enum FormulaResult {
    Number(f64),
    String(String),
    Boolean(bool),
}

/// A single cell: its value and an optional pointer into the styles table.
#[derive(Debug, Clone, PartialEq)]
pub struct WriteCell {
    pub value: WriteCellValue,
    /// Index into `xl/styles.xml`'s `<cellXfs>` block, if styled.
    ///
    /// `None` means the cell inherits default style and the `s` attribute
    /// is omitted (not set to `"0"`). This matches what Excel itself writes.
    pub style_id: Option<u32>,
}

impl WriteCell {
    pub fn new(value: WriteCellValue) -> Self {
        Self {
            value,
            style_id: None,
        }
    }

    pub fn with_style(mut self, style_id: u32) -> Self {
        self.style_id = Some(style_id);
        self
    }
}

impl From<WriteCellValue> for WriteCell {
    fn from(value: WriteCellValue) -> Self {
        WriteCell::new(value)
    }
}
