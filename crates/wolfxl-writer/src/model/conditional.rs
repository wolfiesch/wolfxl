//! Conditional formatting data. Wave 3 module — types only here.

/// A single conditional-formatting block applied to a range.
#[derive(Debug, Clone, PartialEq)]
pub struct ConditionalFormat {
    /// The A1 range this block applies to. Can be a single cell or
    /// multi-area (space-separated).
    pub sqref: String,
    pub rules: Vec<ConditionalRule>,
}

/// One rule within a conditional-formatting block. Priority is set by
/// order in the parent `Vec` — earlier rules win.
#[derive(Debug, Clone, PartialEq)]
pub struct ConditionalRule {
    pub kind: ConditionalKind,
    /// Index into `xl/styles.xml`'s `<dxfs>` (differential formats) table.
    /// `None` for rule types like icon sets that don't use dxf styling.
    pub dxf_id: Option<u32>,
    pub stop_if_true: bool,
}

/// The large sum type covering every CF rule Excel recognizes.
///
/// Kept explicit rather than string-based so the emitter can lean on
/// the Rust match-exhaustiveness check when adding new kinds.
#[derive(Debug, Clone, PartialEq)]
pub enum ConditionalKind {
    CellIs {
        operator: CellIsOperator,
        formula_a: String,
        formula_b: Option<String>,
    },
    Expression {
        formula: String,
    },
    ContainsText {
        text: String,
    },
    NotContainsText {
        text: String,
    },
    BeginsWith {
        text: String,
    },
    EndsWith {
        text: String,
    },
    Duplicate,
    Unique,
    Top10 {
        count: u32,
        bottom: bool,
        percent: bool,
    },
    AboveAverage {
        above: bool,
        /// 0 means "exactly average"; 1 = "one std dev", etc.
        std_dev: u32,
    },
    /// Gradient color scale (2-stop or 3-stop).
    ColorScale {
        stops: Vec<ColorScaleStop>,
    },
    DataBar {
        color_rgb: String,
        min: ConditionalThreshold,
        max: ConditionalThreshold,
    },
    IconSet {
        /// e.g. `"3TrafficLights1"`, `"5Arrows"`, `"4Rating"`.
        set_name: String,
        /// One entry per icon band; length must match the set.
        thresholds: Vec<ConditionalThreshold>,
        /// OOXML `showValue` attribute. Default is `true` (matches the
        /// OOXML spec default), so emit `showValue="0"` only when this
        /// is explicitly `false`.
        show_value: bool,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellIsOperator {
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
    Between,
    NotBetween,
}

/// One stop on a color-scale gradient.
#[derive(Debug, Clone, PartialEq)]
pub struct ColorScaleStop {
    pub threshold: ConditionalThreshold,
    /// ARGB like `"FFRRGGBB"`.
    pub color_rgb: String,
}

/// A threshold value for gradient / bar / icon rules.
#[derive(Debug, Clone, PartialEq)]
pub enum ConditionalThreshold {
    Min,
    Max,
    /// Literal number.
    Number(f64),
    /// Percentage 0-100.
    Percent(f64),
    /// Percentile 0-100.
    Percentile(f64),
    /// Formula that evaluates to a threshold.
    Formula(String),
}
