//! Small carriers for calamine sheet-record emission.

#[derive(Clone, Copy, Debug)]
pub(crate) struct SheetRecordOptions {
    pub(crate) data_only: bool,
    pub(crate) include_format: bool,
    pub(crate) include_empty: bool,
    pub(crate) include_formula_blanks: bool,
    pub(crate) include_coordinate: bool,
    pub(crate) include_style_id: bool,
    pub(crate) include_extended_format: bool,
    pub(crate) include_cached_formula_value: bool,
}
