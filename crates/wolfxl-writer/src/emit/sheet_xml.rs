//! `xl/worksheets/sheet{N}.xml` emitter — rows, cells, merges, freeze,
//! columns, print area, and extension hooks for CF/DV. Wave 2B.

use crate::intern::SstBuilder;
use crate::model::format::StylesBuilder;
use crate::model::worksheet::Worksheet;

pub fn emit(
    _sheet: &Worksheet,
    _sst: &mut SstBuilder,
    _styles: &mut StylesBuilder,
) -> Vec<u8> {
    Vec::new()
}
