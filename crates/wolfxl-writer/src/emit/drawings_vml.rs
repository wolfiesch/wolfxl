//! `xl/drawings/vmlDrawing{N}.vml` emitter — legacy VML anchor shapes for
//! comment boxes. Wave 3A.
//!
//! VML is an ancient Microsoft XML dialect kept around specifically for
//! comment boxes and form controls. Modern OOXML uses DrawingML for
//! everything else, but comments still need VML to show Excel a
//! yellow-rectangle shape anchored to a cell.

use crate::model::worksheet::Worksheet;

pub fn emit(_sheet: &Worksheet) -> Vec<u8> {
    Vec::new()
}
