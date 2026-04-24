//! `xl/sharedStrings.xml` emitter. Wave 2C.
//!
//! Emits the final SST after all sheets have been streamed so that the
//! string-count attributes reflect what was actually referenced.

use crate::intern::SstBuilder;

pub fn emit(_sst: &SstBuilder) -> Vec<u8> {
    Vec::new()
}
