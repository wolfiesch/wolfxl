//! `xl/comments/comments{N}.xml` emitter — multi-author comments with
//! insertion-ordered authors (fixes the rust_xlsxwriter BTreeMap bug).
//! Wave 3A.

use crate::model::comment::CommentAuthorTable;
use crate::model::worksheet::Worksheet;

pub fn emit(_sheet: &Worksheet, _authors: &CommentAuthorTable) -> Vec<u8> {
    Vec::new()
}
