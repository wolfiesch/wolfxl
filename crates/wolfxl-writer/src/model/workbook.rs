//! Top-level workbook data: sheets, doc props, styles, SST, defined names.

use super::comment::CommentAuthorTable;
use super::defined_name::DefinedName;
use super::format::StylesBuilder;
use super::worksheet::Worksheet;
use crate::intern::SstBuilder;

/// A workbook awaiting serialization.
///
/// Build one up by calling [`Workbook::add_sheet`] (and the various property
/// setters), then hand it to the emitter (Wave 4 — the emitter entrypoint
/// arrives with the `NativeWorkbook` pyclass).
#[derive(Debug, Default)]
pub struct Workbook {
    /// Sheets in display order. The first sheet is the default visible tab.
    pub sheets: Vec<Worksheet>,

    /// Workbook-level defined names (formula aliases, print areas at
    /// the workbook scope).
    pub defined_names: Vec<DefinedName>,

    /// Document properties surfaced in `docProps/core.xml` and `app.xml`.
    pub doc_props: DocProperties,

    /// The shared styles table. Styles are deduped on insertion.
    pub styles: StylesBuilder,

    /// The shared string table. Strings are interned on insertion.
    pub sst: SstBuilder,

    /// Workbook-scope comment authors. Every `commentsN.xml` emits this
    /// full `<authors>` block — `authorId` on a `<comment>` indexes into
    /// it. Insertion order is preserved by [`CommentAuthorTable`] so
    /// multi-author workbooks round-trip without the BTreeMap reordering
    /// bug that motivated this rewrite.
    pub comment_authors: CommentAuthorTable,
}

impl Workbook {
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a sheet and return its index.
    pub fn add_sheet(&mut self, sheet: Worksheet) -> usize {
        let idx = self.sheets.len();
        self.sheets.push(sheet);
        idx
    }

    pub fn sheet_mut(&mut self, idx: usize) -> Option<&mut Worksheet> {
        self.sheets.get_mut(idx)
    }

    pub fn sheet_by_name(&self, name: &str) -> Option<&Worksheet> {
        self.sheets.iter().find(|s| s.name == name)
    }
}

/// Document properties surfaced in the two docProps parts.
///
/// Excel shows these in File → Info. They're optional — if everything
/// is `None`, the emitter writes minimal stub parts so the container
/// stays valid.
#[derive(Debug, Clone, Default)]
pub struct DocProperties {
    pub title: Option<String>,
    pub subject: Option<String>,
    pub creator: Option<String>,
    pub keywords: Option<String>,
    pub description: Option<String>,
    pub last_modified_by: Option<String>,
    pub category: Option<String>,

    /// If `None`, the emitter uses the current wall-clock time (or the
    /// `WOLFXL_TEST_EPOCH` override for deterministic output).
    pub created: Option<chrono::NaiveDateTime>,
    pub modified: Option<chrono::NaiveDateTime>,

    /// Company / application metadata shown in `docProps/app.xml`.
    pub company: Option<String>,
    pub manager: Option<String>,
}
