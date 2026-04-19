use std::fs::File;
use std::io::BufReader;
use std::path::{Path, PathBuf};

use calamine_styles::{Reader, Xlsx};

use crate::error::{Error, Result};
use crate::map::{SheetMap, WorkbookMap, classify_sheet};
use crate::sheet::Sheet;

type XlsxReader = Xlsx<BufReader<File>>;

pub struct Workbook {
    inner: XlsxReader,
    sheet_names: Vec<String>,
    path: PathBuf,
}

impl Workbook {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref().to_path_buf();
        let file = File::open(&path)?;
        let reader = BufReader::new(file);
        let mut inner: XlsxReader =
            Xlsx::new(reader).map_err(|e| Error::Xlsx(format!("failed to parse xlsx: {e}")))?;
        // calamine `table_names*` panic if tables haven't been loaded; load
        // eagerly so downstream `&self` accessors stay infallible. The call
        // is idempotent and cheap on workbooks without table parts.
        let _ = inner.load_tables();
        let sheet_names = inner.sheet_names().to_vec();
        Ok(Self {
            inner,
            sheet_names,
            path,
        })
    }

    pub fn path(&self) -> &Path {
        &self.path
    }

    pub fn sheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    /// Load a sheet by name. Reads the entire range eagerly; for huge sheets,
    /// callers should pass a row cap downstream rather than load everything.
    pub fn sheet(&mut self, name: &str) -> Result<Sheet> {
        if !self.sheet_names.iter().any(|n| n == name) {
            return Err(Error::SheetNotFound(name.to_string()));
        }
        Sheet::load(&mut self.inner, name)
    }

    /// Convenience: first sheet in workbook order.
    pub fn first_sheet(&mut self) -> Result<Sheet> {
        let name = self
            .sheet_names
            .first()
            .ok_or_else(|| Error::SheetNotFound("(workbook has no sheets)".to_string()))?
            .clone();
        self.sheet(&name)
    }

    /// Workbook-level defined names as `(name, formula)` pairs, exactly
    /// as calamine surfaces them.
    pub fn named_ranges(&self) -> Vec<(String, String)> {
        self.inner.defined_names().to_vec()
    }

    /// Names of workbook tables anchored on the given sheet. Empty when
    /// the sheet has none, which is the common case.
    pub fn table_names_in_sheet(&self, sheet_name: &str) -> Vec<String> {
        self.inner
            .table_names_in_sheet(sheet_name)
            .into_iter()
            .cloned()
            .collect()
    }

    /// Build a one-page summary: every sheet's dimensions, headers,
    /// classification, and anchored tables, plus workbook-level defined
    /// names. Loads each sheet eagerly to compute density for the
    /// classifier — for huge workbooks the caller bears that IO cost.
    pub fn map(&mut self) -> Result<WorkbookMap> {
        let path = self.path.to_string_lossy().into_owned();
        let named_ranges = self.named_ranges();
        let names = self.sheet_names.clone();
        let mut sheets = Vec::with_capacity(names.len());
        for name in &names {
            let tables = self.table_names_in_sheet(name);
            let sheet = self.sheet(name)?;
            let (rows, cols) = sheet.dimensions();
            sheets.push(SheetMap {
                name: name.clone(),
                rows,
                cols,
                class: classify_sheet(&sheet),
                headers: sheet.headers(),
                tables,
            });
        }
        Ok(WorkbookMap {
            path,
            sheets,
            named_ranges,
        })
    }
}
