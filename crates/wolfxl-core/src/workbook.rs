use std::fs::File;
use std::io::BufReader;
use std::path::{Path, PathBuf};

use calamine_styles::{Reader, Xlsx};

use crate::error::{Error, Result};
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
        let inner: XlsxReader =
            Xlsx::new(reader).map_err(|e| Error::Xlsx(format!("failed to parse xlsx: {e}")))?;
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
}
