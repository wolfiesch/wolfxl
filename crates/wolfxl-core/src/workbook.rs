use std::collections::HashMap;
use std::fs::File;
use std::io::BufReader;
use std::path::{Path, PathBuf};

use calamine_styles::{Reader, Xlsx};
use zip::ZipArchive;

use crate::error::{Error, Result};
use crate::map::{classify_sheet, SheetMap, WorkbookMap};
use crate::ooxml::{
    join_and_normalize, parse_relationship_targets, parse_workbook_sheet_rids,
    zip_read_to_string, zip_read_to_string_opt,
};
use crate::sheet::Sheet;
use crate::styles::{parse_cellxfs, parse_num_fmts, XfEntry};
use crate::worksheet_xml::parse_cell_style_ids;

type XlsxReader = Xlsx<BufReader<File>>;

/// Pre-parsed style tables and worksheet-path lookup for a workbook.
///
/// Populated lazily on first sheet load and then shared across sheets.
/// Each sheet's per-cell `(row, col) → styleId` map is populated on demand
/// via [`WorkbookStyles::sheet_style_ids_mut`] to avoid walking every
/// worksheet XML up-front on workbooks where only one sheet is read.
pub struct WorkbookStyles {
    cell_xfs: Vec<XfEntry>,
    num_fmts: HashMap<u32, String>,
    sheet_xml_paths: HashMap<String, String>,
    per_sheet_style_ids: HashMap<String, HashMap<(u32, u32), u32>>,
    zip_path: PathBuf,
}

impl WorkbookStyles {
    fn load(zip_path: &Path) -> Result<Self> {
        let file = File::open(zip_path)?;
        let mut zip = ZipArchive::new(file)
            .map_err(|e| Error::Xlsx(format!("failed to open xlsx zip: {e}")))?;

        let cell_xfs = match zip_read_to_string_opt(&mut zip, "xl/styles.xml")? {
            Some(xml) => parse_cellxfs(&xml),
            None => Vec::new(),
        };
        let num_fmts = match zip_read_to_string_opt(&mut zip, "xl/styles.xml")? {
            Some(xml) => parse_num_fmts(&xml)?,
            None => HashMap::new(),
        };

        let workbook_xml = zip_read_to_string(&mut zip, "xl/workbook.xml")?;
        let rels_xml = zip_read_to_string(&mut zip, "xl/_rels/workbook.xml.rels")?;
        let sheet_rids = parse_workbook_sheet_rids(&workbook_xml)?;
        let rel_targets = parse_relationship_targets(&rels_xml)?;
        let mut sheet_xml_paths: HashMap<String, String> = HashMap::new();
        for (name, rid) in sheet_rids {
            if let Some(target) = rel_targets.get(&rid) {
                sheet_xml_paths.insert(name, join_and_normalize("xl/", target));
            }
        }

        Ok(Self {
            cell_xfs,
            num_fmts,
            sheet_xml_paths,
            per_sheet_style_ids: HashMap::new(),
            zip_path: zip_path.to_path_buf(),
        })
    }

    /// Resolve a styleId to a number-format string, consulting custom
    /// numFmts first and the Excel built-in table second. Returns `None`
    /// for styleId 0 (default), unknown IDs, or when the resolved code is
    /// the no-op `"General"`.
    pub fn number_format_for_style_id(&self, style_id: u32) -> Option<&str> {
        if style_id == 0 {
            return None;
        }
        let xf = self.cell_xfs.get(style_id as usize)?;
        let code = crate::styles::resolve_num_fmt(xf.num_fmt_id, &self.num_fmts)?;
        if code.trim().is_empty() || code.eq_ignore_ascii_case("General") {
            None
        } else {
            Some(code)
        }
    }

    /// Read-only access to the per-cell styleId map for a sheet. Returns
    /// `None` until [`WorkbookStyles::sheet_style_ids_mut`] has populated
    /// it. Used on the per-cell fast path where mutation would require
    /// exclusive access.
    pub fn sheet_style_ids(&self, sheet_name: &str) -> Option<&HashMap<(u32, u32), u32>> {
        self.per_sheet_style_ids.get(sheet_name)
    }

    /// Lazily populate the per-cell styleId map for a sheet. Returns a
    /// reference to the cached map. Reading the XML is the expensive part;
    /// `&mut self` makes caching explicit.
    pub fn sheet_style_ids_mut(
        &mut self,
        sheet_name: &str,
    ) -> Result<&HashMap<(u32, u32), u32>> {
        if !self.per_sheet_style_ids.contains_key(sheet_name) {
            let Some(path) = self.sheet_xml_paths.get(sheet_name).cloned() else {
                self.per_sheet_style_ids
                    .insert(sheet_name.to_string(), HashMap::new());
                return Ok(self.per_sheet_style_ids.get(sheet_name).unwrap());
            };
            let file = File::open(&self.zip_path)?;
            let mut zip = ZipArchive::new(file)
                .map_err(|e| Error::Xlsx(format!("failed to open xlsx zip: {e}")))?;
            let map = match zip_read_to_string_opt(&mut zip, &path)? {
                Some(xml) => parse_cell_style_ids(&xml)?,
                None => HashMap::new(),
            };
            self.per_sheet_style_ids.insert(sheet_name.to_string(), map);
        }
        Ok(self.per_sheet_style_ids.get(sheet_name).unwrap())
    }

    /// Test-only access to the parsed cellXfs table.
    #[cfg(test)]
    pub fn cell_xfs(&self) -> &[XfEntry] {
        &self.cell_xfs
    }
}

pub struct Workbook {
    inner: XlsxReader,
    sheet_names: Vec<String>,
    path: PathBuf,
    styles: Option<WorkbookStyles>,
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
            styles: None,
        })
    }

    pub fn path(&self) -> &Path {
        &self.path
    }

    pub fn sheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    /// Lazy accessor for the pre-parsed styles bundle. First call walks
    /// `xl/styles.xml` + `xl/workbook.xml` + rels; subsequent calls reuse
    /// the cached [`WorkbookStyles`]. Returns `None` only if the zip is
    /// unreadable (propagated via `Result`).
    pub fn styles(&mut self) -> Result<&mut WorkbookStyles> {
        if self.styles.is_none() {
            self.styles = Some(WorkbookStyles::load(&self.path)?);
        }
        Ok(self.styles.as_mut().unwrap())
    }

    /// Load a sheet by name. Reads the entire range eagerly; for huge sheets,
    /// callers should pass a row cap downstream rather than load everything.
    pub fn sheet(&mut self, name: &str) -> Result<Sheet> {
        if !self.sheet_names.iter().any(|n| n == name) {
            return Err(Error::SheetNotFound(name.to_string()));
        }
        // Attempt styles load lazily. Failure here is non-fatal: older or
        // non-standard xlsx zips without the expected parts still yield
        // value-only sheets via the calamine path.
        if self.styles.is_none() {
            self.styles = WorkbookStyles::load(&self.path).ok();
        }
        // Split borrow: `inner` and `styles` are separate fields.
        Sheet::load(&mut self.inner, name, self.styles.as_mut())
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
