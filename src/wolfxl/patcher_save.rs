//! Shared save-orchestration state for `XlsxPatcher`.

use std::collections::{HashMap, HashSet};
use std::fs::File;

use zip::ZipArchive;

use wolfxl_merger::SheetBlock;

/// Mutable workspace threaded through the ordered save phases.
///
/// Keeping this state in one struct makes `XlsxPatcher::do_save` easier to
/// split without changing phase order or the public PyO3 surface.
pub(super) struct SaveWorkspace {
    pub(super) file_patches: HashMap<String, Vec<u8>>,
    pub(super) part_id_allocator: wolfxl_rels::PartIdAllocator,
    pub(super) cloned_table_names: HashSet<String>,
    pub(super) local_blocks: HashMap<String, Vec<SheetBlock>>,
}

impl SaveWorkspace {
    pub(super) fn new(
        zip: &mut ZipArchive<File>,
        queued_blocks: &HashMap<String, Vec<SheetBlock>>,
    ) -> Self {
        let names: Vec<String> = (0..zip.len())
            .filter_map(|i| zip.by_index(i).ok().map(|e| e.name().to_string()))
            .collect();

        Self {
            file_patches: HashMap::new(),
            part_id_allocator: wolfxl_rels::PartIdAllocator::from_zip_parts(
                names.iter().map(|s| s.as_str()),
            ),
            cloned_table_names: HashSet::new(),
            local_blocks: queued_blocks.clone(),
        }
    }
}
