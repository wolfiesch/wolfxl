use pyo3::exceptions::PyIOError;
use pyo3::prelude::*;

use std::collections::HashMap;
use std::fs::File;
use std::io::Read;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use zip::ZipArchive;

pub fn normalize_zip_path(path: &str) -> String {
    let mut stack: Vec<&str> = Vec::new();
    for part in path.split('/') {
        if part.is_empty() || part == "." {
            continue;
        }
        if part == ".." {
            stack.pop();
            continue;
        }
        stack.push(part);
    }
    stack.join("/")
}

pub fn join_and_normalize(base_dir: &str, target: &str) -> String {
    let t = target.trim_start_matches('/');
    let combined = if t.starts_with("xl/") {
        t.to_string()
    } else {
        format!("{base_dir}{t}")
    };
    normalize_zip_path(&combined)
}

pub fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for a in e.attributes().with_checks(false).flatten() {
        if a.key.as_ref() == key {
            if let Ok(v) = a.unescape_value() {
                return Some(v.to_string());
            }
            return Some(String::from_utf8_lossy(a.value.as_ref()).into_owned());
        }
    }
    None
}

pub fn parse_workbook_sheet_rids(xml: &str) -> PyResult<Vec<(String, String)>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf: Vec<u8> = Vec::new();
    let mut out: Vec<(String, String)> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"sheet" {
                    let name = attr_value(&e, b"name");
                    let rid = attr_value(&e, b"r:id");
                    if let (Some(n), Some(r)) = (name, rid) {
                        out.push((n, r));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(PyErr::new::<PyIOError, _>(format!(
                    "Failed to parse workbook.xml: {e}"
                )))
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

pub fn parse_relationship_targets(xml: &str) -> PyResult<HashMap<String, String>> {
    // Body-only swap: the wolfxl-rels crate is the single source of truth
    // for the rels grammar. Signature (and lenient skipping of malformed
    // entries) are preserved so existing patcher call sites need not change.
    let graph = wolfxl_rels::RelsGraph::parse(xml.as_bytes())
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to parse rels: {e}")))?;
    let mut out: HashMap<String, String> = HashMap::with_capacity(graph.len());
    for r in graph.iter() {
        out.insert(r.id.0.clone(), r.target.clone());
    }
    Ok(out)
}

pub fn zip_read_to_string(zip: &mut ZipArchive<File>, name: &str) -> PyResult<String> {
    let mut f = zip
        .by_name(name)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Missing zip entry {name}: {e}")))?;
    let mut out = String::new();
    f.read_to_string(&mut out)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read {name}: {e}")))?;
    Ok(out)
}

pub fn zip_read_to_string_opt(zip: &mut ZipArchive<File>, name: &str) -> PyResult<Option<String>> {
    match zip.by_name(name) {
        Ok(mut f) => {
            let mut out = String::new();
            f.read_to_string(&mut out)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read {name}: {e}")))?;
            Ok(Some(out))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(e) => Err(PyErr::new::<PyIOError, _>(format!(
            "Zip error reading {name}: {e}"
        ))),
    }
}
