use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

type PyObject = Py<PyAny>;

mod calamine_xlsb_xls_backend;
mod native_reader_backend;
mod native_writer_anchors;
mod native_writer_autofilter;
mod native_writer_backend;
mod native_writer_cells;
mod native_writer_charts;
mod native_writer_formats;
mod native_writer_images;
mod native_writer_rich_text;
mod native_writer_sheet_features;
mod native_writer_sheet_state;
mod native_writer_workbook;
mod native_writer_workbook_metadata;
mod ooxml_util;
mod streaming;
mod util;
mod wolfxl;
mod wolfxl_core_bridge;

#[pyfunction]
fn build_info(py: Python<'_>) -> PyResult<PyObject> {
    let info = PyDict::new(py);
    info.set_item("package", "wolfxl")?;
    info.set_item("package_version", env!("CARGO_PKG_VERSION"))?;

    let enabled = PyList::new(py, ["native-xlsx", "calamine-binary", "wolfxl"])?;
    info.set_item("enabled_backends", enabled)?;

    let versions = PyDict::new(py);
    versions.set_item(
        "calamine-binary",
        option_env!("WOLFXL_DEP_CALAMINE_VERSION"),
    )?;
    info.set_item("backend_versions", versions)?;

    Ok(info.into())
}

#[pymodule]
fn _rust(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    m.add_function(wrap_pyfunction!(build_info, m)?)?;
    m.add_class::<native_reader_backend::NativeXlsxBook>()?;
    m.add_class::<calamine_xlsb_xls_backend::CalamineXlsbBook>()?;
    m.add_class::<calamine_xlsb_xls_backend::CalamineXlsBook>()?;
    m.add_function(wrap_pyfunction!(
        calamine_xlsb_xls_backend::classify_file_format,
        m
    )?)?;
    m.add_class::<native_writer_backend::NativeWorkbook>()?;
    m.add_function(wrap_pyfunction!(
        native_writer_charts::serialize_chart_dict,
        m
    )?)?;
    // Sprint Ο Pod 1D (RFC-058 §10) — workbook security serializer.
    m.add_function(wrap_pyfunction!(
        native_writer_workbook_metadata::serialize_workbook_security_dict,
        m
    )?)?;
    // Sprint Ν Pod-γ (RFC-047 / RFC-048) — pivot serialisers.
    m.add_function(wrap_pyfunction!(
        wolfxl::pivot::serialize_pivot_cache_dict,
        m
    )?)?;
    m.add_function(wrap_pyfunction!(
        wolfxl::pivot::serialize_pivot_records_dict,
        m
    )?)?;
    m.add_function(wrap_pyfunction!(
        wolfxl::pivot::serialize_pivot_table_dict,
        m
    )?)?;
    // Sprint Ο Pod 1B (RFC-056) — autoFilter serialiser + evaluator.
    m.add_function(wrap_pyfunction!(
        wolfxl::autofilter::serialize_autofilter_dict,
        m
    )?)?;
    m.add_function(wrap_pyfunction!(
        wolfxl::autofilter::evaluate_autofilter,
        m
    )?)?;
    // Sprint Ο Pod 3 (RFC-061) — slicer serialisers.
    m.add_function(wrap_pyfunction!(
        wolfxl::pivot::serialize_slicer_cache_dict,
        m
    )?)?;
    m.add_function(wrap_pyfunction!(wolfxl::pivot::serialize_slicer_dict, m)?)?;
    // Sprint Π Pod Π-α (RFC-062) — page breaks + sheet format serialiser.
    m.add_function(wrap_pyfunction!(
        wolfxl::page_breaks::serialize_page_breaks_dict,
        m
    )?)?;
    m.add_class::<streaming::StreamingSheetReader>()?;
    m.add_class::<wolfxl::XlsxPatcher>()?;
    wolfxl_core_bridge::register(m)?;
    Ok(())
}
