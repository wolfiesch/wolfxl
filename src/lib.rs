use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

type PyObject = Py<PyAny>;

mod calamine_styled_backend;
mod native_writer_backend;
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

    let enabled = PyList::new(py, ["calamine-styles", "wolfxl", "native"])?;
    info.set_item("enabled_backends", enabled)?;

    let versions = PyDict::new(py);
    versions.set_item(
        "calamine-styles",
        option_env!("WOLFXL_DEP_CALAMINE_VERSION"),
    )?;
    info.set_item("backend_versions", versions)?;

    Ok(info.into())
}

#[pymodule]
fn _rust(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    m.add_function(wrap_pyfunction!(build_info, m)?)?;
    m.add_class::<calamine_styled_backend::CalamineStyledBook>()?;
    m.add_class::<native_writer_backend::NativeWorkbook>()?;
    m.add_class::<streaming::StreamingSheetReader>()?;
    m.add_class::<wolfxl::XlsxPatcher>()?;
    wolfxl_core_bridge::register(m)?;
    Ok(())
}
