//! Python dictionary payload parsers for patcher queues.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use super::conditional_formatting::{CfRuleKind, CfRulePatch, CfvoPatch, ColorScaleStop, DxfPatch};
use super::styles;
use super::styles::FormatSpec;

pub(crate) fn dict_to_format_spec(d: &Bound<'_, PyDict>) -> PyResult<FormatSpec> {
    let mut spec = FormatSpec::default();

    let bold = extract_bool(d, "bold")?;
    let italic = extract_bool(d, "italic")?;
    let underline = extract_bool(d, "underline")?;
    let strikethrough = extract_bool(d, "strikethrough")?;
    let font_name = extract_str(d, "font_name")?;
    let font_size = extract_u32(d, "font_size")?;
    let font_color = extract_str(d, "font_color")?;

    if bold.is_some()
        || italic.is_some()
        || underline.is_some()
        || strikethrough.is_some()
        || font_name.is_some()
        || font_size.is_some()
        || font_color.is_some()
    {
        spec.font = Some(styles::FontSpec {
            bold: bold.unwrap_or(false),
            italic: italic.unwrap_or(false),
            underline: underline.unwrap_or(false),
            strikethrough: strikethrough.unwrap_or(false),
            name: font_name,
            size: font_size,
            color_rgb: font_color.map(|c| normalize_color(&c)),
        });
    }

    if let Some(color) = extract_str(d, "bg_color")? {
        spec.fill = Some(styles::FillSpec {
            pattern_type: "solid".to_string(),
            fg_color_rgb: Some(normalize_color(&color)),
        });
    }

    spec.number_format = extract_str(d, "number_format")?;

    let horizontal = extract_str(d, "horizontal")?.or(extract_str(d, "h_align")?);
    let vertical = extract_str(d, "vertical")?.or(extract_str(d, "v_align")?);
    let wrap_text = extract_bool(d, "wrap_text")?.or(extract_bool(d, "wrap")?);
    let indent = extract_u32(d, "indent")?;
    let text_rotation = extract_u32(d, "text_rotation")?.or(extract_u32(d, "rotation")?);

    if horizontal.is_some()
        || vertical.is_some()
        || wrap_text.is_some()
        || indent.is_some()
        || text_rotation.is_some()
    {
        spec.alignment = Some(styles::AlignmentSpec {
            horizontal,
            vertical,
            wrap_text: wrap_text.unwrap_or(false),
            indent: indent.unwrap_or(0),
            text_rotation: text_rotation.unwrap_or(0),
        });
    }

    Ok(spec)
}

pub(crate) fn dict_to_border_spec(d: &Bound<'_, PyDict>) -> PyResult<styles::BorderSpec> {
    fn extract_side(d: &Bound<'_, PyDict>, key: &str) -> PyResult<styles::BorderSideSpec> {
        if let Some(side) = d.get_item(key)? {
            if let Ok(sd) = side.cast::<PyDict>() {
                let style = extract_str(sd, "style")?;
                let color = extract_str(sd, "color")?.map(|c| normalize_color(&c));
                return Ok(styles::BorderSideSpec {
                    style,
                    color_rgb: color,
                });
            }
        }
        Ok(styles::BorderSideSpec::default())
    }

    Ok(styles::BorderSpec {
        left: extract_side(d, "left")?,
        right: extract_side(d, "right")?,
        top: extract_side(d, "top")?,
        bottom: extract_side(d, "bottom")?,
    })
}

pub(crate) fn extract_cf_rule(d: &Bound<'_, PyDict>) -> PyResult<CfRulePatch> {
    let kind_tag =
        extract_str(d, "kind")?.ok_or_else(|| PyValueError::new_err("CF rule requires 'kind'"))?;

    let kind = match kind_tag.as_str() {
        "cellIs" => CfRuleKind::CellIs {
            operator: extract_str(d, "operator")?.unwrap_or_else(|| "equal".to_string()),
            formula_a: extract_str(d, "formula_a")?.unwrap_or_default(),
            formula_b: extract_str(d, "formula_b")?,
        },
        "expression" => CfRuleKind::Expression {
            formula: extract_str(d, "formula")?.unwrap_or_default(),
        },
        "colorScale" => {
            let stops_obj = d
                .get_item("stops")?
                .ok_or_else(|| PyValueError::new_err("colorScale rule requires 'stops'"))?;
            let stops_list = stops_obj
                .cast::<PyList>()
                .map_err(|_| PyValueError::new_err("'stops' must be a list of dicts"))?;
            let mut stops: Vec<ColorScaleStop> = Vec::with_capacity(stops_list.len());
            for s in stops_list.iter() {
                let sd = s
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("each stop must be a dict"))?;
                stops.push(ColorScaleStop {
                    cfvo: CfvoPatch {
                        cfvo_type: extract_str(sd, "cfvo_type")?
                            .unwrap_or_else(|| "min".to_string()),
                        val: extract_str(sd, "val")?,
                    },
                    color_rgb: extract_str(sd, "color_rgb")?.unwrap_or_default(),
                });
            }
            CfRuleKind::ColorScale { stops }
        }
        "dataBar" => CfRuleKind::DataBar {
            min: CfvoPatch {
                cfvo_type: extract_str(d, "min_cfvo_type")?.unwrap_or_else(|| "min".to_string()),
                val: extract_str(d, "min_val")?,
            },
            max: CfvoPatch {
                cfvo_type: extract_str(d, "max_cfvo_type")?.unwrap_or_else(|| "max".to_string()),
                val: extract_str(d, "max_val")?,
            },
            color_rgb: extract_str(d, "color_rgb")?.unwrap_or_default(),
        },
        other => {
            return Err(PyValueError::new_err(format!(
                "unsupported CF rule kind: '{other}'"
            )));
        }
    };

    let dxf = match d.get_item("dxf")? {
        Some(v) if !v.is_none() => {
            let dd = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("'dxf' must be a dict or None"))?;
            Some(extract_dxf_patch(dd)?)
        }
        _ => None,
    };

    Ok(CfRulePatch {
        kind,
        dxf,
        stop_if_true: extract_bool(d, "stop_if_true")?.unwrap_or(false),
    })
}

fn extract_dxf_patch(d: &Bound<'_, PyDict>) -> PyResult<DxfPatch> {
    Ok(DxfPatch {
        font_bold: extract_bool(d, "font_bold")?,
        font_italic: extract_bool(d, "font_italic")?,
        font_color_rgb: extract_str(d, "font_color_rgb")?.map(|c| normalize_color(&c)),
        fill_pattern_type: extract_str(d, "fill_pattern_type")?,
        fill_fg_color_rgb: extract_str(d, "fill_fg_color_rgb")?.map(|c| normalize_color(&c)),
        border_top_style: extract_str(d, "border_top_style")?,
        border_bottom_style: extract_str(d, "border_bottom_style")?,
        border_left_style: extract_str(d, "border_left_style")?,
        border_right_style: extract_str(d, "border_right_style")?,
    })
}

pub(crate) fn extract_str(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<String>> {
    d.get_item(key)?.map(|v| v.extract::<String>()).transpose()
}

pub(crate) fn extract_bool(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<bool>> {
    d.get_item(key)?.map(|v| v.extract::<bool>()).transpose()
}

pub(crate) fn extract_u32(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<u32>> {
    d.get_item(key)?.map(|v| v.extract::<u32>()).transpose()
}

pub(crate) fn extract_f64(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<f64>> {
    d.get_item(key)?.map(|v| v.extract::<f64>()).transpose()
}

fn normalize_color(color: &str) -> String {
    let hex = color.trim_start_matches('#');
    if hex.len() == 6 {
        format!("FF{}", hex.to_uppercase())
    } else if hex.len() == 8 {
        hex.to_uppercase()
    } else {
        format!("FF{hex}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normalize_color_adds_alpha() {
        assert_eq!(normalize_color("#c00000"), "FFC00000");
        assert_eq!(normalize_color("FFC00000"), "FFC00000");
    }
}
