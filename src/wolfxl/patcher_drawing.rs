//! Drawing helpers for patcher image and chart queues.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyDict;

use super::patcher_models::{QueuedChartAdd, QueuedImageAdd, QueuedImageAnchor};

pub(crate) fn parse_queued_image_anchor(d: &Bound<'_, PyDict>) -> PyResult<QueuedImageAnchor> {
    let kind: String = d
        .get_item("type")?
        .ok_or_else(|| PyValueError::new_err("anchor dict missing 'type'"))?
        .extract()?;
    let q_int = |key: &str, default: u32| -> PyResult<u32> {
        Ok(d.get_item(key)?
            .and_then(|v| v.extract().ok())
            .unwrap_or(default))
    };
    let q_i64 = |key: &str, default: i64| -> PyResult<i64> {
        Ok(d.get_item(key)?
            .and_then(|v| v.extract().ok())
            .unwrap_or(default))
    };
    match kind.as_str() {
        "one_cell" => Ok(QueuedImageAnchor::OneCell {
            from_col: q_int("from_col", 0)?,
            from_row: q_int("from_row", 0)?,
            from_col_off: q_i64("from_col_off", 0)?,
            from_row_off: q_i64("from_row_off", 0)?,
        }),
        "two_cell" => Ok(QueuedImageAnchor::TwoCell {
            from_col: q_int("from_col", 0)?,
            from_row: q_int("from_row", 0)?,
            from_col_off: q_i64("from_col_off", 0)?,
            from_row_off: q_i64("from_row_off", 0)?,
            to_col: q_int("to_col", 0)?,
            to_row: q_int("to_row", 0)?,
            to_col_off: q_i64("to_col_off", 0)?,
            to_row_off: q_i64("to_row_off", 0)?,
            edit_as: d
                .get_item("edit_as")?
                .and_then(|v| v.extract().ok())
                .unwrap_or_else(|| "oneCell".to_string()),
        }),
        "absolute" => Ok(QueuedImageAnchor::Absolute {
            x_emu: q_i64("x_emu", 0)?,
            y_emu: q_i64("y_emu", 0)?,
            cx_emu: q_i64("cx_emu", 0)?,
            cy_emu: q_i64("cy_emu", 0)?,
        }),
        other => Err(PyValueError::new_err(format!(
            "unknown anchor type: {other:?}"
        ))),
    }
}

/// Build the `xl/drawings/drawingN.xml` body for the queued images.
pub(crate) fn build_drawing_xml(images: &[QueuedImageAdd]) -> String {
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let emu_per_px: i64 = 9525;
    let mut out = String::with_capacity(512 + images.len() * 512);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    out.push_str(&format!(
        "<xdr:wsDr xmlns:xdr=\"{xdr_ns}\" xmlns:a=\"{a_ns}\" xmlns:r=\"{r_ns}\">"
    ));
    for (i, img) in images.iter().enumerate() {
        let pic_id = (i + 1) as u32;
        let rid = format!("rId{}", i + 1);
        match &img.anchor {
            QueuedImageAnchor::OneCell {
                from_col,
                from_row,
                from_col_off,
                from_row_off,
            } => {
                out.push_str("<xdr:oneCellAnchor>");
                out.push_str(&format!(
                    "<xdr:from><xdr:col>{from_col}</xdr:col><xdr:colOff>{from_col_off}</xdr:colOff>\
                     <xdr:row>{from_row}</xdr:row><xdr:rowOff>{from_row_off}</xdr:rowOff></xdr:from>"
                ));
                let cx = img.width_px as i64 * emu_per_px;
                let cy = img.height_px as i64 * emu_per_px;
                out.push_str(&format!("<xdr:ext cx=\"{cx}\" cy=\"{cy}\"/>"));
            }
            QueuedImageAnchor::TwoCell {
                from_col,
                from_row,
                from_col_off,
                from_row_off,
                to_col,
                to_row,
                to_col_off,
                to_row_off,
                edit_as,
            } => {
                out.push_str(&format!("<xdr:twoCellAnchor editAs=\"{edit_as}\">"));
                out.push_str(&format!(
                    "<xdr:from><xdr:col>{from_col}</xdr:col><xdr:colOff>{from_col_off}</xdr:colOff>\
                     <xdr:row>{from_row}</xdr:row><xdr:rowOff>{from_row_off}</xdr:rowOff></xdr:from>"
                ));
                out.push_str(&format!(
                    "<xdr:to><xdr:col>{to_col}</xdr:col><xdr:colOff>{to_col_off}</xdr:colOff>\
                     <xdr:row>{to_row}</xdr:row><xdr:rowOff>{to_row_off}</xdr:rowOff></xdr:to>"
                ));
            }
            QueuedImageAnchor::Absolute {
                x_emu,
                y_emu,
                cx_emu,
                cy_emu,
            } => {
                out.push_str("<xdr:absoluteAnchor>");
                out.push_str(&format!("<xdr:pos x=\"{x_emu}\" y=\"{y_emu}\"/>"));
                out.push_str(&format!("<xdr:ext cx=\"{cx_emu}\" cy=\"{cy_emu}\"/>"));
            }
        }
        let cx = img.width_px as i64 * emu_per_px;
        let cy = img.height_px as i64 * emu_per_px;
        out.push_str(&format!(
            "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"{pic_id}\" name=\"Picture {pic_id}\" descr=\"Picture {pic_id}\"/>\
             <xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>\
             <xdr:blipFill><a:blip xmlns:r=\"{r_ns}\" r:embed=\"{rid}\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>\
             <xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>\
             <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>"
        ));
        out.push_str("<xdr:clientData/>");
        match &img.anchor {
            QueuedImageAnchor::OneCell { .. } => out.push_str("</xdr:oneCellAnchor>"),
            QueuedImageAnchor::TwoCell { .. } => out.push_str("</xdr:twoCellAnchor>"),
            QueuedImageAnchor::Absolute { .. } => out.push_str("</xdr:absoluteAnchor>"),
        }
    }
    out.push_str("</xdr:wsDr>");
    out
}

/// Build `xl/drawings/_rels/drawingN.xml.rels` for the given images.
pub(crate) fn build_drawing_rels_xml(images: &[QueuedImageAdd], image_indices: &[u32]) -> String {
    debug_assert_eq!(images.len(), image_indices.len());
    let mut g = wolfxl_rels::RelsGraph::new();
    for (img, &n) in images.iter().zip(image_indices.iter()) {
        g.add(
            wolfxl_rels::rt::IMAGE,
            &format!("../media/image{n}.{}", img.ext),
            wolfxl_rels::TargetMode::Internal,
        );
    }
    String::from_utf8(g.serialize()).expect("rels serialize is utf8")
}

/// Splice a `<drawing r:id="rIdN"/>` element into a sheet XML body.
pub(crate) fn splice_drawing_ref(sheet_xml: &str, rid: &str) -> Result<String, &'static str> {
    if sheet_xml.contains("<drawing ") || sheet_xml.contains("<drawing/>") {
        return Err("sheet already has a <drawing> element");
    }
    let elem = format!("<drawing r:id=\"{rid}\"/>");
    let with_drawing = if let Some(idx) = sheet_xml.find("<legacyDrawing") {
        let mut out = String::with_capacity(sheet_xml.len() + elem.len());
        out.push_str(&sheet_xml[..idx]);
        out.push_str(&elem);
        out.push_str(&sheet_xml[idx..]);
        out
    } else if let Some(idx) = sheet_xml.rfind("</worksheet>") {
        let mut out = String::with_capacity(sheet_xml.len() + elem.len());
        out.push_str(&sheet_xml[..idx]);
        out.push_str(&elem);
        out.push_str(&sheet_xml[idx..]);
        out
    } else {
        return Err("sheet xml has no </worksheet> closing tag");
    };
    Ok(ensure_xmlns_r_on_worksheet(&with_drawing))
}

/// Ensure the `<worksheet>` root element declares the `r` prefix.
pub(crate) fn ensure_xmlns_r_on_worksheet(sheet_xml: &str) -> String {
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let start = match sheet_xml.find("<worksheet") {
        Some(i) => i,
        None => return sheet_xml.to_string(),
    };
    let end = match sheet_xml[start..].find('>') {
        Some(e) => start + e,
        None => return sheet_xml.to_string(),
    };
    let open_tag = &sheet_xml[start..=end];
    if open_tag.contains("xmlns:r=") {
        return sheet_xml.to_string();
    }
    let inserted = if open_tag.ends_with("/>") {
        format!("{} xmlns:r=\"{r_ns}\"/>", &open_tag[..open_tag.len() - 2])
    } else {
        format!("{} xmlns:r=\"{r_ns}\">", &open_tag[..open_tag.len() - 1])
    };
    let mut out = String::with_capacity(sheet_xml.len() + 80);
    out.push_str(&sheet_xml[..start]);
    out.push_str(&inserted);
    out.push_str(&sheet_xml[end + 1..]);
    out
}

/// Parse an A1-style coordinate into zero-based `(col, row)`.
pub(crate) fn parse_a1_coord(s: &str) -> Option<(u32, u32)> {
    let s = s.trim().trim_start_matches('$');
    let mut col: u32 = 0;
    let mut iter = s.chars().peekable();
    let mut col_chars = 0;
    while let Some(&c) = iter.peek() {
        if c.is_ascii_alphabetic() {
            let v = (c.to_ascii_uppercase() as u32) - ('A' as u32) + 1;
            col = col * 26 + v;
            iter.next();
            col_chars += 1;
        } else {
            break;
        }
    }
    if col_chars == 0 || col == 0 {
        return None;
    }
    let rest: String = iter.collect();
    let rest = rest.trim_start_matches('$');
    let row: u32 = rest.parse().ok()?;
    if row == 0 {
        return None;
    }
    Some((col - 1, row - 1))
}

/// Resolve a relative or absolute OOXML rel target against a base directory.
pub(crate) fn resolve_relative_path(base_dir: &str, target: &str) -> String {
    let (mut parts, target_iter): (Vec<&str>, _) = if let Some(stripped) = target.strip_prefix('/')
    {
        (Vec::new(), stripped.split('/'))
    } else {
        (
            base_dir.split('/').filter(|p| !p.is_empty()).collect(),
            target.split('/'),
        )
    };
    for seg in target_iter {
        match seg {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            other => parts.push(other),
        }
    }
    parts.join("/")
}

/// Best-effort extract `N` from `xl/drawings/drawingN.xml`.
pub(crate) fn drawing_n_from_path(path: &str) -> Option<u32> {
    let fname = path.rsplit('/').next()?;
    let core = fname.strip_suffix(".xml")?;
    let n_str = core.strip_prefix("drawing")?;
    n_str.parse::<u32>().ok()
}

/// Build a fresh `xl/drawings/drawingN.xml` body for queued charts.
pub(crate) fn build_chart_drawing_xml(queued: &[QueuedChartAdd], chart_rids: &[String]) -> String {
    debug_assert_eq!(queued.len(), chart_rids.len());
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let c_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    let mut out = String::with_capacity(512 + queued.len() * 768);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    out.push_str(&format!(
        "<xdr:wsDr xmlns:xdr=\"{xdr_ns}\" xmlns:a=\"{a_ns}\" xmlns:r=\"{r_ns}\" xmlns:c=\"{c_ns}\">"
    ));
    for (i, (chart, rid)) in queued.iter().zip(chart_rids.iter()).enumerate() {
        out.push_str(&render_graphic_frame_anchor(chart, rid, (i + 1) as u32));
    }
    out.push_str("</xdr:wsDr>");
    out
}

/// Append one anchor per queued chart to an existing drawing XML body.
pub(crate) fn append_graphic_frames(
    drawing_xml: &[u8],
    queued: &[QueuedChartAdd],
    chart_rids: &[String],
) -> Result<Vec<u8>, String> {
    debug_assert_eq!(queued.len(), chart_rids.len());
    let body = std::str::from_utf8(drawing_xml).map_err(|e| e.to_string())?;
    let use_xdr_prefix = body.contains("<xdr:wsDr") || body.contains("xmlns:xdr=");
    let existing_count: u32 =
        (body.matches("<graphicFrame").count() + body.matches("<pic").count()) as u32;
    let mut new_anchors = String::with_capacity(queued.len() * 512);
    for (i, (chart, rid)) in queued.iter().zip(chart_rids.iter()).enumerate() {
        new_anchors.push_str(&render_graphic_frame_anchor_styled(
            chart,
            rid,
            existing_count + (i + 1) as u32,
            use_xdr_prefix,
        ));
    }
    let pos_opt = body.rfind("</xdr:wsDr>").or_else(|| body.rfind("</wsDr>"));
    if let Some(pos) = pos_opt {
        let mut out = String::with_capacity(body.len() + new_anchors.len());
        out.push_str(&body[..pos]);
        out.push_str(&new_anchors);
        out.push_str(&body[pos..]);
        Ok(out.into_bytes())
    } else {
        let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
        let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let c_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        let mut out = String::with_capacity(new_anchors.len() + 256);
        out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
        out.push_str(&format!(
            "<xdr:wsDr xmlns:xdr=\"{xdr_ns}\" xmlns:a=\"{a_ns}\" xmlns:r=\"{r_ns}\" xmlns:c=\"{c_ns}\">"
        ));
        out.push_str(&new_anchors);
        out.push_str("</xdr:wsDr>");
        Ok(out.into_bytes())
    }
}

fn render_graphic_frame_anchor(chart: &QueuedChartAdd, chart_rid: &str, unique_id: u32) -> String {
    render_graphic_frame_anchor_styled(chart, chart_rid, unique_id, true)
}

fn render_graphic_frame_anchor_styled(
    chart: &QueuedChartAdd,
    chart_rid: &str,
    unique_id: u32,
    use_xdr_prefix: bool,
) -> String {
    let (col0, row0) = parse_a1_coord(&chart.anchor_a1).unwrap_or((3, 1));
    let cx = chart.width_emu;
    let cy = chart.height_emu;
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let c_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    let p = if use_xdr_prefix { "xdr:" } else { "" };
    let root_xmlns = if use_xdr_prefix {
        String::new()
    } else {
        format!(" xmlns=\"{xdr_ns}\" xmlns:a=\"{a_ns}\" xmlns:r=\"{r_ns}\" xmlns:c=\"{c_ns}\"")
    };

    let mut out = String::with_capacity(640);
    out.push_str(&format!("<{p}oneCellAnchor{root_xmlns}>"));
    out.push_str(&format!(
        "<{p}from><{p}col>{col0}</{p}col><{p}colOff>0</{p}colOff>\
         <{p}row>{row0}</{p}row><{p}rowOff>0</{p}rowOff></{p}from>"
    ));
    out.push_str(&format!("<{p}ext cx=\"{cx}\" cy=\"{cy}\"/>"));
    out.push_str(&format!(
        "<{p}graphicFrame macro=\"\">\
           <{p}nvGraphicFramePr>\
             <{p}cNvPr id=\"{unique_id}\" name=\"Chart {unique_id}\"/>\
             <{p}cNvGraphicFramePr/>\
           </{p}nvGraphicFramePr>\
           <{p}xfrm>\
             <a:off x=\"0\" y=\"0\" xmlns:a=\"{a_ns}\"/>\
             <a:ext cx=\"{cx}\" cy=\"{cy}\" xmlns:a=\"{a_ns}\"/>\
           </{p}xfrm>\
           <a:graphic xmlns:a=\"{a_ns}\">\
             <a:graphicData uri=\"{c_ns}\">\
               <c:chart xmlns:c=\"{c_ns}\" \
                        xmlns:r=\"{r_ns}\" \
                        r:id=\"{chart_rid}\"/>\
             </a:graphicData>\
           </a:graphic>\
         </{p}graphicFrame>"
    ));
    out.push_str(&format!("<{p}clientData/>"));
    out.push_str(&format!("</{p}oneCellAnchor>"));
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn splice_drawing_before_legacy_drawing() {
        let xml = r#"<?xml version="1.0"?><worksheet><sheetData/><legacyDrawing r:id="rId2"/></worksheet>"#;
        let out = splice_drawing_ref(xml, "rId5").unwrap();
        assert!(out.contains("<drawing r:id=\"rId5\"/><legacyDrawing"));
    }

    #[test]
    fn splice_drawing_before_close_when_no_legacy() {
        let xml = r#"<?xml version="1.0"?><worksheet><sheetData/></worksheet>"#;
        let out = splice_drawing_ref(xml, "rId1").unwrap();
        assert!(out.contains("<drawing r:id=\"rId1\"/></worksheet>"));
    }

    #[test]
    fn splice_drawing_errors_when_already_present() {
        let xml =
            r#"<?xml version="1.0"?><worksheet><sheetData/><drawing r:id="rId7"/></worksheet>"#;
        assert!(splice_drawing_ref(xml, "rId1").is_err());
    }

    #[test]
    fn build_drawing_xml_roundtrip() {
        let imgs = vec![QueuedImageAdd {
            data: vec![],
            ext: "png".into(),
            width_px: 10,
            height_px: 5,
            anchor: QueuedImageAnchor::OneCell {
                from_col: 1,
                from_row: 4,
                from_col_off: 0,
                from_row_off: 0,
            },
        }];
        let xml = build_drawing_xml(&imgs);
        assert!(xml.contains("<xdr:oneCellAnchor>"));
        assert!(xml.contains("r:embed=\"rId1\""));
    }

    #[test]
    fn parse_a1_basic_cells() {
        assert_eq!(parse_a1_coord("A1"), Some((0, 0)));
        assert_eq!(parse_a1_coord("D2"), Some((3, 1)));
        assert_eq!(parse_a1_coord("Z1"), Some((25, 0)));
        assert_eq!(parse_a1_coord("AA1"), Some((26, 0)));
        assert_eq!(parse_a1_coord("$D$2"), Some((3, 1)));
        assert!(parse_a1_coord("").is_none());
        assert!(parse_a1_coord("1A").is_none());
    }

    #[test]
    fn resolve_relative_basic() {
        assert_eq!(
            resolve_relative_path("xl/worksheets", "../drawings/drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
        assert_eq!(
            resolve_relative_path("xl/drawings", "../charts/chart1.xml"),
            "xl/charts/chart1.xml"
        );
    }

    #[test]
    fn drawing_n_extract() {
        assert_eq!(drawing_n_from_path("xl/drawings/drawing7.xml"), Some(7));
        assert_eq!(drawing_n_from_path("xl/drawings/drawing.xml"), None);
        assert_eq!(drawing_n_from_path("nope.xml"), None);
    }

    #[test]
    fn build_drawing_xml_for_one_chart() {
        let q = vec![QueuedChartAdd {
            chart_xml: b"<chartSpace/>".to_vec(),
            anchor_a1: "D2".into(),
            width_emu: 4_572_000,
            height_emu: 2_743_200,
        }];
        let rids = vec!["rId1".to_string()];
        let body = build_chart_drawing_xml(&q, &rids);
        assert!(body.contains("<xdr:graphicFrame"));
        assert!(body.contains("r:id=\"rId1\""));
        assert!(body.contains("<xdr:col>3</xdr:col>"));
        assert!(body.contains("<xdr:row>1</xdr:row>"));
    }

    #[test]
    fn append_graphic_frame_inserts_before_close() {
        let original = b"<?xml version=\"1.0\"?><xdr:wsDr xmlns:xdr=\"x\" xmlns:r=\"r\" xmlns:c=\"c\"><xdr:oneCellAnchor/></xdr:wsDr>";
        let q = vec![QueuedChartAdd {
            chart_xml: vec![],
            anchor_a1: "B5".into(),
            width_emu: 100,
            height_emu: 200,
        }];
        let rids = vec!["rId7".to_string()];
        let merged = append_graphic_frames(original, &q, &rids).unwrap();
        let s = std::str::from_utf8(&merged).unwrap();
        assert!(s.contains("<xdr:oneCellAnchor/>"));
        assert!(s.contains("<xdr:graphicFrame"));
        assert!(s.contains("r:id=\"rId7\""));
        assert!(s.ends_with("</xdr:wsDr>"));
    }
}
