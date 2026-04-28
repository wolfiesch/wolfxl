//! `xl/styles.xml` emitter — full styles file assembled from [`StylesBuilder`]. Wave 2A.

use crate::model::format::{
    AlignmentSpec, BorderSideSpec, BorderSpec, DxfRecord, FillSpec, FontSpec, StylesBuilder,
    XfRecord,
};
use crate::xml_escape;

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

/// Emit the complete `xl/styles.xml` bytes for the given [`StylesBuilder`].
pub fn emit(styles: &StylesBuilder) -> Vec<u8> {
    let mut out = String::with_capacity(4096);

    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    out.push_str(
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
    );

    // 1. numFmts — only custom (id >= 164); omit entirely when empty.
    if !styles.num_fmts.is_empty() {
        out.push_str(&format!("<numFmts count=\"{}\">", styles.num_fmts.len()));
        for (id, code) in &styles.num_fmts {
            out.push_str(&format!(
                "<numFmt numFmtId=\"{id}\" formatCode=\"{}\"/>",
                xml_escape::attr(code)
            ));
        }
        out.push_str("</numFmts>");
    }

    // 2. fonts
    out.push_str(&format!("<fonts count=\"{}\">", styles.fonts.len()));
    for font in &styles.fonts {
        out.push_str(&font_to_xml(font));
    }
    out.push_str("</fonts>");

    // 3. fills
    out.push_str(&format!("<fills count=\"{}\">", styles.fills.len()));
    for fill in &styles.fills {
        out.push_str(&fill_to_xml(fill));
    }
    out.push_str("</fills>");

    // 4. borders
    out.push_str(&format!("<borders count=\"{}\">", styles.borders.len()));
    for border in &styles.borders {
        out.push_str(&border_to_xml(border));
    }
    out.push_str("</borders>");

    // 5. cellStyleXfs — singleton required by Excel schema validators.
    out.push_str(
        "<cellStyleXfs count=\"1\">\
         <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\
         </cellStyleXfs>",
    );

    // 6. cellXfs
    out.push_str(&format!("<cellXfs count=\"{}\">", styles.cell_xfs.len()));
    for xf in &styles.cell_xfs {
        out.push_str(&xf_to_xml(xf));
    }
    out.push_str("</cellXfs>");

    // 7. cellStyles — singleton hardcoded Normal style.
    out.push_str(
        "<cellStyles count=\"1\">\
         <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>\
         </cellStyles>",
    );

    // 8. dxfs — differential formats referenced by conditional-formatting
    //    rules. Empty is the common case; when populated, each dxf emits
    //    whichever of <font>/<fill>/<border> is `Some` — number formats
    //    and alignment are not valid children of <dxf>.
    if styles.dxfs.is_empty() {
        out.push_str("<dxfs count=\"0\"/>");
    } else {
        out.push_str(&format!("<dxfs count=\"{}\">", styles.dxfs.len()));
        for dxf in &styles.dxfs {
            out.push_str(&dxf_to_xml(dxf));
        }
        out.push_str("</dxfs>");
    }

    // 9. tableStyles
    out.push_str(
        "<tableStyles count=\"0\" \
         defaultTableStyle=\"TableStyleMedium9\" \
         defaultPivotStyle=\"PivotStyleLight16\"/>",
    );

    out.push_str("</styleSheet>");
    out.into_bytes()
}

// ---------------------------------------------------------------------------
// Per-element helpers (private)
// ---------------------------------------------------------------------------

fn font_to_xml(spec: &FontSpec) -> String {
    let mut parts = String::new();
    if spec.bold {
        parts.push_str("<b/>");
    }
    if spec.italic {
        parts.push_str("<i/>");
    }
    if spec.underline {
        parts.push_str("<u/>");
    }
    if spec.strikethrough {
        parts.push_str("<strike/>");
    }
    if let Some(sz) = spec.size {
        parts.push_str(&format!("<sz val=\"{sz}\"/>"));
    }
    if let Some(ref rgb) = spec.color_rgb {
        // rgb is hex-only, no escape needed; using attr for defense-in-depth.
        parts.push_str(&format!("<color rgb=\"{}\"/>", xml_escape::attr(rgb)));
    }
    if let Some(ref name) = spec.name {
        parts.push_str(&format!("<name val=\"{}\"/>", xml_escape::attr(name)));
    }
    format!("<font>{parts}</font>")
}

fn fill_to_xml(spec: &FillSpec) -> String {
    // OOXML schema: `<patternFill>` may carry `<fgColor>` and/or `<bgColor>`
    // children. Emit both when present, in spec order (fg before bg). When
    // neither is present, self-close the element to match Excel's style.
    let pattern_attr = xml_escape::attr(&spec.pattern_type);
    let has_fg = spec.fg_color_rgb.is_some();
    let has_bg = spec.bg_color_rgb.is_some();

    if !has_fg && !has_bg {
        return format!("<fill><patternFill patternType=\"{pattern_attr}\"/></fill>");
    }

    let mut inner = format!("<patternFill patternType=\"{pattern_attr}\">");
    if let Some(ref rgb) = spec.fg_color_rgb {
        inner.push_str(&format!("<fgColor rgb=\"{}\"/>", xml_escape::attr(rgb)));
    }
    if let Some(ref rgb) = spec.bg_color_rgb {
        inner.push_str(&format!("<bgColor rgb=\"{}\"/>", xml_escape::attr(rgb)));
    }
    inner.push_str("</patternFill>");
    format!("<fill>{inner}</fill>")
}

fn border_side_xml(tag: &str, side: &BorderSideSpec) -> String {
    match (&side.style, &side.color_rgb) {
        (Some(style), Some(rgb)) => {
            format!(
                "<{tag} style=\"{}\"><color rgb=\"{}\"/></{tag}>",
                xml_escape::attr(style),
                xml_escape::attr(rgb)
            )
        }
        (Some(style), None) => format!("<{tag} style=\"{}\"/>", xml_escape::attr(style)),
        _ => format!("<{tag}/>"),
    }
}

fn border_to_xml(spec: &BorderSpec) -> String {
    let mut diag_attrs = String::new();
    if spec.diagonal_up {
        diag_attrs.push_str(" diagonalUp=\"1\"");
    }
    if spec.diagonal_down {
        diag_attrs.push_str(" diagonalDown=\"1\"");
    }
    let mut out = format!("<border{diag_attrs}>");
    out.push_str(&border_side_xml("left", &spec.left));
    out.push_str(&border_side_xml("right", &spec.right));
    out.push_str(&border_side_xml("top", &spec.top));
    out.push_str(&border_side_xml("bottom", &spec.bottom));
    // Diagonal goes through the same renderer as the other sides so its
    // style + color flow through. Empty BorderSideSpec emits `<diagonal/>`
    // which is valid OOXML; non-empty emits the full element.
    out.push_str(&border_side_xml("diagonal", &spec.diagonal));
    out.push_str("</border>");
    out
}

fn alignment_xml(align: &AlignmentSpec) -> String {
    let mut attrs = String::new();
    if let Some(ref h) = align.horizontal {
        attrs.push_str(&format!(" horizontal=\"{}\"", xml_escape::attr(h)));
    }
    if let Some(ref v) = align.vertical {
        attrs.push_str(&format!(" vertical=\"{}\"", xml_escape::attr(v)));
    }
    if align.wrap_text {
        attrs.push_str(" wrapText=\"1\"");
    }
    if align.indent > 0 {
        attrs.push_str(&format!(" indent=\"{}\"", align.indent));
    }
    if align.text_rotation > 0 {
        attrs.push_str(&format!(" textRotation=\"{}\"", align.text_rotation));
    }
    if align.shrink_to_fit {
        attrs.push_str(" shrinkToFit=\"1\"");
    }
    format!("<alignment{attrs}/>")
}

/// Emit one `<dxf>` element. Only the overrides present on the record
/// appear in the output — an all-`None` record emits `<dxf/>`, which is
/// valid but pointless (the builder's dedup prevents that in practice).
fn dxf_to_xml(dxf: &DxfRecord) -> String {
    let mut parts = String::new();
    if let Some(ref font) = dxf.font {
        parts.push_str(&font_to_xml(font));
    }
    if let Some(ref fill) = dxf.fill {
        parts.push_str(&fill_to_xml(fill));
    }
    if let Some(ref border) = dxf.border {
        parts.push_str(&border_to_xml(border));
    }
    if parts.is_empty() {
        "<dxf/>".to_string()
    } else {
        format!("<dxf>{parts}</dxf>")
    }
}

fn xf_to_xml(xf: &XfRecord) -> String {
    let mut attrs = format!(
        "numFmtId=\"{}\" fontId=\"{}\" fillId=\"{}\" borderId=\"{}\" xfId=\"0\"",
        xf.num_fmt_id, xf.font_id, xf.fill_id, xf.border_id
    );
    if xf.apply_font {
        attrs.push_str(" applyFont=\"1\"");
    }
    if xf.apply_fill {
        attrs.push_str(" applyFill=\"1\"");
    }
    if xf.apply_border {
        attrs.push_str(" applyBorder=\"1\"");
    }
    if xf.apply_number_format {
        attrs.push_str(" applyNumberFormat=\"1\"");
    }

    if let Some(ref align) = xf.alignment {
        if xf.apply_alignment {
            attrs.push_str(" applyAlignment=\"1\"");
            let align_xml = alignment_xml(align);
            return format!("<xf {attrs}>{align_xml}</xf>");
        }
    }

    format!("<xf {attrs}/>")
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::format::{
        AlignmentSpec, BorderSideSpec, BorderSpec, FillSpec, FontSpec, FormatSpec, StylesBuilder,
    };
    use quick_xml::events::Event;
    use quick_xml::Reader;

    /// Parse the given bytes as XML. Panics on any parse error.
    fn parse_ok(bytes: &[u8]) {
        let text = std::str::from_utf8(bytes).expect("output must be valid UTF-8");
        let mut reader = Reader::from_str(text);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Eof) => break,
                Err(e) => panic!("XML parse error: {e}\n\nFull output:\n{text}"),
                _ => {}
            }
            buf.clear();
        }
    }

    /// Convenience: emit and convert to String.
    fn emit_str(styles: &StylesBuilder) -> String {
        String::from_utf8(emit(styles)).expect("output must be valid UTF-8")
    }

    // -------------------------------------------------------------------
    // Test 1: empty builder emits valid skeleton
    // -------------------------------------------------------------------
    #[test]
    fn empty_builder_emits_valid_skeleton() {
        let styles = StylesBuilder::default();
        let bytes = emit(&styles);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();

        assert!(text.contains("<styleSheet"), "missing <styleSheet");
        assert!(
            text.contains("<fonts count=\"1\">"),
            "missing <fonts count=\"1\">"
        );
        assert!(
            text.contains("<fills count=\"2\">"),
            "missing <fills count=\"2\">"
        );
        assert!(
            text.contains("<borders count=\"1\">"),
            "missing <borders count=\"1\">"
        );
        assert!(
            text.contains("<cellStyleXfs count=\"1\">"),
            "missing <cellStyleXfs count=\"1\">"
        );
        assert!(
            text.contains("<cellXfs count=\"1\">"),
            "missing <cellXfs count=\"1\">"
        );
        assert!(
            text.contains("<cellStyles count=\"1\">"),
            "missing <cellStyles count=\"1\">"
        );
        assert!(
            text.contains("<dxfs count=\"0\"/>"),
            "missing <dxfs count=\"0\"/>"
        );
        assert!(
            text.contains("<tableStyles count=\"0\""),
            "missing <tableStyles count=\"0\""
        );
        assert!(
            !text.contains("<numFmts"),
            "numFmts should be omitted when empty"
        );
    }

    // -------------------------------------------------------------------
    // Test 2: fills reserved slots always present
    // -------------------------------------------------------------------
    #[test]
    fn fills_reserved_slots_always_present() {
        let styles = StylesBuilder::default();
        let text = emit_str(&styles);

        assert!(
            text.contains("<patternFill patternType=\"none\"/>"),
            "slot 0 (none) missing"
        );
        assert!(
            text.contains("<patternFill patternType=\"gray125\"/>"),
            "slot 1 (gray125) missing"
        );
    }

    // -------------------------------------------------------------------
    // Test 3: font bold+italic+color round-trips
    // -------------------------------------------------------------------
    #[test]
    fn font_bold_italic_color_round_trips() {
        let mut styles = StylesBuilder::default();
        let spec = FontSpec {
            bold: true,
            italic: true,
            color_rgb: Some("FFFF0000".into()),
            size: Some(12),
            ..Default::default()
        };
        styles.intern_font(&spec);
        let bytes = emit(&styles);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();

        // The new font (index 1) should be present.
        assert!(text.contains("<b/>"), "missing <b/>");
        assert!(text.contains("<i/>"), "missing <i/>");
        assert!(text.contains("<color rgb=\"FFFF0000\"/>"), "missing color");
        assert!(text.contains("<sz val=\"12\"/>"), "missing size");
    }

    // -------------------------------------------------------------------
    // Test 4: alignment only emitted when apply_alignment is true
    // -------------------------------------------------------------------
    #[test]
    fn alignment_only_emitted_when_apply_alignment_true() {
        let mut styles = StylesBuilder::default();
        let spec = FormatSpec {
            alignment: Some(AlignmentSpec {
                horizontal: Some("center".into()),
                ..Default::default()
            }),
            ..Default::default()
        };
        styles.intern_format(&spec);
        let text = emit_str(&styles);

        assert!(
            text.contains("applyAlignment=\"1\""),
            "applyAlignment flag missing"
        );
        assert!(
            text.contains("<alignment horizontal=\"center\"/>"),
            "alignment child missing"
        );
    }

    // -------------------------------------------------------------------
    // Test 5: alignment not emitted when absent
    // -------------------------------------------------------------------
    #[test]
    fn alignment_not_emitted_when_absent() {
        let styles = StylesBuilder::default();
        let text = emit_str(&styles);

        // The default xf at index 0 should self-close.
        assert!(
            text.contains(
                "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
            ),
            "default xf should self-close without alignment; got:\n{text}"
        );
    }

    // -------------------------------------------------------------------
    // Test 6: numFmts element omitted when no user formats
    // -------------------------------------------------------------------
    #[test]
    fn numfmts_element_omitted_when_no_user_formats() {
        let styles = StylesBuilder::default();
        let text = emit_str(&styles);
        assert!(!text.contains("<numFmts"), "numFmts should be absent");
    }

    // -------------------------------------------------------------------
    // Test 7: numFmts element present with user formats
    // -------------------------------------------------------------------
    #[test]
    fn numfmts_element_present_with_user_formats() {
        let mut styles = StylesBuilder::default();
        styles.intern_num_fmt("0.0000");
        let bytes = emit(&styles);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();

        assert!(
            text.contains("<numFmts count=\"1\">"),
            "missing numFmts count=1"
        );
        assert!(
            text.contains("numFmtId=\"164\""),
            "custom fmt should have id 164"
        );
        assert!(
            text.contains("formatCode=\"0.0000\""),
            "format code missing"
        );
    }

    // -------------------------------------------------------------------
    // Test 8: custom numFmt id starts at 164
    // -------------------------------------------------------------------
    #[test]
    fn numfmt_custom_id_starts_at_164() {
        let mut styles = StylesBuilder::default();
        let id = styles.intern_num_fmt("my custom");
        assert!(id >= 164, "custom id should be >= 164, got {id}");
    }

    // -------------------------------------------------------------------
    // Test 9: well-formed under quick_xml reader (non-trivial builder)
    // -------------------------------------------------------------------
    #[test]
    fn xml_well_formed_under_quick_xml_reader() {
        let mut styles = StylesBuilder::default();

        // Several fonts.
        styles.intern_font(&FontSpec {
            bold: true,
            size: Some(14),
            name: Some("Arial".into()),
            ..Default::default()
        });
        styles.intern_font(&FontSpec {
            italic: true,
            color_rgb: Some("FF0000FF".into()),
            ..Default::default()
        });

        // Custom fill.
        styles.intern_fill(&FillSpec {
            pattern_type: "solid".into(),
            fg_color_rgb: Some("FFFFFF00".into()),
            ..Default::default()
        });

        // Custom border.
        styles.intern_border(&BorderSpec {
            left: BorderSideSpec {
                style: Some("thin".into()),
                color_rgb: None,
            },
            ..Default::default()
        });

        // Custom numFmt.
        styles.intern_num_fmt("yyyy-mm-dd hh:mm");

        // Alignment format.
        styles.intern_format(&FormatSpec {
            alignment: Some(AlignmentSpec {
                horizontal: Some("right".into()),
                wrap_text: true,
                ..Default::default()
            }),
            ..Default::default()
        });

        let bytes = emit(&styles);
        parse_ok(&bytes);
    }

    // -------------------------------------------------------------------
    // Test 10: applyFont / applyFill / applyBorder flags written
    // -------------------------------------------------------------------
    #[test]
    fn applyfont_applyfill_applyborder_flags_written() {
        let mut styles = StylesBuilder::default();
        let spec = FormatSpec {
            font: Some(FontSpec {
                bold: true,
                ..Default::default()
            }),
            fill: Some(FillSpec {
                pattern_type: "solid".into(),
                fg_color_rgb: Some("FFFF0000".into()),
                ..Default::default()
            }),
            border: Some(BorderSpec {
                left: BorderSideSpec {
                    style: Some("thin".into()),
                    ..Default::default()
                },
                ..Default::default()
            }),
            ..Default::default()
        };
        styles.intern_format(&spec);
        let text = emit_str(&styles);

        assert!(text.contains("applyFont=\"1\""), "applyFont flag missing");
        assert!(text.contains("applyFill=\"1\""), "applyFill flag missing");
        assert!(
            text.contains("applyBorder=\"1\""),
            "applyBorder flag missing"
        );
    }

    // -------------------------------------------------------------------
    // Test 11: xf attributes order is deterministic (byte equality)
    // -------------------------------------------------------------------
    #[test]
    fn xf_attributes_order_deterministic() {
        let mut styles = StylesBuilder::default();
        styles.intern_format(&FormatSpec {
            font: Some(FontSpec {
                bold: true,
                ..Default::default()
            }),
            fill: Some(FillSpec {
                pattern_type: "solid".into(),
                fg_color_rgb: Some("FFABCDEF".into()),
                ..Default::default()
            }),
            ..Default::default()
        });

        let first = emit(&styles);
        let second = emit(&styles);
        assert_eq!(first, second, "emit must be deterministic");
    }

    // -------------------------------------------------------------------
    // Test 12: builtin numFmt does not create custom entry
    // -------------------------------------------------------------------
    #[test]
    fn builtin_numfmt_does_not_create_custom_entry() {
        let mut styles = StylesBuilder::default();
        let id = styles.intern_num_fmt("0.00");
        assert_eq!(id, 2, "0.00 is builtin id 2");
        assert!(
            styles.num_fmts.is_empty(),
            "no custom entry should be created"
        );
        let text = emit_str(&styles);
        assert!(!text.contains("<numFmts"), "numFmts should be absent");
    }

    // -------------------------------------------------------------------
    // Test 13: cellStyles block is always emitted
    // -------------------------------------------------------------------
    #[test]
    fn cellstyles_block_is_always_emitted() {
        let styles = StylesBuilder::default();
        let text = emit_str(&styles);

        assert!(
            text.contains("<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"),
            "Normal cellStyle missing"
        );
        assert!(
            text.contains("<cellStyles count=\"1\">"),
            "cellStyles wrapper missing"
        );
    }

    // -------------------------------------------------------------------
    // Test 14: attribute XML escape in numFmt format code
    // -------------------------------------------------------------------
    #[test]
    fn attribute_xml_escape_in_numfmt_format_code() {
        let mut styles = StylesBuilder::default();
        styles.intern_num_fmt(r#"[Red]"A&B""#);
        let text = emit_str(&styles);

        assert!(
            text.contains("&amp;"),
            "& must be escaped as &amp; in formatCode attribute; got:\n{text}"
        );
        assert!(
            !text.contains("A&B"),
            "raw & must not appear unescaped; got:\n{text}"
        );
    }

    // -------------------------------------------------------------------
    // Test 15: border side with color and style
    // -------------------------------------------------------------------
    #[test]
    fn border_side_color_and_style() {
        let mut styles = StylesBuilder::default();
        let border = BorderSpec {
            left: BorderSideSpec {
                style: Some("thin".into()),
                color_rgb: Some("FF0000FF".into()),
            },
            ..Default::default()
        };
        styles.intern_border(&border);
        let bytes = emit(&styles);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();

        assert!(
            text.contains("<left style=\"thin\"><color rgb=\"FF0000FF\"/></left>"),
            "border side with color not rendered correctly; got:\n{text}"
        );
    }

    // -------------------------------------------------------------------
    // Test 16: border rgb with special chars is attr-escaped
    // -------------------------------------------------------------------
    #[test]
    fn border_rgb_with_special_chars_is_attr_escaped() {
        // A pathological RGB - real RGB is always hex but the emitter is
        // defense-in-depth: if callers ever pass `"` or `&`, output stays valid.
        let mut b = StylesBuilder::default();
        let border = BorderSpec {
            left: BorderSideSpec {
                style: Some("thin".into()),
                color_rgb: Some(r#"FF"00&00"#.into()),
            },
            ..Default::default()
        };
        b.intern_border(&border);
        let text = String::from_utf8(emit(&b)).unwrap();
        // Must not contain the raw quote or ampersand inside the attribute.
        assert!(
            text.contains(r#"<color rgb="FF&quot;00&amp;00"/>"#),
            "got: {text}"
        );
        assert!(!text.contains(r#"rgb="FF"00"#), "raw quote leaked: {text}");
    }

    // -------------------------------------------------------------------
    // Test 17: fgColor rgb with special chars is attr-escaped
    // -------------------------------------------------------------------
    #[test]
    fn fill_fgcolor_rgb_with_special_chars_is_attr_escaped() {
        let mut b = StylesBuilder::default();
        b.intern_fill(&FillSpec {
            pattern_type: "solid".into(),
            fg_color_rgb: Some(r#"FF"00&00"#.into()),
            ..Default::default()
        });
        let text = String::from_utf8(emit(&b)).unwrap();
        assert!(
            text.contains(r#"<fgColor rgb="FF&quot;00&amp;00"/>"#),
            "got: {text}"
        );
        assert!(!text.contains(r#"rgb="FF"00"#), "raw quote leaked: {text}");
    }

    // -------------------------------------------------------------------
    // Test: dxfs empty — emits self-closing <dxfs count="0"/>
    // -------------------------------------------------------------------
    #[test]
    fn dxfs_empty_emits_self_closing() {
        let styles = StylesBuilder::default();
        let text = emit_str(&styles);
        assert!(
            text.contains("<dxfs count=\"0\"/>"),
            "missing empty dxfs self-closing: {text}"
        );
    }

    // -------------------------------------------------------------------
    // Test: dxfs non-empty — emits <dxfs count="N"> with child <dxf>s
    // -------------------------------------------------------------------
    #[test]
    fn dxfs_non_empty_emits_wrapper_and_children() {
        use crate::model::format::DxfRecord;
        let mut styles = StylesBuilder::default();
        let id = styles.intern_dxf(&DxfRecord {
            font: Some(FontSpec {
                bold: true,
                color_rgb: Some("FFFF0000".into()),
                ..Default::default()
            }),
            ..Default::default()
        });
        assert_eq!(id, 0, "first dxf interns at index 0");

        // Dedup should return same index.
        let id2 = styles.intern_dxf(&DxfRecord {
            font: Some(FontSpec {
                bold: true,
                color_rgb: Some("FFFF0000".into()),
                ..Default::default()
            }),
            ..Default::default()
        });
        assert_eq!(id2, 0, "identical dxf should dedup");

        let bytes = emit(&styles);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<dxfs count=\"1\">"),
            "missing count-1 wrapper: {text}"
        );
        assert!(text.contains("<dxf>"), "missing <dxf> child: {text}");
        assert!(text.contains("<b/>"), "missing bold inside dxf: {text}");
        assert!(
            text.contains("FFFF0000"),
            "missing red color inside dxf: {text}"
        );
        assert!(text.contains("</dxfs>"), "missing </dxfs>: {text}");
    }

    // -------------------------------------------------------------------
    // W4E.R2: diagonal style + color must round-trip through `<diagonal>`
    // -------------------------------------------------------------------
    #[test]
    fn border_diagonal_style_and_color_emit() {
        let mut styles = StylesBuilder::default();
        let border = BorderSpec {
            diagonal: BorderSideSpec {
                style: Some("thin".into()),
                color_rgb: Some("FF112233".into()),
            },
            ..Default::default()
        };
        styles.intern_border(&border);
        let bytes = emit(&styles);
        parse_ok(&bytes);
        let text = String::from_utf8(bytes).unwrap();
        assert!(
            text.contains("<diagonal style=\"thin\"><color rgb=\"FF112233\"/></diagonal>"),
            "diagonal style + color must round-trip; got:\n{text}"
        );
    }

    // -------------------------------------------------------------------
    // W4E.R2: diagonalUp / diagonalDown attrs land on the parent <border>
    // -------------------------------------------------------------------
    #[test]
    fn border_diagonal_up_and_down_attrs_present() {
        let mut styles = StylesBuilder::default();
        let border = BorderSpec {
            diagonal: BorderSideSpec {
                style: Some("thin".into()),
                color_rgb: None,
            },
            diagonal_up: true,
            diagonal_down: true,
            ..Default::default()
        };
        styles.intern_border(&border);
        let text = String::from_utf8(emit(&styles)).unwrap();
        assert!(
            text.contains("<border diagonalUp=\"1\" diagonalDown=\"1\">"),
            "diagonalUp + diagonalDown attrs must land on <border>; got:\n{text}"
        );
    }

    // -------------------------------------------------------------------
    // W4E.R3: fill bg_color_rgb must round-trip via `<bgColor>`
    // -------------------------------------------------------------------
    #[test]
    fn fill_with_bg_color_emits_bg_element() {
        let mut styles = StylesBuilder::default();
        styles.intern_fill(&FillSpec {
            pattern_type: "darkHorizontal".into(),
            fg_color_rgb: None,
            bg_color_rgb: Some("FFAABBCC".into()),
        });
        let text = String::from_utf8(emit(&styles)).unwrap();
        assert!(
            text.contains("<bgColor rgb=\"FFAABBCC\"/>"),
            "bg_color_rgb must round-trip; got:\n{text}"
        );
    }

    // -------------------------------------------------------------------
    // W4E.R3: fg + bg colors both round-trip in spec order (fg before bg)
    // -------------------------------------------------------------------
    #[test]
    fn fill_with_fg_and_bg_emits_both_in_order() {
        let mut styles = StylesBuilder::default();
        styles.intern_fill(&FillSpec {
            pattern_type: "darkHorizontal".into(),
            fg_color_rgb: Some("FF111111".into()),
            bg_color_rgb: Some("FF222222".into()),
        });
        let text = String::from_utf8(emit(&styles)).unwrap();
        let fg_pos = text
            .find("<fgColor rgb=\"FF111111\"/>")
            .expect("fg present");
        let bg_pos = text
            .find("<bgColor rgb=\"FF222222\"/>")
            .expect("bg present");
        assert!(fg_pos < bg_pos, "fg must precede bg in OOXML; got:\n{text}");
    }

    // -------------------------------------------------------------------
    // W4E.R3: empty fill (no fg, no bg) still self-closes — no regression
    // -------------------------------------------------------------------
    #[test]
    fn fill_with_no_colors_self_closes() {
        let mut styles = StylesBuilder::default();
        styles.intern_fill(&FillSpec {
            pattern_type: "gray0625".into(),
            fg_color_rgb: None,
            bg_color_rgb: None,
        });
        let text = String::from_utf8(emit(&styles)).unwrap();
        assert!(
            text.contains("<patternFill patternType=\"gray0625\"/>"),
            "empty pattern fill must self-close; got:\n{text}"
        );
    }

    // -------------------------------------------------------------------
    // W4E.R2: empty diagonal still emits valid `<diagonal/>` (no regression)
    // -------------------------------------------------------------------
    #[test]
    fn border_empty_diagonal_emits_self_closing() {
        let mut styles = StylesBuilder::default();
        let border = BorderSpec {
            left: BorderSideSpec {
                style: Some("thin".into()),
                color_rgb: None,
            },
            ..Default::default()
        };
        styles.intern_border(&border);
        let text = String::from_utf8(emit(&styles)).unwrap();
        // No diagonalUp/diagonalDown when both are false.
        assert!(
            !text.contains("diagonalUp") && !text.contains("diagonalDown"),
            "empty-diag border must not emit up/down attrs: {text}"
        );
        assert!(
            text.contains("<diagonal/>"),
            "empty diagonal must remain self-closing: {text}"
        );
    }
}
