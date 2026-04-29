//! `xl/charts/chartN.xml` emitter.
//!
//! One `xl/charts/chartN.xml` per [`Chart`] anchored on a sheet. The
//! emitter produces a `<c:chartSpace>` element with the `c:` (chart) and
//! `a:` (drawingml) namespaces declared on the root, matching openpyxl's
//! emit shape so byte-parity tests can pass downstream.
//!
//! # Element ordering
//!
//! The OOXML spec is strict about child-element order inside
//! `<chart>` and `<plotArea>`. Both elements use a sequence model:
//!
//! ```text
//! <chart>
//!   [title?]
//!   [autoTitleDeleted?]
//!   <plotArea>
//!     [layout?]
//!     <{kind}Chart>
//!       [varyColors?]
//!       <ser>+
//!       [gapWidth? / overlap? / holeSize? / scatterStyle? / radarStyle? / ...]
//!       <axId>+   <!-- 0 axis ids for Pie, 2 for Bar/Line/Area/Radar/Scatter/Bubble -->
//!     </{kind}Chart>
//!     [catAx?] [valAx?] [dateAx?] [serAx?]
//!   </plotArea>
//!   [legend?]
//!   [plotVisOnly?]
//!   [dispBlanksAs?]
//! </chart>
//! ```
//!
//! Inside `<ser>`:
//!
//! ```text
//! <idx/> <order/> [tx?] [spPr?] [marker?] [dPt*] [dLbls?]
//! [errBars*] [trendline*] [cat?] [val?] [xVal?] [yVal?] [bubbleSize?]
//! [smooth?]
//! ```
//!
//! Skipping any optional sub-element produces no XML — this matches
//! openpyxl's "leave it off" rule.

mod axes;
mod layout;
mod pivot;
mod plot;
mod primitives;
mod series;
mod style;
mod text;

use crate::model::chart::Chart;

use axes::{emit_axis, pick_axis_ids};
use layout::{emit_layout, emit_legend, emit_view_3d};
use pivot::emit_pivot_source;
use plot::emit_plot_chart;
use primitives::bool_str;
use text::emit_title;

const C_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Emit `xl/charts/chartN.xml` bytes from `chart`. `axis_id_a` and
/// `axis_id_b` are the per-chart axis ids (typically 1-based, distinct
/// within a chart). Pie/Doughnut ignore them.
pub fn emit_chart_xml(chart: &Chart) -> Vec<u8> {
    // Allocate axis ids deterministically — they only need to be
    // unique within this chart. Use 10 + 100 like openpyxl's golden
    // example to match its emission shape.
    let (ax_id_a, ax_id_b) = pick_axis_ids(chart);

    let mut out = String::with_capacity(2048);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    out.push_str(&format!(
        "<c:chartSpace xmlns:c=\"{C_NS}\" xmlns:a=\"{A_NS}\" xmlns:r=\"{R_NS}\">"
    ));

    out.push_str("<c:chart>");

    // Per ECMA-376 Part 1 section 21.2.2.27, pivotSource is the first
    // optional child in the chart sequence.
    if let Some(ps) = &chart.pivot_source {
        emit_pivot_source(&mut out, ps);
    }

    // Optional <c:title>
    if let Some(t) = &chart.title {
        emit_title(&mut out, t);
    }

    // Auto-title deleted is implicit when title is absent + we emit no element.

    // Optional <c:view3D> for 3D variants.
    if chart.kind.is_3d() {
        if let Some(v) = &chart.view_3d {
            emit_view_3d(&mut out, v);
        }
    }

    // <c:plotArea>
    out.push_str("<c:plotArea>");
    if let Some(layout) = &chart.layout {
        emit_layout(&mut out, layout);
    }

    // Per-type chart element.
    emit_plot_chart(&mut out, chart, ax_id_a, ax_id_b);

    // Axes (after the chart-type element).
    if !chart.kind.is_axis_free() {
        if let Some(x) = &chart.x_axis {
            emit_axis(&mut out, x);
        }
        if let Some(y) = &chart.y_axis {
            emit_axis(&mut out, y);
        }
    }

    out.push_str("</c:plotArea>");

    // Legend.
    if let Some(legend) = &chart.legend {
        emit_legend(&mut out, legend);
    }

    if let Some(v) = chart.plot_visible_only {
        out.push_str(&format!("<c:plotVisOnly val=\"{}\"/>", bool_str(v)));
    }
    if let Some(d) = chart.display_blanks_as {
        out.push_str(&format!("<c:dispBlanksAs val=\"{}\"/>", d.as_str()));
    }

    out.push_str("</c:chart>");
    out.push_str("</c:chartSpace>");
    out.into_bytes()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::chart::{
        Axis, AxisCommon, AxisPos, CategoryAxis, Chart, ChartKind, DataLabels, ErrorBarType,
        ErrorBarValType, ErrorBars, Legend, LegendPosition, Marker, MarkerSymbol, PivotSource,
        Reference, Series, SeriesTitle, Title, Trendline, TrendlineKind, ValueAxis,
    };
    use crate::model::image::ImageAnchor;
    use quick_xml::events::Event;
    use quick_xml::Reader;

    fn parse_ok(bytes: &[u8]) {
        let text = std::str::from_utf8(bytes).expect("utf8");
        let mut reader = Reader::from_str(text);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Eof) => break,
                Err(e) => panic!("parse error: {e}\nBytes:\n{text}"),
                _ => (),
            }
            buf.clear();
        }
    }

    fn bar_chart_with_one_series() -> Chart {
        let mut c = Chart::new(ChartKind::Bar, ImageAnchor::one_cell(0, 0));
        c.title = Some(Title::plain("Sales"));
        let mut s = Series::new(0);
        s.title = Some(SeriesTitle::StrRef(Reference::new("Sheet", "B1")));
        s.categories = Some(Reference::new("Sheet", "A2:A6"));
        s.values = Some(Reference::new("Sheet", "B2:B6"));
        c.add_series(s);
        c.x_axis = Some(Axis::Category(CategoryAxis {
            common: AxisCommon::new(10, 100, AxisPos::Bottom),
            lbl_offset: Some(100),
            lbl_algn: None,
        }));
        c.y_axis = Some(Axis::Value(ValueAxis {
            common: {
                let mut a = AxisCommon::new(100, 10, AxisPos::Left);
                a.major_gridlines = true;
                a
            },
            min: None,
            max: None,
            major_unit: None,
            minor_unit: None,
            display_units: None,
            crosses: None,
        }));
        c
    }

    #[test]
    fn bar_chart_has_correct_plot_element() {
        let c = bar_chart_with_one_series();
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<c:barChart>"), "missing barChart: {text}");
        assert!(text.contains("<c:barDir val=\"col\"/>"));
        assert!(text.contains("<c:grouping val=\"clustered\"/>"));
        assert!(text.contains("<c:gapWidth val=\"150\"/>"));
        // Both axes referenced.
        assert!(text.contains("<c:axId val=\"10\"/>"));
        assert!(text.contains("<c:axId val=\"100\"/>"));
        // Title present.
        assert!(text.contains("<a:t>Sales</a:t>"));
    }

    #[test]
    fn pie_chart_has_no_axes() {
        let mut c = Chart::new(ChartKind::Pie, ImageAnchor::one_cell(0, 0));
        let mut s = Series::new(0);
        s.values = Some(Reference::new("Sheet", "B2:B6"));
        s.categories = Some(Reference::new("Sheet", "A2:A6"));
        c.add_series(s);
        c.vary_colors = Some(true);
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<c:pieChart>"));
        assert!(text.contains("<c:varyColors val=\"1\"/>"));
        // No axId / catAx / valAx for Pie.
        assert!(!text.contains("<c:catAx>"));
        assert!(!text.contains("<c:valAx>"));
        assert!(!text.contains("<c:axId"));
    }

    #[test]
    fn doughnut_emits_hole_size() {
        let mut c = Chart::new(ChartKind::Doughnut, ImageAnchor::one_cell(0, 0));
        let mut s = Series::new(0);
        s.values = Some(Reference::new("S", "B2:B5"));
        c.add_series(s);
        c.hole_size = Some(60);
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<c:doughnutChart>"));
        assert!(text.contains("<c:holeSize val=\"60\"/>"));
    }

    #[test]
    fn scatter_uses_x_y_val() {
        let mut c = Chart::new(ChartKind::Scatter, ImageAnchor::one_cell(0, 0));
        let mut s = Series::new(0);
        s.x_values = Some(Reference::new("S", "A2:A6"));
        s.values = Some(Reference::new("S", "B2:B6"));
        c.add_series(s);
        c.x_axis = Some(Axis::Value(ValueAxis {
            common: AxisCommon::new(10, 100, AxisPos::Bottom),
            min: None,
            max: None,
            major_unit: None,
            minor_unit: None,
            display_units: None,
            crosses: None,
        }));
        c.y_axis = Some(Axis::Value(ValueAxis {
            common: AxisCommon::new(100, 10, AxisPos::Left),
            min: None,
            max: None,
            major_unit: None,
            minor_unit: None,
            display_units: None,
            crosses: None,
        }));
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<c:scatterChart>"));
        assert!(!text.contains("<c:scatterStyle"));
        assert!(text.contains("<c:xVal>"));
        assert!(text.contains("<c:yVal>"));
        assert!(!text.contains("<c:cat>"));
    }

    #[test]
    fn legend_position_emitted() {
        let mut c = bar_chart_with_one_series();
        c.legend = Some(Legend {
            position: LegendPosition::Top,
            overlay: None,
            layout: None,
        });
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<c:legendPos val=\"t\"/>"));
    }

    #[test]
    fn data_labels_emitted_on_series() {
        let mut c = bar_chart_with_one_series();
        c.series[0].data_labels = Some(DataLabels {
            show_val: Some(true),
            show_cat_name: Some(true),
            position: Some("outEnd".to_string()),
            ..Default::default()
        });
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<c:showVal val=\"1\"/>"));
        assert!(text.contains("<c:showCatName val=\"1\"/>"));
        assert!(text.contains("<c:dLblPos val=\"outEnd\"/>"));
    }

    #[test]
    fn error_bars_emitted() {
        let mut c = bar_chart_with_one_series();
        c.series[0].error_bars.push(ErrorBars {
            bar_type: ErrorBarType::Both,
            val_type: ErrorBarValType::FixedVal,
            value: Some(1.5),
            no_end_cap: Some(false),
        });
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<c:errBars>"));
        assert!(text.contains("<c:errBarType val=\"both\"/>"));
        assert!(text.contains("<c:errValType val=\"fixedVal\"/>"));
        assert!(text.contains("<c:val val=\"1.5\"/>"));
    }

    #[test]
    fn trendline_polynomial_emits_order() {
        let mut c = bar_chart_with_one_series();
        c.series[0].trendlines.push(Trendline {
            kind: TrendlineKind::Polynomial,
            order: Some(3),
            period: None,
            forward: None,
            backward: None,
            display_equation: Some(true),
            display_r_squared: None,
            name: Some("My Fit".to_string()),
        });
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<c:trendline>"));
        assert!(text.contains("<c:trendlineType val=\"poly\"/>"));
        assert!(text.contains("<c:order val=\"3\"/>"));
        assert!(text.contains("<c:dispEq val=\"1\"/>"));
        assert!(text.contains("<c:name>My Fit</c:name>"));
    }

    #[test]
    fn marker_emitted_on_series() {
        let mut c = Chart::new(ChartKind::Line, ImageAnchor::one_cell(0, 0));
        let mut s = Series::new(0);
        s.values = Some(Reference::new("S", "B2:B6"));
        s.marker = Some(Marker {
            symbol: MarkerSymbol::Diamond,
            size: Some(7),
            graphical_properties: None,
        });
        c.add_series(s);
        c.x_axis = Some(Axis::Category(CategoryAxis {
            common: AxisCommon::new(10, 100, AxisPos::Bottom),
            lbl_offset: None,
            lbl_algn: None,
        }));
        c.y_axis = Some(Axis::Value(ValueAxis {
            common: AxisCommon::new(100, 10, AxisPos::Left),
            min: None,
            max: None,
            major_unit: None,
            minor_unit: None,
            display_units: None,
            crosses: None,
        }));
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<c:marker>"));
        assert!(text.contains("<c:symbol val=\"diamond\"/>"));
        assert!(text.contains("<c:size val=\"7\"/>"));
    }

    #[test]
    fn smooth_only_on_line_or_scatter() {
        let mut c = Chart::new(ChartKind::Line, ImageAnchor::one_cell(0, 0));
        let mut s = Series::new(0);
        s.smooth = Some(true);
        s.values = Some(Reference::new("S", "B2:B6"));
        c.add_series(s);
        c.x_axis = Some(Axis::Category(CategoryAxis {
            common: AxisCommon::new(10, 100, AxisPos::Bottom),
            lbl_offset: None,
            lbl_algn: None,
        }));
        c.y_axis = Some(Axis::Value(ValueAxis {
            common: AxisCommon::new(100, 10, AxisPos::Left),
            min: None,
            max: None,
            major_unit: None,
            minor_unit: None,
            display_units: None,
            crosses: None,
        }));
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<c:smooth val=\"1\"/>"));
    }

    // ----------------------------------------------------------------
    // Pivot-chart linkage.
    // ----------------------------------------------------------------

    #[test]
    fn pivot_source_emitted_at_top_of_chart_with_per_series_fmt_id() {
        let mut c = bar_chart_with_one_series();
        c.pivot_source = Some(PivotSource {
            name: "MyPivot".into(),
            fmt_id: 0,
        });
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        // 1) `<c:pivotSource>` block appears immediately after
        //    `<c:chart>` open and BEFORE `<c:title>`.
        let chart_open = text.find("<c:chart>").expect("chart open");
        let pivot_src = text
            .find("<c:pivotSource>")
            .expect("pivotSource missing when set");
        let title_open = text.find("<c:title>").expect("title open");
        assert!(
            chart_open < pivot_src && pivot_src < title_open,
            "ordering wrong: chart={chart_open} pivotSource={pivot_src} title={title_open}\n{text}"
        );
        // 2) Block content matches the §10.1 byte-shape exactly.
        assert!(text.contains(
            "<c:pivotSource><c:name>MyPivot</c:name><c:fmtId val=\"0\"/></c:pivotSource>"
        ));
        // 3) Per-series `<c:fmtId val="0"/>` injected RIGHT AFTER the
        //    series-order block. Excel rejects pivot charts whose
        //    series lack `<c:fmtId>`.
        assert!(
            text.contains("<c:order val=\"0\"/><c:fmtId val=\"0\"/>"),
            "missing per-series fmtId after order: {text}"
        );
    }

    #[test]
    fn pivot_source_omitted_when_none() {
        let c = bar_chart_with_one_series();
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(
            !text.contains("<c:pivotSource"),
            "should not emit pivotSource when None"
        );
        assert!(
            !text.contains("<c:fmtId"),
            "should not emit per-series fmtId when no pivot_source"
        );
    }

    #[test]
    fn pivot_source_name_xml_escaped() {
        let mut c = bar_chart_with_one_series();
        c.pivot_source = Some(PivotSource {
            name: "Sheet & Co".into(),
            fmt_id: 7,
        });
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<c:name>Sheet &amp; Co</c:name>"));
        assert!(text.contains("<c:fmtId val=\"7\"/>"));
    }
}
