//! Integration tests for the chart emit pipeline — Sprint Μ Pod-α
//! (RFC-046).
//!
//! Each test builds a [`Chart`] in memory, emits its
//! `xl/charts/chartN.xml` via [`wolfxl_writer::emit::charts::emit_chart_xml`],
//! then re-parses the bytes with `quick-xml` to assert structural
//! invariants. Eight kind-specific tests + 12 sub-feature tests.
//!
//! Pod-β (Python) and Pod-δ (cross-language byte parity) will layer
//! their own tests on top; this file owns the Rust-only contract.

use quick_xml::events::Event;
use quick_xml::Reader;

use wolfxl_writer::emit::charts::emit_chart_xml;
use wolfxl_writer::model::chart::{
    Axis, AxisCommon, AxisOrientation, AxisPos, BarGrouping, CategoryAxis, Chart, ChartKind,
    DataLabels, DateAxis, DisplayBlanksAs, ErrorBarType, ErrorBarValType, ErrorBars,
    GraphicalProperties, Layout, Legend, LegendPosition, Marker, MarkerSymbol, RadarStyle,
    Reference, Series, SeriesAxis, SeriesTitle, TickMark, Title, TitleRun, Trendline,
    TrendlineKind, ValueAxis,
};
use wolfxl_writer::model::image::ImageAnchor;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn parse_ok(bytes: &[u8]) {
    let text = std::str::from_utf8(bytes).expect("emit produced invalid utf-8");
    let mut reader = Reader::from_str(text);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Err(e) => panic!("emit produced invalid xml: {e}\n\n{text}"),
            _ => (),
        }
        buf.clear();
    }
}

fn assert_contains(bytes: &[u8], needle: &str) {
    let text = std::str::from_utf8(bytes).expect("utf-8");
    assert!(
        text.contains(needle),
        "expected output to contain {needle:?}, got:\n{text}"
    );
}

fn assert_not_contains(bytes: &[u8], needle: &str) {
    let text = std::str::from_utf8(bytes).expect("utf-8");
    assert!(
        !text.contains(needle),
        "expected output NOT to contain {needle:?}, got:\n{text}"
    );
}

fn anchor_at(col: u32, row: u32) -> ImageAnchor {
    ImageAnchor::one_cell(col, row)
}

fn cat_value_pair() -> (Axis, Axis) {
    let cat = Axis::Category(CategoryAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        lbl_offset: Some(100),
        lbl_algn: None,
    });
    let val = Axis::Value(ValueAxis {
        common: {
            let mut a = AxisCommon::new(100, 10, AxisPos::Left);
            a.major_gridlines = true;
            a
        },
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        crosses: None,
    });
    (cat, val)
}

fn dual_value_pair() -> (Axis, Axis) {
    let x = Axis::Value(ValueAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        crosses: None,
    });
    let y = Axis::Value(ValueAxis {
        common: AxisCommon::new(100, 10, AxisPos::Left),
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        crosses: None,
    });
    (x, y)
}

fn series_with_cat_val() -> Series {
    let mut s = Series::new(0);
    s.title = Some(SeriesTitle::StrRef(Reference::new("Sheet", "B1")));
    s.categories = Some(Reference::new("Sheet", "A2:A6"));
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    s
}

// ---------------------------------------------------------------------------
// 8 chart-kind round-trips
// ---------------------------------------------------------------------------

#[test]
fn bar_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.title = Some(Title::plain("Sales"));
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:barChart>");
    assert_contains(&bytes, "<c:barDir val=\"col\"/>");
    assert_contains(&bytes, "<c:grouping val=\"clustered\"/>");
    assert_contains(&bytes, "<c:gapWidth val=\"150\"/>");
    assert_contains(&bytes, "<c:catAx>");
    assert_contains(&bytes, "<c:valAx>");
    assert_contains(&bytes, "<c:axId val=\"10\"/>");
    assert_contains(&bytes, "<c:axId val=\"100\"/>");
    assert_contains(&bytes, "'Sheet'!$A$2:$A$6");
    assert_contains(&bytes, "'Sheet'!$B$2:$B$6");
}

#[test]
fn line_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    c.title = Some(Title::plain("Trend"));
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:lineChart>");
    assert_contains(&bytes, "<c:grouping val=\"standard\"/>");
    // No barDir on line.
    assert_not_contains(&bytes, "<c:barDir");
}

#[test]
fn pie_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Pie, anchor_at(0, 0));
    c.vary_colors = Some(true);
    let mut s = Series::new(0);
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    s.categories = Some(Reference::new("Sheet", "A2:A6"));
    c.add_series(s);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:pieChart>");
    assert_contains(&bytes, "<c:varyColors val=\"1\"/>");
    // Pie has no axes.
    assert_not_contains(&bytes, "<c:catAx>");
    assert_not_contains(&bytes, "<c:valAx>");
    assert_not_contains(&bytes, "<c:axId");
}

#[test]
fn doughnut_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Doughnut, anchor_at(0, 0));
    c.hole_size = Some(75);
    c.vary_colors = Some(true);
    let mut s = Series::new(0);
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    c.add_series(s);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:doughnutChart>");
    assert_contains(&bytes, "<c:holeSize val=\"75\"/>");
    assert_not_contains(&bytes, "<c:catAx>");
}

#[test]
fn area_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Area, anchor_at(0, 0));
    c.grouping = Some(BarGrouping::PercentStacked);
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:areaChart>");
    assert_contains(&bytes, "<c:grouping val=\"percentStacked\"/>");
    assert_contains(&bytes, "<c:catAx>");
    assert_contains(&bytes, "<c:valAx>");
}

#[test]
fn scatter_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Scatter, anchor_at(0, 0));
    let mut s = Series::new(0);
    s.x_values = Some(Reference::new("Sheet", "A2:A6"));
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    c.add_series(s);
    let (x, y) = dual_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:scatterChart>");
    assert_contains(&bytes, "<c:scatterStyle val=\"lineMarker\"/>");
    assert_contains(&bytes, "<c:xVal>");
    assert_contains(&bytes, "<c:yVal>");
    assert_not_contains(&bytes, "<c:cat>");
    // Scatter uses two valAx, no catAx.
    assert_not_contains(&bytes, "<c:catAx>");
    let text = std::str::from_utf8(&bytes).unwrap();
    assert_eq!(text.matches("<c:valAx>").count(), 2, "two valAx for scatter");
}

#[test]
fn bubble_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Bubble, anchor_at(0, 0));
    c.bubble_scale = Some(120);
    c.show_neg_bubbles = Some(false);
    let mut s = Series::new(0);
    s.x_values = Some(Reference::new("Sheet", "A2:A6"));
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    s.bubble_size = Some(Reference::new("Sheet", "C2:C6"));
    c.add_series(s);
    let (x, y) = dual_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:bubbleChart>");
    assert_contains(&bytes, "<c:bubbleScale val=\"120\"/>");
    assert_contains(&bytes, "<c:bubbleSize>");
    assert_contains(&bytes, "<c:xVal>");
    assert_contains(&bytes, "<c:yVal>");
}

#[test]
fn radar_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Radar, anchor_at(0, 0));
    c.radar_style = Some(RadarStyle::Filled);
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:radarChart>");
    assert_contains(&bytes, "<c:radarStyle val=\"filled\"/>");
    assert_contains(&bytes, "<c:catAx>");
}

// ---------------------------------------------------------------------------
// 12 sub-feature tests
// ---------------------------------------------------------------------------

#[test]
fn category_axis_emits_lbl_offset() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Category(CategoryAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        lbl_offset: Some(75),
        lbl_algn: Some("ctr".to_string()),
    }));
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: AxisCommon::new(100, 10, AxisPos::Left),
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        crosses: None,
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:lblOffset val=\"75\"/>");
    assert_contains(&bytes, "<c:lblAlgn val=\"ctr\"/>");
}

#[test]
fn value_axis_emits_min_max_units() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Category(CategoryAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        lbl_offset: None,
        lbl_algn: None,
    }));
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: AxisCommon::new(100, 10, AxisPos::Left),
        min: Some(0.0),
        max: Some(100.0),
        major_unit: Some(20.0),
        minor_unit: Some(5.0),
        crosses: Some("autoZero".to_string()),
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:min val=\"0\"/>");
    assert_contains(&bytes, "<c:max val=\"100\"/>");
    assert_contains(&bytes, "<c:majorUnit val=\"20\"/>");
    assert_contains(&bytes, "<c:minorUnit val=\"5\"/>");
    assert_contains(&bytes, "<c:crosses val=\"autoZero\"/>");
}

#[test]
fn date_axis_emits_base_time_unit() {
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Date(DateAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        min: None,
        max: None,
        major_unit: Some(1.0),
        minor_unit: None,
        base_time_unit: Some("months".to_string()),
    }));
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: AxisCommon::new(100, 10, AxisPos::Left),
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        crosses: None,
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:dateAx>");
    assert_contains(&bytes, "<c:baseTimeUnit val=\"months\"/>");
    assert_contains(&bytes, "<c:majorUnit val=\"1\"/>");
}

#[test]
fn title_with_rich_text_runs() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
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
        crosses: None,
    }));
    c.title = Some(Title {
        runs: vec![
            TitleRun {
                text: "Q4 ".into(),
                bold: Some(true),
                italic: None,
                underline: None,
                size_pt: Some(14),
                color: Some("FF0000FF".into()),
                font_name: Some("Calibri".into()),
            },
            TitleRun {
                text: "Sales".into(),
                bold: None,
                italic: Some(true),
                underline: None,
                size_pt: Some(14),
                color: None,
                font_name: None,
            },
        ],
        overlay: Some(false),
        layout: None,
    });
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<a:t>Q4 </a:t>");
    assert_contains(&bytes, "<a:t>Sales</a:t>");
    // Bold on first run.
    assert_contains(&bytes, "b=\"1\"");
    // Italic on second run.
    assert_contains(&bytes, "i=\"1\"");
    // Color on first run (alpha stripped).
    assert_contains(&bytes, "<a:srgbClr val=\"0000FF\"/>");
    assert_contains(&bytes, "Calibri");
    // Size encoded as 14*100 = 1400.
    assert_contains(&bytes, "sz=\"1400\"");
}

#[test]
fn legend_each_position() {
    let positions = [
        (LegendPosition::Right, "r"),
        (LegendPosition::Left, "l"),
        (LegendPosition::Top, "t"),
        (LegendPosition::Bottom, "b"),
        (LegendPosition::TopRight, "tr"),
    ];
    for (pos, expected) in positions {
        let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
        c.add_series(series_with_cat_val());
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
            crosses: None,
        }));
        c.legend = Some(Legend {
            position: pos,
            overlay: None,
            layout: None,
        });
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        assert_contains(
            &bytes,
            &format!("<c:legendPos val=\"{expected}\"/>"),
        );
    }
}

#[test]
fn data_labels_show_val_cat_position() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    let mut s = series_with_cat_val();
    s.data_labels = Some(DataLabels {
        show_val: Some(true),
        show_cat_name: Some(true),
        show_ser_name: Some(false),
        show_percent: None,
        show_legend_key: None,
        show_bubble_size: None,
        position: Some("outEnd".to_string()),
        number_format: Some("0.00".to_string()),
        separator: Some(", ".to_string()),
    });
    c.add_series(s);
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:showVal val=\"1\"/>");
    assert_contains(&bytes, "<c:showCatName val=\"1\"/>");
    assert_contains(&bytes, "<c:showSerName val=\"0\"/>");
    assert_contains(&bytes, "<c:dLblPos val=\"outEnd\"/>");
    assert_contains(&bytes, "formatCode=\"0.00\"");
    assert_contains(&bytes, "<c:separator>, </c:separator>");
}

#[test]
fn error_bars_each_type_and_val_type() {
    let cases = [
        (ErrorBarType::Plus, ErrorBarValType::FixedVal, 1.5_f64),
        (ErrorBarType::Minus, ErrorBarValType::Percentage, 5.0),
        (ErrorBarType::Both, ErrorBarValType::StdDev, 2.0),
    ];
    for (bt, vt, value) in cases {
        let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
        let mut s = series_with_cat_val();
        s.error_bars.push(ErrorBars {
            bar_type: bt,
            val_type: vt,
            value: Some(value),
            no_end_cap: Some(false),
        });
        c.add_series(s);
        let (x, y) = cat_value_pair();
        c.x_axis = Some(x);
        c.y_axis = Some(y);
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        assert_contains(
            &bytes,
            &format!("<c:errBarType val=\"{}\"/>", bt.as_str()),
        );
        assert_contains(
            &bytes,
            &format!("<c:errValType val=\"{}\"/>", vt.as_str()),
        );
    }
}

#[test]
fn trendlines_linear_polynomial_movingavg() {
    let cases = [
        Trendline {
            kind: TrendlineKind::Linear,
            order: None,
            period: None,
            forward: Some(2.0),
            backward: Some(1.0),
            display_equation: Some(true),
            display_r_squared: Some(true),
            name: None,
        },
        Trendline {
            kind: TrendlineKind::Polynomial,
            order: Some(3),
            period: None,
            forward: None,
            backward: None,
            display_equation: None,
            display_r_squared: None,
            name: Some("Cubic".into()),
        },
        Trendline {
            kind: TrendlineKind::MovingAvg,
            order: None,
            period: Some(3),
            forward: None,
            backward: None,
            display_equation: None,
            display_r_squared: None,
            name: None,
        },
    ];
    for tl in cases {
        let expected_kind = tl.kind.as_str().to_string();
        let order = tl.order;
        let period = tl.period;
        let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
        let mut s = series_with_cat_val();
        s.trendlines.push(tl);
        c.add_series(s);
        let (x, y) = cat_value_pair();
        c.x_axis = Some(x);
        c.y_axis = Some(y);
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        assert_contains(
            &bytes,
            &format!("<c:trendlineType val=\"{expected_kind}\"/>"),
        );
        if let Some(o) = order {
            assert_contains(&bytes, &format!("<c:order val=\"{o}\"/>"));
        }
        if let Some(p) = period {
            assert_contains(&bytes, &format!("<c:period val=\"{p}\"/>"));
        }
    }
}

#[test]
fn manual_layout_emits_emu_style_fractions() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    c.layout = Some(Layout {
        x: 0.1,
        y: 0.2,
        w: 0.7,
        h: 0.6,
        layout_target: None,
    });
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:manualLayout>");
    assert_contains(&bytes, "<c:xMode val=\"edge\"/>");
    assert_contains(&bytes, "<c:x val=\"0.1\"/>");
    assert_contains(&bytes, "<c:y val=\"0.2\"/>");
    assert_contains(&bytes, "<c:w val=\"0.7\"/>");
    assert_contains(&bytes, "<c:h val=\"0.6\"/>");
}

#[test]
fn major_and_minor_gridlines_emitted() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Category(CategoryAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        lbl_offset: None,
        lbl_algn: None,
    }));
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: {
            let mut a = AxisCommon::new(100, 10, AxisPos::Left);
            a.major_gridlines = true;
            a.minor_gridlines = true;
            a
        },
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        crosses: None,
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:majorGridlines/>");
    assert_contains(&bytes, "<c:minorGridlines/>");
}

#[test]
fn graphical_properties_line_and_fill() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    let mut s = series_with_cat_val();
    s.graphical_properties = Some(GraphicalProperties {
        line_color: Some("FFFF0000".into()),
        line_width_emu: Some(12700),
        line_dash: Some("dash".into()),
        fill_color: Some("FF00FF00".into()),
        no_fill: false,
        no_line: false,
    });
    c.add_series(s);
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    // Fill color (alpha stripped).
    assert_contains(&bytes, "<a:srgbClr val=\"00FF00\"/>");
    // Line width.
    assert_contains(&bytes, "<a:ln w=\"12700\">");
    // Line color (alpha stripped).
    assert_contains(&bytes, "<a:srgbClr val=\"FF0000\"/>");
    // Dash style.
    assert_contains(&bytes, "<a:prstDash val=\"dash\"/>");
}

#[test]
fn marker_symbol_size_emitted() {
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    let mut s = Series::new(0);
    s.values = Some(Reference::new("S", "B2:B6"));
    s.marker = Some(Marker {
        symbol: MarkerSymbol::Triangle,
        size: Some(9),
        graphical_properties: None,
    });
    c.add_series(s);
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:symbol val=\"triangle\"/>");
    assert_contains(&bytes, "<c:size val=\"9\"/>");
}

#[test]
fn multi_series_chart_emits_distinct_idx_order() {
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    for i in 0..3 {
        let mut s = Series::new(i);
        s.title = Some(SeriesTitle::Literal(format!("Series {i}")));
        s.values = Some(Reference::new("Sheet", &format!("B{}:B{}", i + 2, i + 6)));
        s.categories = Some(Reference::new("Sheet", "A2:A6"));
        c.add_series(s);
    }
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    let text = std::str::from_utf8(&bytes).unwrap();
    // Three <c:ser> entries with distinct idx values.
    assert_eq!(text.matches("<c:ser>").count(), 3);
    assert_contains(&bytes, "<c:idx val=\"0\"/>");
    assert_contains(&bytes, "<c:idx val=\"1\"/>");
    assert_contains(&bytes, "<c:idx val=\"2\"/>");
    assert_contains(&bytes, "Series 0");
    assert_contains(&bytes, "Series 1");
    assert_contains(&bytes, "Series 2");
}

#[test]
fn vary_colors_true_for_pie() {
    let mut c = Chart::new(ChartKind::Pie, anchor_at(0, 0));
    c.vary_colors = Some(true);
    let mut s = Series::new(0);
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    c.add_series(s);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:varyColors val=\"1\"/>");
}

#[test]
fn series_axis_shape_emitted() {
    // Series axis is rare but present in 3D and surface charts.
    // Even for 2D kinds the shape has to round-trip if a user
    // attaches one.
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Series(SeriesAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
    }));
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: AxisCommon::new(100, 10, AxisPos::Left),
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        crosses: None,
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:serAx>");
}

#[test]
fn axis_orientation_and_tick_marks() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Category(CategoryAxis {
        common: {
            let mut a = AxisCommon::new(10, 100, AxisPos::Bottom);
            a.orientation = AxisOrientation::MaxMin;
            a.major_tick_mark = Some(TickMark::Out);
            a.minor_tick_mark = Some(TickMark::None);
            a
        },
        lbl_offset: None,
        lbl_algn: None,
    }));
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: AxisCommon::new(100, 10, AxisPos::Left),
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        crosses: None,
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:orientation val=\"maxMin\"/>");
    assert_contains(&bytes, "<c:majorTickMark val=\"out\"/>");
    assert_contains(&bytes, "<c:minorTickMark val=\"none\"/>");
}

#[test]
fn display_blanks_as_emitted() {
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    c.display_blanks_as = Some(DisplayBlanksAs::Span);
    c.plot_visible_only = Some(true);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:dispBlanksAs val=\"span\"/>");
    assert_contains(&bytes, "<c:plotVisOnly val=\"1\"/>");
}

#[test]
fn bar_overlap_emitted_when_set() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.overlap = Some(-25);
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:overlap val=\"-25\"/>");
}

#[test]
fn end_to_end_emits_chart_part_via_emit_xlsx() {
    use wolfxl_writer::emit_xlsx;
    use wolfxl_writer::model::Worksheet;
    use wolfxl_writer::Workbook;
    use std::io::Cursor;

    let mut wb = Workbook::new();
    let mut sheet = Worksheet::new("Sheet");
    let mut chart = Chart::new(ChartKind::Bar, anchor_at(3, 1));
    chart.title = Some(Title::plain("End-to-end"));
    chart.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    chart.x_axis = Some(x);
    chart.y_axis = Some(y);
    sheet.charts.push(chart);
    wb.add_sheet(sheet);

    let bytes = emit_xlsx(&mut wb);

    // Crack open the ZIP and verify the chart part exists.
    let archive = ::zip::ZipArchive::new(Cursor::new(bytes)).expect("open archive");
    let names: Vec<String> = archive.file_names().map(|s| s.to_string()).collect();
    assert!(
        names.iter().any(|n| n == "xl/charts/chart1.xml"),
        "archive missing chart part; entries:\n{names:#?}"
    );
    assert!(names.iter().any(|n| n == "xl/drawings/drawing1.xml"));
    assert!(names
        .iter()
        .any(|n| n == "xl/drawings/_rels/drawing1.xml.rels"));
}
