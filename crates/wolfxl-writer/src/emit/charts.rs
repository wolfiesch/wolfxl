//! `xl/charts/chartN.xml` emitter — Sprint Μ Pod-α (RFC-046 §4).
//!
//! One `xl/charts/chartN.xml` per [`Chart`] anchored on a sheet. The
//! emitter produces a `<c:chartSpace>` element with the `c:` (chart) and
//! `a:` (drawingml) namespaces declared on the root, matching openpyxl's
//! emit shape so byte-parity tests can pass downstream (Pod-δ).
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

use crate::model::chart::{
    Axis, AxisCommon, BarDir, BarGrouping, CategoryAxis, Chart, ChartKind, DataLabels, DateAxis,
    ErrorBars, GraphicalProperties, Gridlines, Layout, Legend, Marker, PivotSource, RadarStyle,
    Reference, ScatterStyle, Series, SeriesAxis, SeriesTitle, Title, TitleRun, Trendline,
    TrendlineKind, ValueAxis, View3D,
};
use crate::xml_escape;

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

    // Sprint Ν Pod-δ — RFC-049 §10. ``<c:pivotSource>`` MUST be the
    // first child of `<c:chart>` per ECMA-376 Part 1 §21.2.2.27 (the
    // chart sequence is `pivotSource? title? autoTitleDeleted?
    // view3D? plotArea legend? plotVisOnly? dispBlanksAs? extLst?`).
    if let Some(ps) = &chart.pivot_source {
        emit_pivot_source(&mut out, ps);
    }

    // Optional <c:title>
    if let Some(t) = &chart.title {
        emit_title(&mut out, t);
    }

    // Auto-title deleted is implicit when title is absent + we emit no element.

    // Optional <c:view3D> for 3D variants. Sprint Μ-prime / RFC-046 §10.10.
    if chart.kind.is_3d() {
        if let Some(v) = &chart.view_3d {
            emit_view_3d(&mut out, v);
        }
    }

    // <c:plotArea>
    out.push_str("<c:plotArea>");
    if let Some(layout) = &chart.layout {
        emit_layout(&mut out, layout);
    } else {
        // openpyxl always emits an empty <layout/> inside plotArea.
        out.push_str("<c:layout/>");
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

fn pick_axis_ids(chart: &Chart) -> (u32, u32) {
    // Prefer ids from the explicit axis blocks if set, so callers who
    // pre-assign axis ids see them honored. Otherwise fall back to
    // openpyxl's familiar 10 / 100 pair.
    let from_explicit = (
        chart.x_axis.as_ref().map(axis_common).map(|c| c.ax_id),
        chart.y_axis.as_ref().map(axis_common).map(|c| c.ax_id),
    );
    match from_explicit {
        (Some(a), Some(b)) => (a, b),
        _ => (10, 100),
    }
}

fn axis_common(a: &Axis) -> &AxisCommon {
    match a {
        Axis::Category(c) => &c.common,
        Axis::Value(c) => &c.common,
        Axis::Date(c) => &c.common,
        Axis::Series(c) => &c.common,
    }
}

fn emit_plot_chart(out: &mut String, chart: &Chart, ax_a: u32, ax_b: u32) {
    let elem = chart.kind.plot_element_name();
    out.push_str(&format!("<c:{elem}>"));

    // varyColors comes first in plot-area children (openpyxl order).
    if let Some(v) = chart.vary_colors {
        out.push_str(&format!("<c:varyColors val=\"{}\"/>", bool_str(v)));
    }

    // Type-specific shape header.
    match chart.kind {
        ChartKind::Bar | ChartKind::Bar3D => {
            if let Some(d) = chart.bar_dir {
                out.push_str(&format!("<c:barDir val=\"{}\"/>", d.as_str()));
            } else {
                out.push_str(&format!("<c:barDir val=\"{}\"/>", BarDir::Col.as_str()));
            }
            if let Some(g) = chart.grouping {
                out.push_str(&format!("<c:grouping val=\"{}\"/>", g.as_str()));
            } else {
                out.push_str(&format!(
                    "<c:grouping val=\"{}\"/>",
                    BarGrouping::Clustered.as_str()
                ));
            }
        }
        ChartKind::Line | ChartKind::Line3D => {
            // Default grouping for Line is "standard".
            let g = chart.grouping.unwrap_or(BarGrouping::Standard);
            out.push_str(&format!("<c:grouping val=\"{}\"/>", g.as_str()));
        }
        ChartKind::Area | ChartKind::Area3D => {
            let g = chart.grouping.unwrap_or(BarGrouping::Standard);
            out.push_str(&format!("<c:grouping val=\"{}\"/>", g.as_str()));
        }
        ChartKind::Scatter => {
            let s = chart.scatter_style.unwrap_or(ScatterStyle::LineMarker);
            out.push_str(&format!("<c:scatterStyle val=\"{}\"/>", s.as_str()));
        }
        ChartKind::Radar => {
            let s = chart.radar_style.unwrap_or(RadarStyle::Standard);
            out.push_str(&format!("<c:radarStyle val=\"{}\"/>", s.as_str()));
        }
        ChartKind::OfPie => {
            // ofPieType comes first inside <ofPieChart>.
            let t = chart.of_pie_type.as_deref().unwrap_or("pie");
            out.push_str(&format!("<c:ofPieType val=\"{}\"/>", xml_escape::attr(t)));
        }
        ChartKind::Surface | ChartKind::Surface3D => {
            if let Some(w) = chart.wireframe {
                out.push_str(&format!("<c:wireframe val=\"{}\"/>", bool_str(w)));
            }
        }
        // Pie/Doughnut/Pie3D/Bubble/Stock have no opening style attr beyond varyColors.
        _ => {}
    }

    // Series. Sprint Ν Pod-δ — when the chart has a `pivot_source`,
    // each `<c:ser>` MUST carry a `<c:fmtId val="0"/>` element matching
    // the pivot source's fmt_id; Excel rejects pivot charts whose
    // series lack `<c:fmtId>`. RFC-049 §2.
    let pivot_fmt_id = chart.pivot_source.as_ref().map(|ps| ps.fmt_id);
    for ser in &chart.series {
        emit_series(out, ser, chart.kind, pivot_fmt_id);
    }

    // Type-specific trailing properties.
    match chart.kind {
        ChartKind::Bar | ChartKind::Bar3D => {
            if let Some(g) = chart.gap_width {
                out.push_str(&format!("<c:gapWidth val=\"{g}\"/>"));
            }
            if let Some(o) = chart.overlap {
                out.push_str(&format!("<c:overlap val=\"{o}\"/>"));
            }
        }
        ChartKind::Doughnut => {
            if let Some(a) = chart.first_slice_ang {
                out.push_str(&format!("<c:firstSliceAng val=\"{a}\"/>"));
            }
            if let Some(h) = chart.hole_size {
                out.push_str(&format!("<c:holeSize val=\"{h}\"/>"));
            }
        }
        ChartKind::Pie | ChartKind::Pie3D => {
            if let Some(a) = chart.first_slice_ang {
                out.push_str(&format!("<c:firstSliceAng val=\"{a}\"/>"));
            }
        }
        ChartKind::OfPie => {
            // gapWidth, splitType, splitPos, secondPieSize.
            if let Some(g) = chart.gap_width {
                out.push_str(&format!("<c:gapWidth val=\"{g}\"/>"));
            }
            let st = chart.split_type.as_deref().unwrap_or("auto");
            out.push_str(&format!("<c:splitType val=\"{}\"/>", xml_escape::attr(st)));
            if let Some(p) = chart.split_pos {
                out.push_str(&format!("<c:splitPos val=\"{}\"/>", fmt_f64(p)));
            }
            if let Some(s) = chart.second_pie_size {
                out.push_str(&format!("<c:secondPieSize val=\"{s}\"/>"));
            }
        }
        ChartKind::Bubble => {
            if let Some(s) = chart.bubble_scale {
                out.push_str(&format!("<c:bubbleScale val=\"{s}\"/>"));
            }
            if let Some(b) = chart.show_neg_bubbles {
                out.push_str(&format!("<c:showNegBubbles val=\"{}\"/>", bool_str(b)));
            }
            if let Some(b) = chart.bubble3d {
                out.push_str(&format!("<c:bubble3D val=\"{}\"/>", bool_str(b)));
            }
        }
        ChartKind::Line | ChartKind::Line3D => {
            if let Some(s) = chart.smoothing {
                out.push_str(&format!("<c:smooth val=\"{}\"/>", bool_str(s)));
            }
        }
        ChartKind::Stock => {
            // Stock charts emit hiLowLines + upDownBars decorators.
            out.push_str("<c:hiLowLines/>");
            out.push_str("<c:upDownBars><c:gapWidth val=\"150\"/></c:upDownBars>");
        }
        _ => {}
    }

    // Axis ids — Pie/Doughnut emit none; everything else emits both.
    if !chart.kind.is_axis_free() {
        out.push_str(&format!("<c:axId val=\"{ax_a}\"/>"));
        out.push_str(&format!("<c:axId val=\"{ax_b}\"/>"));
    }

    out.push_str(&format!("</c:{elem}>"));
}

fn emit_series(out: &mut String, ser: &Series, kind: ChartKind, pivot_fmt_id: Option<u32>) {
    out.push_str("<c:ser>");
    out.push_str(&format!("<c:idx val=\"{}\"/>", ser.idx));
    out.push_str(&format!("<c:order val=\"{}\"/>", ser.order));

    // Sprint Ν Pod-δ — RFC-049 §2. Pivot-chart series carry `fmtId`
    // immediately after the order block. Excel rejects pivot charts
    // whose series lack this element.
    if let Some(fmt_id) = pivot_fmt_id {
        out.push_str(&format!("<c:fmtId val=\"{fmt_id}\"/>"));
    }

    if let Some(t) = &ser.title {
        emit_series_title(out, t);
    }

    if let Some(g) = &ser.graphical_properties {
        emit_graphical_props(out, g);
    }

    if let Some(m) = &ser.marker {
        emit_marker(out, m);
    }

    if let Some(b) = ser.invert_if_negative {
        out.push_str(&format!("<c:invertIfNegative val=\"{}\"/>", bool_str(b)));
    }

    if let Some(d) = &ser.data_labels {
        emit_data_labels(out, d);
    }

    for eb in &ser.error_bars {
        emit_error_bars(out, eb);
    }

    for tl in &ser.trendlines {
        emit_trendline(out, tl);
    }

    // x/y/cat/val/bubbleSize depending on chart kind.
    match kind {
        ChartKind::Scatter => {
            if let Some(x) = &ser.x_values {
                out.push_str("<c:xVal>");
                emit_num_ref(out, x);
                out.push_str("</c:xVal>");
            }
            if let Some(y) = &ser.values {
                out.push_str("<c:yVal>");
                emit_num_ref(out, y);
                out.push_str("</c:yVal>");
            }
            if let Some(s) = ser.smooth {
                out.push_str(&format!("<c:smooth val=\"{}\"/>", bool_str(s)));
            }
        }
        ChartKind::Bubble => {
            if let Some(x) = &ser.x_values {
                out.push_str("<c:xVal>");
                emit_num_ref(out, x);
                out.push_str("</c:xVal>");
            }
            if let Some(y) = &ser.values {
                out.push_str("<c:yVal>");
                emit_num_ref(out, y);
                out.push_str("</c:yVal>");
            }
            if let Some(b) = &ser.bubble_size {
                out.push_str("<c:bubbleSize>");
                emit_num_ref(out, b);
                out.push_str("</c:bubbleSize>");
            }
        }
        _ => {
            if let Some(c) = &ser.categories {
                out.push_str("<c:cat>");
                emit_num_ref(out, c);
                out.push_str("</c:cat>");
            }
            if let Some(v) = &ser.values {
                out.push_str("<c:val>");
                emit_num_ref(out, v);
                out.push_str("</c:val>");
            }
            if matches!(kind, ChartKind::Line | ChartKind::Radar | ChartKind::Line3D) {
                if let Some(s) = ser.smooth {
                    out.push_str(&format!("<c:smooth val=\"{}\"/>", bool_str(s)));
                }
            }
        }
    }

    out.push_str("</c:ser>");
}

fn emit_series_title(out: &mut String, title: &SeriesTitle) {
    out.push_str("<c:tx>");
    match title {
        SeriesTitle::StrRef(r) => {
            out.push_str("<c:strRef>");
            out.push_str(&format!(
                "<c:f>{}</c:f>",
                xml_escape::text(&r.to_formula_string())
            ));
            out.push_str("</c:strRef>");
        }
        SeriesTitle::Literal(s) => {
            out.push_str("<c:v>");
            out.push_str(&xml_escape::text(s));
            out.push_str("</c:v>");
        }
    }
    out.push_str("</c:tx>");
}

fn emit_num_ref(out: &mut String, r: &Reference) {
    out.push_str("<c:numRef>");
    out.push_str(&format!(
        "<c:f>{}</c:f>",
        xml_escape::text(&r.to_formula_string())
    ));
    out.push_str("</c:numRef>");
}

/// Sprint Ν Pod-δ — RFC-049 §10.1. Emits the chart-level
/// `<c:pivotSource><c:name>…</c:name><c:fmtId val="…"/></c:pivotSource>`
/// block. `name` is XML-escaped (per the §2 OOXML spec, the inner
/// element is text-content, not an attribute).
fn emit_pivot_source(out: &mut String, ps: &PivotSource) {
    out.push_str("<c:pivotSource>");
    out.push_str("<c:name>");
    out.push_str(&xml_escape::text(&ps.name));
    out.push_str("</c:name>");
    out.push_str(&format!("<c:fmtId val=\"{}\"/>", ps.fmt_id));
    out.push_str("</c:pivotSource>");
}

fn emit_title(out: &mut String, t: &Title) {
    out.push_str("<c:title>");
    out.push_str("<c:tx>");
    out.push_str("<c:rich>");
    out.push_str("<a:bodyPr/>");
    out.push_str("<a:lstStyle/>");
    out.push_str("<a:p>");
    out.push_str("<a:pPr><a:defRPr/></a:pPr>");
    for run in &t.runs {
        emit_run(out, run);
    }
    out.push_str("</a:p>");
    out.push_str("</c:rich>");
    out.push_str("</c:tx>");
    if let Some(layout) = &t.layout {
        emit_layout(out, layout);
    }
    if let Some(o) = t.overlay {
        out.push_str(&format!("<c:overlay val=\"{}\"/>", bool_str(o)));
    }
    out.push_str("</c:title>");
}

fn emit_run(out: &mut String, run: &TitleRun) {
    out.push_str("<a:r>");
    // <a:rPr> contains run-level formatting.
    let mut rpr = String::new();
    if let Some(sz) = run.size_pt {
        // Excel encodes pt as 100 * pt.
        rpr.push_str(&format!(" sz=\"{}\"", sz * 100));
    }
    if let Some(b) = run.bold {
        rpr.push_str(&format!(" b=\"{}\"", bool_str(b)));
    }
    if let Some(i) = run.italic {
        rpr.push_str(&format!(" i=\"{}\"", bool_str(i)));
    }
    if let Some(u) = run.underline {
        rpr.push_str(if u { " u=\"sng\"" } else { " u=\"none\"" });
    }
    out.push_str(&format!("<a:rPr lang=\"en-US\"{rpr}>"));
    if let Some(c) = &run.color {
        out.push_str(&format!(
            "<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>",
            strip_alpha(c)
        ));
    }
    if let Some(f) = &run.font_name {
        out.push_str(&format!("<a:latin typeface=\"{}\"/>", xml_escape::attr(f)));
    }
    out.push_str("</a:rPr>");
    out.push_str("<a:t>");
    out.push_str(&xml_escape::text(&run.text));
    out.push_str("</a:t>");
    out.push_str("</a:r>");
}

fn emit_legend(out: &mut String, l: &Legend) {
    out.push_str("<c:legend>");
    out.push_str(&format!("<c:legendPos val=\"{}\"/>", l.position.as_str()));
    if let Some(layout) = &l.layout {
        emit_layout(out, layout);
    }
    if let Some(o) = l.overlay {
        out.push_str(&format!("<c:overlay val=\"{}\"/>", bool_str(o)));
    }
    out.push_str("</c:legend>");
}

fn emit_layout(out: &mut String, layout: &Layout) {
    out.push_str("<c:layout>");
    out.push_str("<c:manualLayout>");
    if let Some(t) = layout.layout_target {
        out.push_str(&format!("<c:layoutTarget val=\"{}\"/>", t.as_str()));
    }
    out.push_str("<c:xMode val=\"edge\"/>");
    out.push_str("<c:yMode val=\"edge\"/>");
    out.push_str(&format!("<c:x val=\"{}\"/>", fmt_f64(layout.x)));
    out.push_str(&format!("<c:y val=\"{}\"/>", fmt_f64(layout.y)));
    out.push_str(&format!("<c:w val=\"{}\"/>", fmt_f64(layout.w)));
    out.push_str(&format!("<c:h val=\"{}\"/>", fmt_f64(layout.h)));
    out.push_str("</c:manualLayout>");
    out.push_str("</c:layout>");
}

/// Emit a `<c:majorGridlines/>` or `<c:minorGridlines/>` element.
///
/// `tag` is the bare element name (without prefix). Sprint Μ-prime
/// (RFC-046 §10.7.1). The bool flag is the legacy short-form flag; the
/// `obj` (when present) takes precedence and may carry graphical
/// properties.
fn emit_gridlines(out: &mut String, tag: &str, flag: bool, obj: Option<&Gridlines>) {
    if let Some(g) = obj {
        if let Some(gp) = &g.graphical_properties {
            out.push_str(&format!("<c:{tag}>"));
            emit_graphical_props(out, gp);
            out.push_str(&format!("</c:{tag}>"));
        } else {
            out.push_str(&format!("<c:{tag}/>"));
        }
        return;
    }
    if flag {
        out.push_str(&format!("<c:{tag}/>"));
    }
}

/// Emit `<c:view3D>` (chart-level, before plotArea). RFC-046 §10.10.
fn emit_view_3d(out: &mut String, v: &View3D) {
    out.push_str("<c:view3D>");
    if let Some(rx) = v.rot_x {
        out.push_str(&format!("<c:rotX val=\"{rx}\"/>"));
    }
    if let Some(ry) = v.rot_y {
        out.push_str(&format!("<c:rotY val=\"{ry}\"/>"));
    }
    if let Some(p) = v.perspective {
        out.push_str(&format!("<c:perspective val=\"{p}\"/>"));
    }
    if let Some(b) = v.right_angle_axes {
        out.push_str(&format!("<c:rAngAx val=\"{}\"/>", bool_str(b)));
    }
    if let Some(b) = v.auto_scale {
        out.push_str(&format!("<c:autoScale val=\"{}\"/>", bool_str(b)));
    }
    if let Some(d) = v.depth_percent {
        out.push_str(&format!("<c:depthPercent val=\"{d}\"/>"));
    }
    if let Some(h) = v.h_percent {
        out.push_str(&format!("<c:hPercent val=\"{h}\"/>"));
    }
    out.push_str("</c:view3D>");
}

fn emit_data_labels(out: &mut String, d: &DataLabels) {
    out.push_str("<c:dLbls>");
    if let Some(nf) = &d.number_format {
        out.push_str(&format!(
            "<c:numFmt formatCode=\"{}\" sourceLinked=\"0\"/>",
            xml_escape::attr(nf)
        ));
    }
    if let Some(p) = &d.position {
        out.push_str(&format!("<c:dLblPos val=\"{}\"/>", xml_escape::attr(p)));
    }
    if let Some(b) = d.show_legend_key {
        out.push_str(&format!("<c:showLegendKey val=\"{}\"/>", bool_str(b)));
    }
    if let Some(b) = d.show_val {
        out.push_str(&format!("<c:showVal val=\"{}\"/>", bool_str(b)));
    }
    if let Some(b) = d.show_cat_name {
        out.push_str(&format!("<c:showCatName val=\"{}\"/>", bool_str(b)));
    }
    if let Some(b) = d.show_ser_name {
        out.push_str(&format!("<c:showSerName val=\"{}\"/>", bool_str(b)));
    }
    if let Some(b) = d.show_percent {
        out.push_str(&format!("<c:showPercent val=\"{}\"/>", bool_str(b)));
    }
    if let Some(b) = d.show_bubble_size {
        out.push_str(&format!("<c:showBubbleSize val=\"{}\"/>", bool_str(b)));
    }
    if let Some(s) = &d.separator {
        out.push_str(&format!(
            "<c:separator>{}</c:separator>",
            xml_escape::text(s)
        ));
    }
    out.push_str("</c:dLbls>");
}

fn emit_error_bars(out: &mut String, eb: &ErrorBars) {
    out.push_str("<c:errBars>");
    out.push_str(&format!("<c:errBarType val=\"{}\"/>", eb.bar_type.as_str()));
    out.push_str(&format!("<c:errValType val=\"{}\"/>", eb.val_type.as_str()));
    if let Some(b) = eb.no_end_cap {
        out.push_str(&format!("<c:noEndCap val=\"{}\"/>", bool_str(b)));
    }
    if let Some(v) = eb.value {
        out.push_str(&format!("<c:val val=\"{}\"/>", fmt_f64(v)));
    }
    out.push_str("</c:errBars>");
}

fn emit_trendline(out: &mut String, tl: &Trendline) {
    out.push_str("<c:trendline>");
    if let Some(name) = &tl.name {
        out.push_str(&format!("<c:name>{}</c:name>", xml_escape::text(name)));
    }
    out.push_str(&format!("<c:trendlineType val=\"{}\"/>", tl.kind.as_str()));
    if let Some(o) = tl.order {
        if matches!(tl.kind, TrendlineKind::Polynomial) {
            out.push_str(&format!("<c:order val=\"{o}\"/>"));
        }
    }
    if let Some(p) = tl.period {
        if matches!(tl.kind, TrendlineKind::MovingAvg) {
            out.push_str(&format!("<c:period val=\"{p}\"/>"));
        }
    }
    if let Some(f) = tl.forward {
        out.push_str(&format!("<c:forward val=\"{}\"/>", fmt_f64(f)));
    }
    if let Some(b) = tl.backward {
        out.push_str(&format!("<c:backward val=\"{}\"/>", fmt_f64(b)));
    }
    if let Some(b) = tl.display_equation {
        out.push_str(&format!("<c:dispEq val=\"{}\"/>", bool_str(b)));
    }
    if let Some(b) = tl.display_r_squared {
        out.push_str(&format!("<c:dispRSqr val=\"{}\"/>", bool_str(b)));
    }
    out.push_str("</c:trendline>");
}

fn emit_marker(out: &mut String, m: &Marker) {
    out.push_str("<c:marker>");
    out.push_str(&format!("<c:symbol val=\"{}\"/>", m.symbol.as_str()));
    if let Some(s) = m.size {
        out.push_str(&format!("<c:size val=\"{s}\"/>"));
    }
    if let Some(g) = &m.graphical_properties {
        emit_graphical_props(out, g);
    }
    out.push_str("</c:marker>");
}

fn emit_graphical_props(out: &mut String, g: &GraphicalProperties) {
    out.push_str("<c:spPr>");
    // Fill.
    if g.no_fill {
        out.push_str("<a:noFill/>");
    } else if let Some(c) = &g.fill_color {
        out.push_str(&format!(
            "<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>",
            strip_alpha(c)
        ));
    }
    // Line.
    let has_line_attrs =
        g.line_color.is_some() || g.line_width_emu.is_some() || g.line_dash.is_some() || g.no_line;
    if has_line_attrs {
        if let Some(w) = g.line_width_emu {
            out.push_str(&format!("<a:ln w=\"{w}\">"));
        } else {
            out.push_str("<a:ln>");
        }
        if g.no_line {
            out.push_str("<a:noFill/>");
        } else if let Some(c) = &g.line_color {
            out.push_str(&format!(
                "<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>",
                strip_alpha(c)
            ));
        }
        if let Some(d) = &g.line_dash {
            out.push_str(&format!("<a:prstDash val=\"{}\"/>", xml_escape::attr(d)));
        }
        out.push_str("</a:ln>");
    }
    out.push_str("</c:spPr>");
}

fn emit_axis(out: &mut String, axis: &Axis) {
    match axis {
        Axis::Category(c) => emit_category_axis(out, c),
        Axis::Value(v) => emit_value_axis(out, v),
        Axis::Date(d) => emit_date_axis(out, d),
        Axis::Series(s) => emit_series_axis(out, s),
    }
}

fn emit_axis_common_pre(out: &mut String, common: &AxisCommon) {
    out.push_str(&format!("<c:axId val=\"{}\"/>", common.ax_id));
    out.push_str("<c:scaling>");
    out.push_str(&format!(
        "<c:orientation val=\"{}\"/>",
        common.orientation.as_str()
    ));
    out.push_str("</c:scaling>");
    if let Some(d) = common.delete {
        out.push_str(&format!("<c:delete val=\"{}\"/>", bool_str(d)));
    }
    out.push_str(&format!("<c:axPos val=\"{}\"/>", common.ax_pos.as_str()));
    emit_gridlines(
        out,
        "majorGridlines",
        common.major_gridlines,
        common.major_gridlines_obj.as_ref(),
    );
    emit_gridlines(
        out,
        "minorGridlines",
        common.minor_gridlines,
        common.minor_gridlines_obj.as_ref(),
    );
    if let Some(t) = &common.title {
        emit_title(out, t);
    }
    if let Some(nf) = &common.number_format {
        out.push_str(&format!(
            "<c:numFmt formatCode=\"{}\" sourceLinked=\"0\"/>",
            xml_escape::attr(nf)
        ));
    }
    if let Some(t) = common.major_tick_mark {
        out.push_str(&format!("<c:majorTickMark val=\"{}\"/>", t.as_str()));
    }
    if let Some(t) = common.minor_tick_mark {
        out.push_str(&format!("<c:minorTickMark val=\"{}\"/>", t.as_str()));
    }
    out.push_str(&format!("<c:crossAx val=\"{}\"/>", common.cross_ax));
}

fn emit_category_axis(out: &mut String, c: &CategoryAxis) {
    out.push_str("<c:catAx>");
    emit_axis_common_pre(out, &c.common);
    if let Some(o) = c.lbl_offset {
        out.push_str(&format!("<c:lblOffset val=\"{o}\"/>"));
    }
    if let Some(a) = &c.lbl_algn {
        out.push_str(&format!("<c:lblAlgn val=\"{}\"/>", xml_escape::attr(a)));
    }
    out.push_str("</c:catAx>");
}

fn emit_value_axis(out: &mut String, v: &ValueAxis) {
    out.push_str("<c:valAx>");
    // <scaling> wraps min/max/orientation; we re-emit it because the
    // shared writer above only handled orientation. Reopen the scaling
    // subtree by doing a simple custom emit.
    out.push_str(&format!("<c:axId val=\"{}\"/>", v.common.ax_id));
    out.push_str("<c:scaling>");
    out.push_str(&format!(
        "<c:orientation val=\"{}\"/>",
        v.common.orientation.as_str()
    ));
    if let Some(m) = v.min {
        out.push_str(&format!("<c:min val=\"{}\"/>", fmt_f64(m)));
    }
    if let Some(m) = v.max {
        out.push_str(&format!("<c:max val=\"{}\"/>", fmt_f64(m)));
    }
    out.push_str("</c:scaling>");
    if let Some(d) = v.common.delete {
        out.push_str(&format!("<c:delete val=\"{}\"/>", bool_str(d)));
    }
    out.push_str(&format!("<c:axPos val=\"{}\"/>", v.common.ax_pos.as_str()));
    emit_gridlines(
        out,
        "majorGridlines",
        v.common.major_gridlines,
        v.common.major_gridlines_obj.as_ref(),
    );
    emit_gridlines(
        out,
        "minorGridlines",
        v.common.minor_gridlines,
        v.common.minor_gridlines_obj.as_ref(),
    );
    if let Some(t) = &v.common.title {
        emit_title(out, t);
    }
    if let Some(nf) = &v.common.number_format {
        out.push_str(&format!(
            "<c:numFmt formatCode=\"{}\" sourceLinked=\"0\"/>",
            xml_escape::attr(nf)
        ));
    }
    if let Some(t) = v.common.major_tick_mark {
        out.push_str(&format!("<c:majorTickMark val=\"{}\"/>", t.as_str()));
    }
    if let Some(t) = v.common.minor_tick_mark {
        out.push_str(&format!("<c:minorTickMark val=\"{}\"/>", t.as_str()));
    }
    out.push_str(&format!("<c:crossAx val=\"{}\"/>", v.common.cross_ax));
    if let Some(c) = &v.crosses {
        out.push_str(&format!("<c:crosses val=\"{}\"/>", xml_escape::attr(c)));
    }
    if let Some(u) = v.major_unit {
        out.push_str(&format!("<c:majorUnit val=\"{}\"/>", fmt_f64(u)));
    }
    if let Some(u) = v.minor_unit {
        out.push_str(&format!("<c:minorUnit val=\"{}\"/>", fmt_f64(u)));
    }
    out.push_str("</c:valAx>");
}

fn emit_date_axis(out: &mut String, d: &DateAxis) {
    out.push_str("<c:dateAx>");
    out.push_str(&format!("<c:axId val=\"{}\"/>", d.common.ax_id));
    out.push_str("<c:scaling>");
    out.push_str(&format!(
        "<c:orientation val=\"{}\"/>",
        d.common.orientation.as_str()
    ));
    if let Some(m) = d.min {
        out.push_str(&format!("<c:min val=\"{}\"/>", fmt_f64(m)));
    }
    if let Some(m) = d.max {
        out.push_str(&format!("<c:max val=\"{}\"/>", fmt_f64(m)));
    }
    out.push_str("</c:scaling>");
    if let Some(de) = d.common.delete {
        out.push_str(&format!("<c:delete val=\"{}\"/>", bool_str(de)));
    }
    out.push_str(&format!("<c:axPos val=\"{}\"/>", d.common.ax_pos.as_str()));
    emit_gridlines(
        out,
        "majorGridlines",
        d.common.major_gridlines,
        d.common.major_gridlines_obj.as_ref(),
    );
    emit_gridlines(
        out,
        "minorGridlines",
        d.common.minor_gridlines,
        d.common.minor_gridlines_obj.as_ref(),
    );
    if let Some(t) = &d.common.title {
        emit_title(out, t);
    }
    if let Some(nf) = &d.common.number_format {
        out.push_str(&format!(
            "<c:numFmt formatCode=\"{}\" sourceLinked=\"0\"/>",
            xml_escape::attr(nf)
        ));
    }
    if let Some(t) = d.common.major_tick_mark {
        out.push_str(&format!("<c:majorTickMark val=\"{}\"/>", t.as_str()));
    }
    if let Some(t) = d.common.minor_tick_mark {
        out.push_str(&format!("<c:minorTickMark val=\"{}\"/>", t.as_str()));
    }
    out.push_str(&format!("<c:crossAx val=\"{}\"/>", d.common.cross_ax));
    if let Some(b) = &d.base_time_unit {
        out.push_str(&format!(
            "<c:baseTimeUnit val=\"{}\"/>",
            xml_escape::attr(b)
        ));
    }
    if let Some(u) = d.major_unit {
        out.push_str(&format!("<c:majorUnit val=\"{}\"/>", fmt_f64(u)));
    }
    if let Some(u) = d.minor_unit {
        out.push_str(&format!("<c:minorUnit val=\"{}\"/>", fmt_f64(u)));
    }
    out.push_str("</c:dateAx>");
}

fn emit_series_axis(out: &mut String, s: &SeriesAxis) {
    out.push_str("<c:serAx>");
    emit_axis_common_pre(out, &s.common);
    out.push_str("</c:serAx>");
}

fn bool_str(b: bool) -> &'static str {
    if b {
        "1"
    } else {
        "0"
    }
}

/// Strip the leading alpha from an 8-char ARGB color, leaving the 6-char
/// RGB. Drawingml's `<a:srgbClr val>` expects RGB, not ARGB.
fn strip_alpha(c: &str) -> String {
    if c.len() == 8 {
        c[2..].to_string()
    } else {
        c.to_string()
    }
}

/// Format f64 deterministically for OOXML — drop trailing zeros so
/// `1.0` becomes `"1"`, but keep precision for fractional values.
fn fmt_f64(v: f64) -> String {
    if v == v.trunc() && v.abs() < 1e16 {
        format!("{}", v as i64)
    } else {
        // Use `{}` Rust default; for now this is good enough.
        format!("{v}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::chart::{
        Axis, AxisCommon, AxisPos, BarDir, BarGrouping, CategoryAxis, Chart, ChartKind, DataLabels,
        ErrorBarType, ErrorBarValType, ErrorBars, Legend, LegendPosition, Marker, MarkerSymbol,
        Reference, ScatterStyle, Series, SeriesTitle, Title, Trendline, TrendlineKind, ValueAxis,
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
            crosses: None,
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
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<c:scatterChart>"));
        assert!(text.contains("<c:scatterStyle val=\"lineMarker\"/>"));
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
            crosses: None,
        }));
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<c:smooth val=\"1\"/>"));
    }

    // ----------------------------------------------------------------
    // Sprint Ν Pod-δ — pivot-chart linkage (RFC-049)
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
        //    series-order block. RFC-049 §2 — Excel rejects pivot
        //    charts whose series lack `<c:fmtId>`.
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
