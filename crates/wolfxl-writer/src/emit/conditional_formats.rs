//! `<conditionalFormatting>` emitter for worksheet XML.

use std::collections::BTreeSet;

use crate::model::conditional::{CellIsOperator, ConditionalKind, ConditionalThreshold};
use crate::model::worksheet::Worksheet;
use crate::xml_escape;

/// Emit all supported conditional-formatting blocks for a worksheet.
///
/// Unsupported/stub variants are intentionally skipped. Emitting guessed XML
/// for those variants causes Excel repair prompts, so all-stub blocks also
/// omit their `<conditionalFormatting>` wrapper.
pub fn emit(out: &mut String, sheet: &Worksheet) {
    if sheet.conditional_formats.is_empty() {
        return;
    }

    let mut dropped: BTreeSet<&'static str> = BTreeSet::new();

    for cf in &sheet.conditional_formats {
        let mut rules_buf = String::new();

        for (priority_0, rule) in cf.rules.iter().enumerate() {
            let priority = priority_0 + 1;

            match &rule.kind {
                ConditionalKind::CellIs {
                    operator,
                    formula_a,
                    formula_b,
                } => {
                    let op_str = match operator {
                        CellIsOperator::Equal => "equal",
                        CellIsOperator::NotEqual => "notEqual",
                        CellIsOperator::GreaterThan => "greaterThan",
                        CellIsOperator::GreaterThanOrEqual => "greaterThanOrEqual",
                        CellIsOperator::LessThan => "lessThan",
                        CellIsOperator::LessThanOrEqual => "lessThanOrEqual",
                        CellIsOperator::Between => "between",
                        CellIsOperator::NotBetween => "notBetween",
                    };

                    rules_buf.push_str(&format!(
                        "<cfRule type=\"cellIs\" priority=\"{}\" operator=\"{}\"",
                        priority, op_str
                    ));
                    if let Some(dxf_id) = rule.dxf_id {
                        rules_buf.push_str(&format!(" dxfId=\"{}\"", dxf_id));
                    }
                    if rule.stop_if_true {
                        rules_buf.push_str(" stopIfTrue=\"1\"");
                    }
                    rules_buf.push('>');
                    rules_buf.push_str(&format!(
                        "<formula>{}</formula>",
                        xml_escape::text(formula_a)
                    ));
                    if matches!(
                        operator,
                        CellIsOperator::Between | CellIsOperator::NotBetween
                    ) {
                        if let Some(fb) = formula_b {
                            rules_buf
                                .push_str(&format!("<formula>{}</formula>", xml_escape::text(fb)));
                        }
                    }
                    rules_buf.push_str("</cfRule>");
                }

                ConditionalKind::Expression { formula } => {
                    rules_buf.push_str(&format!(
                        "<cfRule type=\"expression\" priority=\"{}\"",
                        priority
                    ));
                    if let Some(dxf_id) = rule.dxf_id {
                        rules_buf.push_str(&format!(" dxfId=\"{}\"", dxf_id));
                    }
                    if rule.stop_if_true {
                        rules_buf.push_str(" stopIfTrue=\"1\"");
                    }
                    rules_buf.push('>');
                    rules_buf
                        .push_str(&format!("<formula>{}</formula>", xml_escape::text(formula)));
                    rules_buf.push_str("</cfRule>");
                }

                ConditionalKind::DataBar {
                    color_rgb,
                    min,
                    max,
                } => {
                    rules_buf.push_str(&format!(
                        "<cfRule type=\"dataBar\" priority=\"{}\">",
                        priority
                    ));
                    rules_buf.push_str("<dataBar>");
                    emit_cfvo(&mut rules_buf, min);
                    emit_cfvo(&mut rules_buf, max);
                    rules_buf.push_str(&format!("<color rgb=\"{}\"/>", color_rgb));
                    rules_buf.push_str("</dataBar>");
                    rules_buf.push_str("</cfRule>");
                }

                ConditionalKind::ColorScale { stops } => {
                    rules_buf.push_str(&format!(
                        "<cfRule type=\"colorScale\" priority=\"{}\">",
                        priority
                    ));
                    rules_buf.push_str("<colorScale>");
                    for stop in stops {
                        emit_cfvo(&mut rules_buf, &stop.threshold);
                    }
                    for stop in stops {
                        rules_buf.push_str(&format!("<color rgb=\"{}\"/>", stop.color_rgb));
                    }
                    rules_buf.push_str("</colorScale>");
                    rules_buf.push_str("</cfRule>");
                }

                ConditionalKind::ContainsText { .. } => {
                    dropped.insert("ContainsText");
                    continue;
                }
                ConditionalKind::NotContainsText { .. } => {
                    dropped.insert("NotContainsText");
                    continue;
                }
                ConditionalKind::BeginsWith { .. } => {
                    dropped.insert("BeginsWith");
                    continue;
                }
                ConditionalKind::EndsWith { .. } => {
                    dropped.insert("EndsWith");
                    continue;
                }
                ConditionalKind::Duplicate => {
                    dropped.insert("Duplicate");
                    continue;
                }
                ConditionalKind::Unique => {
                    dropped.insert("Unique");
                    continue;
                }
                ConditionalKind::Top10 { .. } => {
                    dropped.insert("Top10");
                    continue;
                }
                ConditionalKind::AboveAverage { .. } => {
                    dropped.insert("AboveAverage");
                    continue;
                }
                ConditionalKind::IconSet { .. } => {
                    dropped.insert("IconSet");
                    continue;
                }
            }
        }

        if !rules_buf.is_empty() {
            out.push_str(&format!(
                "<conditionalFormatting sqref=\"{}\">",
                xml_escape::attr(&cf.sqref)
            ));
            out.push_str(&rules_buf);
            out.push_str("</conditionalFormatting>");
        }
    }

    if !dropped.is_empty() {
        let names: Vec<&str> = dropped.iter().copied().collect();
        eprintln!(
            "wolfxl-writer: dropped {} conditional-format rule kind{} on sheet {:?} \
             (variants: {}). Wave 3 ships only CellIs/Expression/DataBar/ColorScale; \
             other kinds are pending a future CF expansion wave.",
            names.len(),
            if names.len() == 1 { "" } else { "s" },
            sheet.name,
            names.join(", "),
        );
    }
}

fn emit_cfvo(out: &mut String, threshold: &ConditionalThreshold) {
    match threshold {
        ConditionalThreshold::Min => {
            out.push_str("<cfvo type=\"min\"/>");
        }
        ConditionalThreshold::Max => {
            out.push_str("<cfvo type=\"max\"/>");
        }
        ConditionalThreshold::Number(x) => {
            out.push_str(&format!("<cfvo type=\"num\" val=\"{}\"/>", format_f64(*x)));
        }
        ConditionalThreshold::Percent(x) => {
            out.push_str(&format!(
                "<cfvo type=\"percent\" val=\"{}\"/>",
                format_f64(*x)
            ));
        }
        ConditionalThreshold::Percentile(x) => {
            out.push_str(&format!(
                "<cfvo type=\"percentile\" val=\"{}\"/>",
                format_f64(*x)
            ));
        }
        ConditionalThreshold::Formula(s) => {
            out.push_str(&format!(
                "<cfvo type=\"formula\" val=\"{}\"/>",
                xml_escape::attr(s)
            ));
        }
    }
}

fn format_f64(n: f64) -> String {
    if n == (n as i64) as f64 && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        format!("{}", n)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::conditional::{ConditionalFormat, ConditionalRule};

    fn rule(kind: ConditionalKind) -> ConditionalRule {
        ConditionalRule {
            kind,
            dxf_id: None,
            stop_if_true: false,
        }
    }

    #[test]
    fn cell_is_rule_emits_wrapper_and_formula() {
        let mut sheet = Worksheet::new("S");
        sheet.conditional_formats.push(ConditionalFormat {
            sqref: "A1:A3".into(),
            rules: vec![rule(ConditionalKind::CellIs {
                operator: CellIsOperator::GreaterThan,
                formula_a: "10".into(),
                formula_b: None,
            })],
        });
        let mut out = String::new();

        emit(&mut out, &sheet);

        assert!(out.contains("<conditionalFormatting sqref=\"A1:A3\">"));
        assert!(out.contains("type=\"cellIs\""));
        assert!(out.contains("<formula>10</formula>"));
    }

    #[test]
    fn all_stub_rules_emit_no_wrapper() {
        let mut sheet = Worksheet::new("S");
        sheet.conditional_formats.push(ConditionalFormat {
            sqref: "A1:A3".into(),
            rules: vec![rule(ConditionalKind::Duplicate)],
        });
        let mut out = String::new();

        emit(&mut out, &sheet);

        assert!(out.is_empty());
    }
}
