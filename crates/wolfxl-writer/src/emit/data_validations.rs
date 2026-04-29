//! `<dataValidations>` emitter for worksheet XML.

use crate::model::validation::{ErrorStyle, ValidationOperator, ValidationType};
use crate::model::worksheet::Worksheet;
use crate::xml_escape;

/// Emit `<dataValidations count="N">…</dataValidations>`.
///
/// The caller is responsible for invoking this at CT_Worksheet slot 18,
/// between conditional formatting and hyperlinks.
pub fn emit(out: &mut String, sheet: &Worksheet) {
    if sheet.validations.is_empty() {
        return;
    }

    out.push_str(&format!(
        "<dataValidations count=\"{}\">",
        sheet.validations.len()
    ));

    for dv in &sheet.validations {
        let type_str = match dv.validation_type {
            ValidationType::Any => "any",
            ValidationType::Whole => "whole",
            ValidationType::Decimal => "decimal",
            ValidationType::List => "list",
            ValidationType::Date => "date",
            ValidationType::Time => "time",
            ValidationType::TextLength => "textLength",
            ValidationType::Custom => "custom",
        };

        out.push_str(&format!("<dataValidation type=\"{}\"", type_str));

        let needs_operator = !matches!(
            dv.validation_type,
            ValidationType::List | ValidationType::Custom
        );
        if needs_operator {
            let op_str = match dv.operator {
                ValidationOperator::Between => "between",
                ValidationOperator::NotBetween => "notBetween",
                ValidationOperator::Equal => "equal",
                ValidationOperator::NotEqual => "notEqual",
                ValidationOperator::GreaterThan => "greaterThan",
                ValidationOperator::LessThan => "lessThan",
                ValidationOperator::GreaterThanOrEqual => "greaterThanOrEqual",
                ValidationOperator::LessThanOrEqual => "lessThanOrEqual",
            };
            out.push_str(&format!(" operator=\"{}\"", op_str));
        }

        if dv.allow_blank {
            out.push_str(" allowBlank=\"1\"");
        }
        if dv.show_dropdown {
            out.push_str(" showDropDown=\"1\"");
        }
        if dv.show_input_message {
            out.push_str(" showInputMessage=\"1\"");
        }
        if dv.show_error_message {
            out.push_str(" showErrorMessage=\"1\"");
        }

        match dv.error_style {
            ErrorStyle::Stop => {}
            ErrorStyle::Warning => {
                out.push_str(" errorStyle=\"warning\"");
            }
            ErrorStyle::Information => {
                out.push_str(" errorStyle=\"information\"");
            }
        }

        if let Some(ref title) = dv.error_title {
            out.push_str(&format!(" errorTitle=\"{}\"", xml_escape::attr(title)));
        }
        if let Some(ref msg) = dv.error_message {
            out.push_str(&format!(" error=\"{}\"", xml_escape::attr(msg)));
        }
        if let Some(ref title) = dv.input_title {
            out.push_str(&format!(" promptTitle=\"{}\"", xml_escape::attr(title)));
        }
        if let Some(ref msg) = dv.input_message {
            out.push_str(&format!(" prompt=\"{}\"", xml_escape::attr(msg)));
        }

        out.push_str(&format!(" sqref=\"{}\">", xml_escape::attr(&dv.sqref)));

        if let Some(ref formula) = dv.formula_a {
            out.push_str(&format!(
                "<formula1>{}</formula1>",
                xml_escape::text(formula)
            ));
        }

        let is_between = matches!(
            dv.operator,
            ValidationOperator::Between | ValidationOperator::NotBetween
        );
        if is_between {
            if let Some(ref formula) = dv.formula_b {
                out.push_str(&format!(
                    "<formula2>{}</formula2>",
                    xml_escape::text(formula)
                ));
            }
        }

        out.push_str("</dataValidation>");
    }

    out.push_str("</dataValidations>");
}
