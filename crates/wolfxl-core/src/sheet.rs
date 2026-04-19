use std::fs::File;
use std::io::BufReader;

use calamine_styles::{Data, Reader, Xlsx};
use chrono::{NaiveDate, NaiveDateTime, NaiveTime};

use crate::cell::{Cell, CellValue};
use crate::error::{Error, Result};

type XlsxReader = Xlsx<BufReader<File>>;

pub struct Sheet {
    pub name: String,
    rows: Vec<Vec<Cell>>,
}

impl Sheet {
    pub(crate) fn load(wb: &mut XlsxReader, name: &str) -> Result<Self> {
        let value_range = wb
            .worksheet_range(name)
            .map_err(|e| Error::Xlsx(format!("read range for {name:?}: {e}")))?;

        let style_range = wb.worksheet_style(name).ok();

        let (h, w) = value_range.get_size();
        let mut rows: Vec<Vec<Cell>> = Vec::with_capacity(h);
        for r in 0..h {
            let mut row: Vec<Cell> = Vec::with_capacity(w);
            for c in 0..w {
                let value = value_range
                    .get((r, c))
                    .map(data_to_cell_value)
                    .unwrap_or(CellValue::Empty);
                let number_format = style_range
                    .as_ref()
                    .and_then(|sr| sr.get((r, c)))
                    .and_then(extract_number_format);
                row.push(Cell {
                    value,
                    number_format,
                });
            }
            rows.push(row);
        }

        Ok(Self {
            name: name.to_string(),
            rows,
        })
    }

    pub fn dimensions(&self) -> (usize, usize) {
        let h = self.rows.len();
        let w = self.rows.first().map(|r| r.len()).unwrap_or(0);
        (h, w)
    }

    pub fn rows(&self) -> &[Vec<Cell>] {
        &self.rows
    }

    pub fn row(&self, idx: usize) -> Option<&[Cell]> {
        self.rows.get(idx).map(|r| r.as_slice())
    }

    /// First row stringified - the conventional "header" row for table-shaped
    /// sheets. Empty cells become empty strings so position is preserved.
    pub fn headers(&self) -> Vec<String> {
        self.rows
            .first()
            .map(|row| {
                row.iter()
                    .map(|c| match &c.value {
                        CellValue::String(s) => s.clone(),
                        CellValue::Empty => String::new(),
                        other => format_value_plain(other),
                    })
                    .collect()
            })
            .unwrap_or_default()
    }
}

fn format_value_plain(v: &CellValue) -> String {
    match v {
        CellValue::Empty => String::new(),
        CellValue::String(s) => s.clone(),
        CellValue::Int(n) => n.to_string(),
        CellValue::Float(n) => n.to_string(),
        CellValue::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
        CellValue::Date(d) => d.format("%Y-%m-%d").to_string(),
        CellValue::DateTime(dt) => dt.format("%Y-%m-%d %H:%M:%S").to_string(),
        CellValue::Time(t) => t.format("%H:%M:%S").to_string(),
        CellValue::Error(e) => e.clone(),
    }
}

fn data_to_cell_value(d: &Data) -> CellValue {
    match d {
        Data::Empty => CellValue::Empty,
        Data::String(s) => CellValue::String(s.clone()),
        Data::Int(i) => CellValue::Int(*i),
        Data::Float(f) => {
            if f.fract() == 0.0 && f.abs() < (i64::MAX as f64) {
                CellValue::Int(*f as i64)
            } else {
                CellValue::Float(*f)
            }
        }
        Data::Bool(b) => CellValue::Bool(*b),
        Data::DateTime(dt) => excel_serial_to_datetime(dt.as_f64()),
        Data::DateTimeIso(s) => parse_iso_datetime_or_string(s),
        Data::DurationIso(s) => CellValue::String(s.clone()),
        Data::Error(e) => CellValue::Error(format!("{e:?}")),
        Data::RichText(rt) => CellValue::String(rt.plain_text().to_string()),
    }
}

/// Excel serial date → chrono. Sub-day fractions become Time; ≥1.0 with no
/// fractional part becomes Date; otherwise DateTime.
fn excel_serial_to_datetime(serial: f64) -> CellValue {
    if serial < 1.0 && serial > 0.0 {
        let secs = (serial * 86_400.0).round() as u32;
        let h = secs / 3600;
        let m = (secs % 3600) / 60;
        let s = secs % 60;
        return NaiveTime::from_hms_opt(h, m, s)
            .map(CellValue::Time)
            .unwrap_or_else(|| CellValue::Float(serial));
    }
    let mut days = serial.trunc() as i64;
    let frac = serial - (days as f64);
    // Excel's 1900 leap-year bug: serial 60 maps to the non-existent
    // 1900-02-29. openpyxl uses base 1899-12-30 (instead of 1899-12-31) to
    // dodge the bug for serials >= 60, but that leaves serials 1..59 off by
    // one day. The +1 correction restores serial 1 -> 1900-01-01 etc., which
    // matches openpyxl.utils.datetime.from_excel.
    if serial > 0.0 && serial < 60.0 {
        days += 1;
    }
    let base = NaiveDate::from_ymd_opt(1899, 12, 30).expect("static date");
    let date = if days >= 0 {
        base.checked_add_days(chrono::Days::new(days as u64))
    } else {
        // u64 cast on negative i64 wraps; subtract instead.
        base.checked_sub_days(chrono::Days::new((-days) as u64))
    }
    .unwrap_or(base);
    if frac.abs() < f64::EPSILON {
        return CellValue::Date(date);
    }
    let secs = (frac * 86_400.0).round() as u32;
    let h = secs / 3600;
    let m = (secs % 3600) / 60;
    let s = secs % 60;
    let time = NaiveTime::from_hms_opt(h.min(23), m.min(59), s.min(59))
        .unwrap_or_else(|| NaiveTime::from_hms_opt(0, 0, 0).unwrap());
    CellValue::DateTime(NaiveDateTime::new(date, time))
}

fn parse_iso_datetime_or_string(s: &str) -> CellValue {
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S%.f") {
        return CellValue::DateTime(dt);
    }
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S") {
        return CellValue::DateTime(dt);
    }
    if let Ok(d) = NaiveDate::parse_from_str(s, "%Y-%m-%d") {
        return CellValue::Date(d);
    }
    CellValue::String(s.to_string())
}

/// Pull the resolved format-code string off a calamine-styles `Style`. The
/// upstream crate handles built-in vs custom (>=164) resolution; we just
/// normalize the result and skip the no-op "General".
fn extract_number_format(style: &calamine_styles::Style) -> Option<String> {
    let nf = style.get_number_format()?;
    let code = nf.format_code.trim();
    if code.is_empty() || code.eq_ignore_ascii_case("general") {
        None
    } else {
        Some(code.to_string())
    }
}

#[cfg(test)]
mod tests {
    use chrono::Datelike;

    use super::*;

    fn date(value: CellValue) -> NaiveDate {
        match value {
            CellValue::Date(d) => d,
            other => panic!("expected Date, got {other:?}"),
        }
    }

    #[test]
    fn excel_serial_matches_openpyxl_for_pre_leap_serials() {
        // openpyxl maps serial 1 -> 1900-01-01 thanks to its +1 correction
        // for serials in (0, 60). Serial 59 -> 1900-02-28.
        assert_eq!(
            date(excel_serial_to_datetime(1.0)),
            NaiveDate::from_ymd_opt(1900, 1, 1).unwrap()
        );
        assert_eq!(
            date(excel_serial_to_datetime(59.0)),
            NaiveDate::from_ymd_opt(1900, 2, 28).unwrap()
        );
        // Serial 61 -> 1900-03-01 (Excel's fake serial-60 leap day is skipped).
        assert_eq!(
            date(excel_serial_to_datetime(61.0)),
            NaiveDate::from_ymd_opt(1900, 3, 1).unwrap()
        );
        // A modern serial: 44197 -> 2021-01-01.
        assert_eq!(
            date(excel_serial_to_datetime(44197.0)),
            NaiveDate::from_ymd_opt(2021, 1, 1).unwrap()
        );
    }

    #[test]
    fn excel_serial_negative_does_not_wrap() {
        // Bad/sentinel serials shouldn't panic or produce a date in the
        // far future via u64 wrap. Fall back to the epoch.
        let value = excel_serial_to_datetime(-100.0);
        let d = date(value);
        assert!(d.year() < 1900, "got {d}");
    }
}
