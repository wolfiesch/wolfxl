"""``openpyxl.styles.numbers`` — number-format helpers + the builtin catalog.

Wolfxl's ``is_date_format`` lives at :mod:`wolfxl.utils.numbers`; this module
re-exports it under the openpyxl-shaped path and bundles the canonical
``BUILTIN_FORMATS`` mapping (numFmtId → format string) verbatim from
openpyxl 3.1.x's catalog.

Pod 2 (RFC-060).
"""

from __future__ import annotations

from wolfxl.utils.numbers import is_date_format

# openpyxl 3.1.x ``openpyxl/styles/numbers.py`` — frozen dict of every
# builtin (Excel-reserved) numFmtId → format-string mapping.  Indexes 0..49
# come from the spec; the rest are openpyxl additions for the formats Excel
# emits in practice.
BUILTIN_FORMATS = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    41: r'_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)',
    42: r'_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)',
    43: r'_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)',
    44: r'_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}

# Per-purpose convenience constants — openpyxl exposes these as the
# canonical strings users assign to ``cell.number_format``.
FORMAT_GENERAL = BUILTIN_FORMATS[0]
FORMAT_TEXT = BUILTIN_FORMATS[49]
FORMAT_NUMBER = BUILTIN_FORMATS[1]
FORMAT_NUMBER_00 = BUILTIN_FORMATS[2]
FORMAT_NUMBER_COMMA_SEPARATED1 = BUILTIN_FORMATS[4]
FORMAT_PERCENTAGE = BUILTIN_FORMATS[9]
FORMAT_PERCENTAGE_00 = BUILTIN_FORMATS[10]
FORMAT_DATE_YYYYMMDD2 = "yyyy-mm-dd"
FORMAT_DATE_YYMMDD = "yy-mm-dd"
FORMAT_DATE_DDMMYY = "dd/mm/yy"
FORMAT_DATE_DMYSLASH = "d/m/y"
FORMAT_DATE_DMYMINUS = "d-m-y"
FORMAT_DATE_DMMINUS = "d-m"
FORMAT_DATE_MYMINUS = "m-y"
FORMAT_DATE_XLSX14 = BUILTIN_FORMATS[14]
FORMAT_DATE_XLSX15 = BUILTIN_FORMATS[15]
FORMAT_DATE_XLSX16 = BUILTIN_FORMATS[16]
FORMAT_DATE_XLSX17 = BUILTIN_FORMATS[17]
FORMAT_DATE_XLSX22 = BUILTIN_FORMATS[22]
FORMAT_DATE_DATETIME = "yyyy-mm-dd h:mm:ss"
FORMAT_DATE_TIME1 = BUILTIN_FORMATS[18]
FORMAT_DATE_TIME2 = BUILTIN_FORMATS[19]
FORMAT_DATE_TIME3 = BUILTIN_FORMATS[20]
FORMAT_DATE_TIME4 = BUILTIN_FORMATS[21]
FORMAT_DATE_TIME5 = BUILTIN_FORMATS[45]
FORMAT_DATE_TIME6 = BUILTIN_FORMATS[21]
FORMAT_CURRENCY_USD_SIMPLE = '"$"#,##0.00_-'
FORMAT_CURRENCY_USD = '$#,##0_-'
FORMAT_CURRENCY_EUR_SIMPLE = '[$EUR ]#,##0.00_-'


__all__ = [
    "BUILTIN_FORMATS",
    "FORMAT_CURRENCY_EUR_SIMPLE",
    "FORMAT_CURRENCY_USD",
    "FORMAT_CURRENCY_USD_SIMPLE",
    "FORMAT_DATE_DATETIME",
    "FORMAT_DATE_DDMMYY",
    "FORMAT_DATE_DMMINUS",
    "FORMAT_DATE_DMYMINUS",
    "FORMAT_DATE_DMYSLASH",
    "FORMAT_DATE_MYMINUS",
    "FORMAT_DATE_TIME1",
    "FORMAT_DATE_TIME2",
    "FORMAT_DATE_TIME3",
    "FORMAT_DATE_TIME4",
    "FORMAT_DATE_TIME5",
    "FORMAT_DATE_TIME6",
    "FORMAT_DATE_XLSX14",
    "FORMAT_DATE_XLSX15",
    "FORMAT_DATE_XLSX16",
    "FORMAT_DATE_XLSX17",
    "FORMAT_DATE_XLSX22",
    "FORMAT_DATE_YYMMDD",
    "FORMAT_DATE_YYYYMMDD2",
    "FORMAT_GENERAL",
    "FORMAT_NUMBER",
    "FORMAT_NUMBER_00",
    "FORMAT_NUMBER_COMMA_SEPARATED1",
    "FORMAT_PERCENTAGE",
    "FORMAT_PERCENTAGE_00",
    "FORMAT_TEXT",
    "is_date_format",
]
