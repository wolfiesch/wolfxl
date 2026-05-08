# Public Sample Fidelity Audit

This note tracks one-off internet-sourced workbook audits that should remain
reproducible without vendoring third-party binaries into the repository.

## 2026-05-08 Microsoft Power BI Excel Samples

Source page: [Microsoft Learn: Power BI samples as Excel workbooks](https://learn.microsoft.com/sr-latn-rs/power-bi/create-reports/sample-datasets/)

The Microsoft Learn page links public sample `.xlsx` workbooks hosted in the
`microsoft/powerbi-desktop-samples` GitHub repository. The page's sample-use
notice limits use to internal reference purposes, so the workbooks were
downloaded to `/tmp` for audit evidence only and were not committed.

| Workbook | Source URL | Bytes | SHA-256 |
|---|---|---:|---|
| `customer-profitability-no-pv.xlsx` | `https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/powerbi-service-samples/Customer%20Profitability%20Sample-no-PV.xlsx` | 2,906,427 | `76f21c59d631e95bbad5489350695a46d903061aedb179e88b72f038772666d4` |
| `it-spend-analysis-no-pv.xlsx` | `https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/powerbi-service-samples/IT%20Spend%20Analysis%20Sample-no-PV.xlsx` | 1,408,921 | `2adaf60667b42610d0a7d4d8ef141ea815d649acf60ba6e4f9f27376845ad4d5` |

Feature radar classified both workbooks as Excel-authored PowerPivot/Data Model
files with custom XML, workbook connections, extension payloads, drawing/media
parts, printer settings, workbook calc/global state, and style/theme content.
`audit_ooxml_gap_radar.py` reported no unknown content types, unknown
relationships, unknown extension URIs, unknown part families, or app-unsupported
features for this mini corpus.

Evidence artifacts:

| Check | Artifact | Result |
|---|---|---|
| Gap radar | `/tmp/wolfxl-public-powerbi-samples-20260508` | clear, `fixture_count=2` |
| Package mutation sweep | `/tmp/wolfxl-public-powerbi-mutations-20260508/report.json` | `36` results, `0` failures, `24` passed, `12` expected drift |
| Excel source app smoke | `/tmp/wolfxl-excel-public-powerbi-source-smoke-20260508-clean/app-smoke-report.json` | `2` results, `0` failures |
| Excel mutation app smoke | `/tmp/wolfxl-excel-public-powerbi-all-mutations-smoke-20260508/app-smoke-report.json` | `36` results, `0` failures |

The first two-file source app-smoke attempt produced one transient `Book1`
active-workbook mismatch for `it-spend-analysis-no-pv.xlsx`; rerunning from a
clean Excel process and running the workbook alone both opened it under the
expected filename with no repair prompt. The clean rerun above is the retained
evidence artifact.
