# Public Sample Fidelity Audit

This note tracks one-off internet-sourced workbook audits that should remain
reproducible without vendoring third-party binaries into the repository.

## 2026-05-08 Microsoft Power BI Excel Samples

Source page: [Microsoft Learn: Power BI samples as Excel workbooks](https://learn.microsoft.com/sr-latn-rs/power-bi/create-reports/sample-datasets/)

The Microsoft Learn page links public sample `.xlsx` workbooks hosted in the
`microsoft/powerbi-desktop-samples` GitHub repository. The page's sample-use
notice limits use to internal reference purposes, so the workbooks were
downloaded to `/tmp` for audit evidence only and were not committed. The full
no-PowerView `.xlsx` set was tested, not just a hand-picked subset.

| Workbook | Source URL | Bytes | SHA-256 |
|---|---|---:|---|
| `customer-profitability-no-pv.xlsx` | `https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/powerbi-service-samples/Customer%20Profitability%20Sample-no-PV.xlsx` | 2,906,427 | `76f21c59d631e95bbad5489350695a46d903061aedb179e88b72f038772666d4` |
| `human-resources-no-pv.xlsx` | `https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/powerbi-service-samples/Human%20Resources%20Sample-no-PV.xlsx` | 10,377,251 | `d837a4af057b450510ca3d1b00ce1ecc6ced32a18255c99e4b10b5aea4f1f1aa` |
| `it-spend-analysis-no-pv.xlsx` | `https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/powerbi-service-samples/IT%20Spend%20Analysis%20Sample-no-PV.xlsx` | 1,408,921 | `2adaf60667b42610d0a7d4d8ef141ea815d649acf60ba6e4f9f27376845ad4d5` |
| `opportunity-tracking-no-pv.xlsx` | `https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/powerbi-service-samples/Opportunity%20Tracking%20Sample%20no%20PV.xlsx` | 700,910 | `16e924bc6ca72b89e4b1dbb09516778cca6c8cce8652d3e79ed166313b62ec7c` |
| `procurement-analysis-no-pv.xlsx` | `https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/powerbi-service-samples/Procurement%20Analysis%20Sample-no-PV.xlsx` | 15,220,519 | `6192e15c4b75942bc663c49fa2c971edf2b868eb8f194b54fb113e47a9e32aa5` |
| `retail-analysis-no-pv.xlsx` | `https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/powerbi-service-samples/Retail%20Analysis%20Sample-no-PV.xlsx` | 13,509,178 | `82f07475d5980321a1ce20b495cbd2ac2d26fe1ec13bf02ff405d7c4f704217c` |
| `sales-and-marketing-no-pv.xlsx` | `https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/powerbi-service-samples/Sales%20and%20Marketing%20Sample-no-PV.xlsx` | 8,525,431 | `ce2ca572d76d2725cf65da3edb08486a268b3548ef7ebaf5348b00e415f4ebcc` |
| `supplier-quality-analysis-no-pv.xlsx` | `https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/powerbi-service-samples/Supplier%20Quality%20Analysis%20Sample-no-PV.xlsx` | 794,651 | `0bf6079ea5345f31ba186bf7cfdf8db79cd2f8b673a9065919615583803ec43a` |

Feature radar classified the set as Excel-authored PowerPivot/Data Model files
with custom XML, workbook connections, extension payloads, drawing/media parts,
printer settings, workbook calc/global state, and style/theme content. The
larger set also covers pivot caches, chart/chart-style parts, and sheet metadata.

This pass discovered `{4F2E5C28-24EA-4eb8-9CBF-B6C8F9C3D259}` in
`xl/pivotCache/pivotCacheDefinition*.xml` as an `x15:cachedUniqueNames`
pivot-cache extension. The gap radar now classifies that extension explicitly,
and the expanded corpus reports no unknown content types, unknown relationships,
unknown extension URIs, unknown part families, or app-unsupported features.

Evidence artifacts:

| Check | Artifact | Result |
|---|---|---|
| Gap radar | `/tmp/wolfxl-public-powerbi-expanded-20260508` | clear, `fixture_count=8` |
| Package mutation sweep | `/tmp/wolfxl-public-powerbi-expanded-mutations-20260508/report.json` | `144` results, `0` failures, `96` passed, `48` expected drift |
| Excel source app smoke | `/tmp/wolfxl-excel-public-powerbi-expanded-source-smoke-20260508/app-smoke-report.json` | `8` results, `0` failures |
| Excel mutation app smoke | `/tmp/wolfxl-excel-public-powerbi-expanded-all-mutations-smoke-20260508/app-smoke-report.json` | `144` results, `0` failures |

The first two-file source app-smoke attempt produced one transient `Book1`
active-workbook mismatch for `it-spend-analysis-no-pv.xlsx`; rerunning from a
clean Excel process and running the workbook alone both opened it under the
expected filename with no repair prompt. The clean rerun above is the retained
evidence artifact. The later eight-file source smoke also opened all files under
their expected filenames.

## 2026-05-08 Microsoft Power BI Desktop Repository Remainder

Repository tree: [microsoft/powerbi-desktop-samples](https://github.com/microsoft/powerbi-desktop-samples/tree/main)

After the eight Learn-linked service samples passed, the repository tree was
enumerated for every remaining `.xlsx` workbook. Two additional public
workbooks were downloaded to `/tmp` and audited without committing binaries.

| Workbook | Source URL | Bytes | SHA-256 |
|---|---|---:|---|
| `adventureworks-sales.xlsx` | `https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/AdventureWorks%20Sales%20Sample/AdventureWorks%20Sales.xlsx` | 14,322,931 | `76fe718fb7806bc06f96b07f8e6835af22c1667f59157225cae6163827856df8` |
| `customerfeedback.xlsx` | `https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/Monthly%20Desktop%20Blog%20Samples/2019/customerfeedback.xlsx` | 7,848,484 | `76ab40c7b772976117f963d79034096c8ff96429f244c1bc6162a590b306b319` |

`adventureworks-sales.xlsx` introduced external-data query table package
surfaces: `xl/queryTables/queryTable*.xml`, query-table content types, and
query-table relationships. The gap radar now classifies those as workbook
connection/query-table evidence rather than unknown OOXML.

This pass also exposed a real app-level fidelity gap that the package oracle
had not caught: renaming the first AdventureWorks sheet left a sheet-scoped,
hidden `ExternalData_6` defined name pointing at the old sheet title, and Excel
timed out while opening the mutated file. The rename path now retargets
sheet-scoped defined-name formulas for the renamed sheet while preserving the
existing defined-name attributes. The mutation runner classifies the resulting
workbook-global fingerprint change as expected rename drift.

Evidence artifacts:

| Check | Artifact | Result |
|---|---|---|
| Gap radar | `/tmp/wolfxl-public-microsoft-extra-samples-20260508` | clear, `fixture_count=2` |
| Excel source app smoke | `/tmp/wolfxl-excel-public-microsoft-extra-source-smoke-20260508/app-smoke-report.json` | `2` results, `0` failures |
| Package mutation sweep | `/tmp/wolfxl-public-microsoft-extra-mutations-retarget-all-20260508/report.json` | `36` results, `0` failures, `23` passed, `13` expected drift |
| AdventureWorks rename retarget check | `/tmp/wolfxl-adventureworks-rename-retarget-check-20260508/report.json` | `1` result, `0` failures, `1` expected workbook-global drift |
| AdventureWorks renamed Excel app smoke | `/tmp/wolfxl-excel-single-mutated-open-probe-retarget-20260508/app-smoke-report.json` | `1` result, `0` failures |
