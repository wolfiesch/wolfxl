# Fixture Design

Fixtures are Excel-authored source files used as ground truth.

## Principles

1. Prefer real Excel-generated files for oracle behavior.
2. Keep fixtures small, focused, and feature-specific.
3. Store canonical fixtures in version control.
4. Separate ephemeral benchmark output from canonical fixtures.

## In the ExcelBench repository

Fixtures live in the [ExcelBench](https://github.com/SynthGL/ExcelBench) repository:

- Canonical xlsx fixtures: `fixtures/excel/`
- Canonical xls fixtures: `fixtures/excel_xls/`
- Throughput fixtures: `fixtures/throughput_xlsx/`

## Regeneration

From the ExcelBench repo root:

```bash
uv run excelbench generate --output fixtures/excel
uv run excelbench generate-xls --output fixtures/excel_xls
```

These commands require Excel for fixture generation steps that rely on xlwings/Excel behavior.
