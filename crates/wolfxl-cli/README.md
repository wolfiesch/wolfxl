# wolfxl-cli

[![crates.io](https://img.shields.io/crates/v/wolfxl-cli.svg)](https://crates.io/crates/wolfxl-cli)

Command-line previewer for Excel xlsx files. Installs the `wolfxl` binary.

```bash
cargo install wolfxl-cli
```

## Subcommands

- `wolfxl peek <file>` - styled preview or text/csv/json export
- `wolfxl map <file>` - one-page workbook overview (sheets, dims, headers)
- `wolfxl agent <file> --max-tokens N` - token-budgeted workbook briefing
- `wolfxl schema <file>` - per-column type, cardinality, and format inference

## Usage

```bash
wolfxl peek workbook.xlsx                   # styled box preview (default)
wolfxl peek workbook.xlsx -e text           # tab-separated for awk/cut
wolfxl peek workbook.xlsx -e csv            # RFC 4180 CSV
wolfxl peek workbook.xlsx -e json           # machine-readable JSON
wolfxl peek workbook.xlsx -n 20 -w 30       # 20 rows, 30-char column cap
wolfxl peek workbook.xlsx -s "Balance Sheet"
wolfxl map workbook.xlsx                    # workbook inventory for agents
wolfxl map workbook.xlsx --format text      # terminal-friendly summary
wolfxl agent workbook.xlsx --max-tokens 400 # fit a briefing to a token budget
wolfxl schema workbook.xlsx                 # JSON schema-style column summary
wolfxl schema workbook.xlsx -s "P&L" -f text
```

The `text`, `csv`, and `json` exporters are tuned for piping into LLM /
agent contexts and Unix tooling: integer thousand-grouping, two-decimal
floats, ISO dates, RFC 4180 CSV quoting (including embedded `\r`/`\n`),
and stable JSON shape (`{sheet, rows, columns, headers, data}`).

The default `box` exporter is wolfxl-branded with `╔═╗` banner and `┌─┬─┐`
table borders.

`map` emits workbook-level metadata for planning a next step: sheet names,
dimensions, detected sheet class, headers, and named ranges. `agent` uses the
same core metadata plus a stratified row sample to compose a briefing that fits
within a `cl100k_base` token budget. `schema` emits per-column type,
cardinality, null-count, format-category, and sample-value inference for one
sheet or the whole workbook.

## Built on

- [`wolfxl-core`](https://crates.io/crates/wolfxl-core) — pure-Rust xlsx
  reader.
- [`calamine-styles`](https://crates.io/crates/calamine-styles) — xlsx
  parser with style metadata.

## License

MIT
