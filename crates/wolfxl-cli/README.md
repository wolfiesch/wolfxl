# wolfxl-cli

[![crates.io](https://img.shields.io/crates/v/wolfxl-cli.svg)](https://crates.io/crates/wolfxl-cli)

Command-line previewer for Excel xlsx files. Installs the `wolfxl` binary.

```bash
cargo install wolfxl-cli
```

## Usage

```bash
wolfxl peek workbook.xlsx                   # styled box preview (default)
wolfxl peek workbook.xlsx -e text           # tab-separated for awk/cut
wolfxl peek workbook.xlsx -e csv            # RFC 4180 CSV
wolfxl peek workbook.xlsx -e json           # machine-readable JSON
wolfxl peek workbook.xlsx -n 20 -w 30       # 20 rows, 30-char column cap
wolfxl peek workbook.xlsx -s "Balance Sheet"
```

The `text`, `csv`, and `json` exporters target byte-identical output with
[`xleak 0.2.5`](https://crates.io/crates/xleak) so existing tooling pipelines
work unchanged. 77 of 81 in-repo fixtures × format combinations are byte-equal;
the four diffs are cases where wolfxl is more correct than xleak (RFC 4180
header quoting, calamine line-ending normalization).

The default `box` exporter is wolfxl-branded with `╔═╗` banner and `┌─┬─┐`
table borders.

## Built on

- [`wolfxl-core`](https://crates.io/crates/wolfxl-core) — pure-Rust xlsx
  reader.
- [`calamine-styles`](https://crates.io/crates/calamine-styles) — xlsx
  parser with style metadata.

## License

MIT
