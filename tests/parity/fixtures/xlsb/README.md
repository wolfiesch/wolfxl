# `.xlsb` parity fixtures (Sprint Κ Pod-γ)

Five committed `.xlsb` files used by `tests/parity/test_xlsb_reads.py` to
assert wolfxl reads match `pandas.read_excel(engine="calamine")`.

| File              | Source                          | Notes                              |
|-------------------|---------------------------------|------------------------------------|
| `numbers.xlsb`    | calamine `tests/issue_419.xlsb` | numeric values                     |
| `strings.xlsb`    | calamine `tests/issue_186.xlsb` | string content                     |
| `dates.xlsb`      | calamine `tests/date.xlsb`      | dates / datetimes / time-deltas    |
| `formulas.xlsb`   | calamine `tests/issue127.xlsb`  | multi-sheet, mixed types           |
| `multisheet.xlsb` | calamine `tests/any_sheets.xlsb`| visible / hidden / very-hidden     |

## License & attribution

These fixtures are vendored from upstream
[`tafia/calamine`](https://github.com/tafia/calamine), MIT License,
Copyright © 2016 Johann Tuffe. The MIT license text is reproduced upstream.
We keep them committed (~50 KB total) so CI runners don't need network
access to validate parity.

## Why vendored, not generated?

LibreOffice 26.x exposes filter names like `Calc MS Excel 2007 Binary` and
`Calc Office Open XML Binary`, but the actual export call returns
`Error Area:Io Class:Parameter Code:26`. Neither `openpyxl` nor `xlsxwriter`
write `.xlsb`. Producing a real `.xlsb` requires either Excel itself or a
commercial library (Aspose). Vendoring real-world `.xlsb` fixtures from
calamine's MIT-licensed corpus is the canonical workaround, and matches
the read-only nature of wolfxl's `.xlsb` support.

## Refreshing the corpus

```bash
git clone https://github.com/tafia/calamine.git /tmp/calamine
python3 scripts/sprint_kappa_build_fixtures.py /tmp/calamine
```

The script copies five upstream `tests/*.xlsb` files into this directory
under stable wolfxl-side names (see the table above).
