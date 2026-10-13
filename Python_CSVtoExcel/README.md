# Python CSV to Excel Converter

Convert CSV file(s) to styled Excel (.xlsx) with type inference, auto-fit columns,
freeze header row, and autofilter.

## Requirements

```bash
pip install openpyxl
```

## Usage

```bash
# Single CSV
python csv_to_excel.py data.csv output.xlsx

# Multiple CSV → multiple sheets in one .xlsx
python csv_to_excel.py customers.csv sales.csv report.xlsx

# Wildcard
python csv_to_excel.py *.csv combined.xlsx

# Custom encoding + delimiter
python csv_to_excel.py data.csv out.xlsx --encoding utf-8-sig --delimiter ";"
```

## Features

- **Multiple CSV → multiple sheets** — each CSV becomes a separate sheet
- **Auto-detect delimiter** — comma, semicolon, tab, pipe
- **Auto-detect encoding** — UTF-8, UTF-8-sig (BOM), cp1252, latin-1
- **Type inference** — int, float, bool, date, datetime
- **Styled header** — bold white text on blue background
- **Auto-fit columns** — width based on content
- **Freeze header row** — row 1 frozen when scrolling
- **Autofilter** — filter dropdown on header row

## Options

| Option | Default | Description |
|---|---|---|
| `--encoding` | `auto` | CSV encoding: `auto`, `utf-8`, `utf-8-sig`, `cp1252`, `latin-1` |
| `--delimiter` | auto-detect | CSV delimiter: `,` `;` `\t` `\|` |

## License

MIT
