#!/usr/bin/env python3
"""
csv_to_excel.py — Convert CSV file(s) to Excel (.xlsx)

Usage:
  python csv_to_excel.py input.csv output.xlsx
  python csv_to_excel.py input1.csv input2.csv output.xlsx
  python csv_to_excel.py *.csv output.xlsx
  python csv_to_excel.py input.csv output.xlsx --encoding utf-8-sig --delimiter ";"

Features:
  - Multiple CSV → multiple sheets in one .xlsx
  - Auto-detect delimiter (comma, semicolon, tab)
  - Type inference (int, float, bool, date)
  - Styled header row (bold + blue background)
  - Auto-fit column widths
  - Freeze header row + autofilter
  - Handles UTF-8, UTF-8-sig (BOM), cp1252, etc.
"""
import argparse
import csv
import datetime
import os
import sys
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Error: openpyxl not installed. Run: pip install openpyxl", file=sys.stderr)
    sys.exit(1)


def detect_delimiter(filepath, encoding='utf-8'):
    """Auto-detect delimiter by reading first few lines."""
    with open(filepath, 'r', encoding=encoding, newline='') as f:
        sample = ''.join(f.readline() for _ in range(5))
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=',;\t|')
        return dialect.delimiter
    except csv.Error:
        # Fallback: count common delimiters
        counts = {d: sample.count(d) for d in [',', ';', '\t', '|']}
        return max(counts, key=counts.get) if any(counts.values()) else ','


def infer_type(value):
    """Infer Python type from string value."""
    if value is None or value == '':
        return None
    
    # Boolean
    low = value.strip().lower()
    if low in ('true', 'yes', 'ya', '1'):
        if low in ('true', 'yes', 'ya'):
            return True
    if low in ('false', 'no', 'tidak', '0'):
        if low in ('false', 'no', 'tidak'):
            return False
    
    # Integer
    try:
        return int(value.replace(',', '').replace('.', ''))
    except (ValueError, TypeError):
        pass
    
    # Float (handle comma as decimal separator)
    cleaned = value.replace(',', '.').replace(' ', '')
    try:
        return float(cleaned)
    except (ValueError, TypeError):
        pass
    
    # Date formats: YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d'):
        try:
            return datetime.datetime.strptime(value.strip(), fmt).date()
        except (ValueError, TypeError):
            pass
    
    # Datetime
    for fmt in ('%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M:%S', '%Y-%m-%dT%H:%M:%S'):
        try:
            return datetime.datetime.strptime(value.strip(), fmt)
        except (ValueError, TypeError):
            pass
    
    return value


def sheet_name_from_filename(filepath):
    """Generate sheet name from filename (max 31 chars, Excel limit)."""
    name = Path(filepath).stem
    # Excel sheet name restrictions
    for char in '[]:*?/\\':
        name = name.replace(char, '_')
    return name[:31]


def convert_csv_to_sheet(wb, csv_path, sheet_name=None, encoding='utf-8', delimiter=None):
    """Convert a single CSV file to a sheet in the workbook."""
    if sheet_name is None:
        sheet_name = sheet_name_from_filename(csv_path)
    
    # Handle duplicate sheet names
    base_name = sheet_name
    counter = 1
    while sheet_name in wb.sheetnames:
        suffix = f'_{counter}'
        sheet_name = f'{base_name[:31-len(suffix)]}{suffix}'
        counter += 1
    
    ws = wb.create_sheet(title=sheet_name)
    
    # Detect encoding if not specified
    if encoding == 'auto':
        for enc in ['utf-8-sig', 'utf-8', 'cp1252', 'latin-1']:
            try:
                with open(csv_path, 'r', encoding=enc) as test:
                    test.read(1024)
                encoding = enc
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        else:
            encoding = 'latin-1'  # Fallback that never fails
    
    # Detect delimiter
    if delimiter is None:
        delimiter = detect_delimiter(csv_path, encoding)
    
    # Read CSV
    with open(csv_path, 'r', encoding=encoding, newline='') as f:
        reader = csv.reader(f, delimiter=delimiter)
        rows = list(reader)
    
    if not rows:
        print(f"  WARNING: {csv_path} is empty, skipping", file=sys.stderr)
        return 0
    
    # Header style
    header_font = Font(bold=True, color='FFFFFF', size=11)
    header_fill = PatternFill(start_color='2563EB', end_color='2563EB', fill_type='solid')
    header_align = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        bottom=Side(style='thin', color='CCCCCC')
    )
    
    # Write header
    headers = rows[0]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    
    # Write data rows with type inference
    data_count = 0
    for row_idx, row in enumerate(rows[1:], 2):
        for col_idx, value in enumerate(row, 1):
            typed_value = infer_type(value)
            cell = ws.cell(row=row_idx, column=col_idx, value=typed_value)
            # Format dates
            if isinstance(typed_value, (datetime.date, datetime.datetime)):
                cell.number_format = 'YYYY-MM-DD' if isinstance(typed_value, datetime.date) and not isinstance(typed_value, datetime.datetime) else 'YYYY-MM-DD HH:MM:SS'
        data_count += 1
    
    # Auto-fit column widths
    for col_idx in range(1, len(headers) + 1):
        max_length = len(str(headers[col_idx - 1])) if col_idx <= len(headers) else 0
        for row_idx in range(2, min(data_count + 2, 100)):  # Sample first 100 rows
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is not None:
                max_length = max(max_length, len(str(cell_value)))
        # Set width with some padding, cap at 50
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 3, 50)
    
    # Freeze header row
    ws.freeze_panes = 'A2'
    
    # Add autofilter
    if len(headers) > 0 and data_count > 0:
        last_col = get_column_letter(len(headers))
        ws.auto_filter.ref = f'A1:{last_col}{data_count + 1}'
    
    return data_count


def main():
    parser = argparse.ArgumentParser(
        description='Convert CSV file(s) to Excel (.xlsx) with styling',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s data.csv output.xlsx
  %(prog)s sales.csv customers.csv report.xlsx
  %(prog)s *.csv combined.xlsx
  %(prog)s data.csv out.xlsx --encoding utf-8-sig --delimiter ";"
        """
    )
    parser.add_argument('files', nargs='+', help='CSV file(s) + output .xlsx (last argument)')
    parser.add_argument('--encoding', default='auto', help='CSV encoding (auto, utf-8, utf-8-sig, cp1252). Default: auto')
    parser.add_argument('--delimiter', default=None, help='CSV delimiter (auto-detect by default). E.g. "," ";" "\\t"')
    
    args = parser.parse_args()
    
    # Last file is the output
    *csv_files, output_file = args.files
    
    if not csv_files:
        parser.error('Need at least one CSV file and one output .xlsx file')
    
    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'
    
    # Validate CSV files exist
    for csv_file in csv_files:
        if not os.path.exists(csv_file):
            print(f"Error: File not found: {csv_file}", file=sys.stderr)
            sys.exit(1)
    
    # Handle delimiter escape
    delimiter = args.delimiter
    if delimiter == '\\t':
        delimiter = '\t'
    
    # Create workbook
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    
    print(f"Converting {len(csv_files)} CSV file(s) to {output_file}")
    print(f"Encoding: {args.encoding}, Delimiter: {delimiter or 'auto-detect'}")
    print("-" * 60)
    
    total_rows = 0
    for csv_file in csv_files:
        sheet_name = sheet_name_from_filename(csv_file)
        rows = convert_csv_to_sheet(wb, csv_file, sheet_name, args.encoding, delimiter)
        total_rows += rows
        print(f"  {csv_file} → sheet '{sheet_name}' ({rows} rows)")
    
    # Save
    wb.save(output_file)
    print("-" * 60)
    print(f"Done! {output_file} ({len(csv_files)} sheets, {total_rows} total rows)")
    print(f"Size: {os.path.getsize(output_file):,} bytes")


if __name__ == '__main__':
    main()
