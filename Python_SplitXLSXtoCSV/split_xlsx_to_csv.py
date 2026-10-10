#!/usr/bin/env python3
"""
Split XLSX/XLS to CSV
======================
Tool untuk membaca file Excel (.xlsx atau .xls), memilih kolom yang
dibutuhkan, lalu split menjadi multiple file CSV berdasarkan jumlah
baris per file.

Cara pakai: python split_xlsx_to_csv.py

Author: Hermes Agent for Rholly
"""

import os
import sys
import csv
import zipfile
from datetime import datetime

# ─── Helper: warna CLI ───
class c:
    HEADER = '\033[95m'
    OK = '\033[92m'
    WARN = '\033[93m'
    ERR = '\033[91m'
    CYAN = '\033[96m'
    BOLD = '\033[1m'
    END = '\033[0m'


def banner():
    print(f"""
{c.CYAN}{c.BOLD}╔══════════════════════════════════════════╗
║   Split XLSX to CSV — Interactive Tool   ║
╚══════════════════════════════════════════╝{c.END}
""")


def ask_file_path() -> str:
    """Tanya path file ke user, validasi sampai benar."""
    while True:
        path = input(f"{c.BOLD}📂 Masukkan path file Excel (.xlsx/.xls): {c.END}").strip()
        if not path:
            print(f"{c.WARN}  ⚠ Path tidak boleh kosong.{c.END}")
            continue
        if not os.path.isfile(path):
            print(f"{c.ERR}  ❌ File tidak ditemukan: {path}{c.END}")
            continue
        if not path.lower().endswith(('.xlsx', '.xlsm', '.xls')):
            print(f"{c.WARN}  ⚠ File bukan .xlsx/.xls. Yakin lanjut? (y/n): {c.END}", end=' ')
            if input().strip().lower() != 'y':
                continue
        return os.path.abspath(path)


def detect_format(filepath: str) -> str:
    """Deteksi format file: 'xlsx' (zip-based) atau 'xls' (BIFF) atau 'unknown'."""
    # Cek apakah file adalah ZIP (format xlsx sebenarnya)
    if zipfile.is_zipfile(filepath):
        return 'xlsx'
    # Cek magic number untuk XLS lama (BIFF): D0 CF 11 E0
    try:
        with open(filepath, 'rb') as f:
            magic = f.read(8)
        if magic == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
            return 'xls'
    except Exception:
        pass
    return 'unknown'


def read_headers(filepath: str) -> tuple:
    """Baca file, auto-detect format, return (data_source, headers)."""
    fmt = detect_format(filepath)

    if fmt == 'xlsx':
        return read_xlsx(filepath)
    elif fmt == 'xls':
        print(f"{c.WARN}  ⚠ File terdeteksi sebagai format .xls lama (BIFF).{c.END}")
        print(f"{c.CYAN}  → Menggunakan xlrd untuk membaca...{c.END}")
        return read_xls(filepath)
    else:
        # Mungkin CSV/TSV yang di-rename, atau file corrupt
        print(f"{c.ERR}  ❌ Format file tidak dikenali.{c.END}")
        print(f"{c.WARN}  Tips:{c.END}")
        print(f"    1. Buka file di Excel → Save As → .xlsx (bukan .xls)")
        print(f"    2. Atau pastikan file tidak corrupt")
        sys.exit(1)


def read_xlsx(filepath: str) -> tuple:
    """Baca file .xlsx (format zip-based) menggunakan openpyxl."""
    try:
        from openpyxl import load_workbook
    except ImportError:
        print(f"{c.ERR}  ❌ Package 'openpyxl' belum terpasang.{c.END}")
        print(f"  Install: pip install openpyxl")
        sys.exit(1)

    wb = load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    print(f"\n{c.OK}  ✓ Sheet: {ws.title} | Dimensi: {ws.dimensions} | Format: XLSX{c.END}")

    headers = []
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        headers = [str(cell) if cell is not None else f"Col_{i+1}" for i, cell in enumerate(row)]
        break

    if not headers:
        print(f"{c.ERR}  ❌ File tidak punya header (kosong?).{c.END}")
        sys.exit(1)

    return wb, ws, headers, 'xlsx'


def read_xls(filepath: str) -> tuple:
    """Baca file .xls lama (BIFF) menggunakan xlrd."""
    try:
        import xlrd
    except ImportError:
        print(f"{c.ERR}  ❌ Package 'xlrd' belum terpasang.{c.END}")
        print(f"  Install: pip install xlrd==2.0.1")
        print(f"\n{c.WARN}  Atau alternatif: buka file di Excel → Save As → .xlsx{c.END}")
        sys.exit(1)

    wb = xlrd.open_workbook(filepath)
    ws = wb.sheet_by_index(0)
    print(f"\n{c.OK}  ✓ Sheet: {ws.name} | Baris: {ws.nrows} | Kolom: {ws.ncols} | Format: XLS{c.END}")

    headers = []
    if ws.nrows > 0:
        headers = [str(ws.cell_value(0, i)) if ws.cell_value(0, i) else f"Col_{i+1}"
                    for i in range(ws.ncols)]

    if not headers:
        print(f"{c.ERR}  ❌ File tidak punya header (kosong?).{c.END}")
        sys.exit(1)

    return wb, ws, headers, 'xls'


def get_row_count(ws, fmt) -> int:
    """Hitung total baris data berdasarkan format."""
    if fmt == 'xlsx':
        return ws.max_row - 1
    else:
        return ws.nrows - 1


def iterate_rows(ws, selected_indices, fmt):
    """Generator yang yield baris data sesuai format."""
    if fmt == 'xlsx':
        for row in ws.iter_rows(min_row=2, values_only=True):
            yield [str(row[i]) if i < len(row) and row[i] is not None else ''
                   for i in selected_indices]
    else:
        import xlrd
        for row_idx in range(1, ws.nrows):
            yield [str(ws.cell_value(row_idx, i)) if ws.cell_value(row_idx, i) != '' else ''
                   for i in selected_indices]


def show_columns(headers: list) -> list:
    """Tampilkan semua kolom, biarkan user pilih mana yang dipakai."""
    print(f"\n{c.BOLD}📋 Kolom yang tersedia:{c.END}\n")
    for i, h in enumerate(headers, 1):
        print(f"  {c.CYAN}{i:>3}.{c.END} {h}")
    print()

    while True:
        inp = input(
            f"{c.BOLD}Pilih kolom (pisahkan dengan koma, contoh: 1,3,5 "
            f"atau ketik 'all' untuk semua): {c.END}"
        ).strip()

        if inp.lower() == 'all':
            return list(range(len(headers)))

        try:
            indices = [int(x.strip()) - 1 for x in inp.split(',')]
            if all(0 <= i < len(headers) for i in indices):
                selected = [(i, headers[i]) for i in indices]
                print(f"\n{c.OK}  ✓ Kolom terpilih:{c.END}")
                for i, name in selected:
                    print(f"    {c.CYAN}{i+1}.{c.END} {name}")
                print()
                return indices
            else:
                print(f"{c.ERR}  ❌ Nomor kolom di luar range.{c.END}")
        except ValueError:
            print(f"{c.ERR}  ❌ Format salah. Gunakan angka dipisah koma, contoh: 1,3,5{c.END}")


def ask_split_size(total_rows: int) -> int:
    """Tanya berapa baris per file CSV."""
    print(f"{c.BOLD}📊 Total baris data: {total_rows}{c.END}\n")
    while True:
        inp = input(
            f"{c.BOLD}Split menjadi berapa baris per file? "
            f"(contoh: 1000, atau ketik 'all' untuk 1 file): {c.END}"
        ).strip()

        if inp.lower() == 'all':
            return total_rows

        try:
            size = int(inp)
            if size <= 0:
                print(f"{c.ERR}  ❌ Harus angka positif.{c.END}")
                continue
            files_count = (total_rows + size - 1) // size
            print(f"\n{c.CYAN}  → Akan menghasilkan {files_count} file CSV "
                  f"({size} baris/file){c.END}\n")
            confirm = input(f"{c.BOLD}  Lanjut? (y/n): {c.END}").strip().lower()
            if confirm == 'y':
                return size
        except ValueError:
            print(f"{c.ERR}  ❌ Masukkan angka yang valid.{c.END}")


def ask_output_dir(filepath: str) -> str:
    """Tanya direktori output untuk file CSV."""
    default_dir = os.path.join(os.path.dirname(filepath), "output_csv")
    inp = input(
        f"\n{c.BOLD}📁 Direktori output "
        f"(Enter untuk default: {default_dir}): {c.END}"
    ).strip()
    out_dir = inp if inp else default_dir
    os.makedirs(out_dir, exist_ok=True)
    print(f"{c.OK}  ✓ Output: {out_dir}{c.END}\n")
    return out_dir


def split_and_export(ws, headers, selected_indices, split_size, out_dir, base_name, fmt):
    """Baca data, split, dan export ke CSV."""
    total_written = 0
    file_num = 1
    batch = []
    selected_headers = [headers[i] for i in selected_indices]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    print(f"{c.CYAN}⏳ Memproses data...{c.END}")

    for row in iterate_rows(ws, selected_indices, fmt):
        batch.append(row)

        if len(batch) >= split_size:
            fname = f"{base_name}_part{file_num:03d}_{timestamp}.csv"
            fpath = os.path.join(out_dir, fname)
            _write_csv(fpath, selected_headers, batch)
            total_written += len(batch)
            print(f"  {c.OK}✓ {fname} ({len(batch)} rows){c.END}")
            batch = []
            file_num += 1

    if batch:
        fname = f"{base_name}_part{file_num:03d}_{timestamp}.csv"
        fpath = os.path.join(out_dir, fname)
        _write_csv(fpath, selected_headers, batch)
        total_written += len(batch)
        print(f"  {c.OK}✓ {fname} ({len(batch)} rows){c.END}")

    return total_written, file_num


def _write_csv(filepath, headers, rows):
    """Tulis satu file CSV dengan BOM (utf-8-sig) agar Excel bisa baca."""
    with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)


def main():
    banner()

    # 1. Tanya path file
    filepath = ask_file_path()
    base_name = os.path.splitext(os.path.basename(filepath))[0]

    # 2. Auto-detect format & baca headers
    wb, ws, headers, fmt = read_headers(filepath)

    # 3. Tanya kolom mana yang dibutuhkan
    selected_indices = show_columns(headers)

    # 4. Hitung total baris data
    total_rows = get_row_count(ws, fmt)
    if total_rows <= 0:
        print(f"{c.ERR}  ❌ Tidak ada baris data.{c.END}")
        sys.exit(1)

    # 5. Tanya split size
    split_size = ask_split_size(total_rows)

    # 6. Tanya direktori output
    out_dir = ask_output_dir(filepath)

    # 7. Proses split & export
    print(f"\n{c.BOLD}🚀 Mulai proses split...{c.END}\n")
    total_written, file_count = split_and_export(
        ws, headers, selected_indices, split_size, out_dir, base_name, fmt
    )

    # 8. Summary
    print(f"\n{c.OK}{c.BOLD}══════════════════════════════════════")
    print(f"  ✅ SELESAI!")
    print(f"  📊 Total baris: {total_written}")
    print(f"  📁 Total file:  {file_count}")
    print(f"  📂 Lokasi:      {out_dir}")
    print(f"══════════════════════════════════════{c.END}\n")

    if hasattr(wb, 'close'):
        wb.close()


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n{c.WARN}Dibatalkan oleh user.{c.END}")
        sys.exit(0)
    except Exception as e:
        print(f"\n{c.ERR}Error: {e}{c.END}")
        sys.exit(1)
