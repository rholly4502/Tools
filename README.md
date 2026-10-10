# Tools

Kumpulan tool/script utility untuk pekerjaan sehari-hari.

## Python

### Python_SplitXLSXtoCSV
Tool interaktif untuk membaca file Excel (.xlsx), memilih kolom yang dibutuhkan, lalu split menjadi multiple file CSV berdasarkan jumlah baris per file.

**Cara pakai:**
```bash
pip install openpyxl
python Python_SplitXLSXtoCSV/split_xlsx_to_csv.py
```

**Fitur:**
- Pilih kolom tertentu atau semua kolom
- Split berdasarkan jumlah baris per file
- Output CSV dengan UTF-8 BOM (Excel-friendly)
- Auto-naming dengan timestamp dan nomor part
