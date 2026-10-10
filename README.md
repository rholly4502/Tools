# Tools

Kumpulan tool/script utility untuk pekerjaan sehari-hari.

## Scripts

### split_xlsx_to_csv.py
Tool interaktif untuk membaca file Excel (.xlsx), memilih kolom yang dibutuhkan, lalu split menjadi multiple file CSV berdasarkan jumlah baris per file.

**Cara pakai:**
```bash
pip install openpyxl
python split_xlsx_to_csv.py
```

**Fitur:**
- Pilih kolom tertentu atau semua kolom
- Split berdasarkan jumlah baris per file
- Output CSV dengan UTF-8 BOM (Excel-friendly)
- Auto-naming dengan timestamp dan nomor part
