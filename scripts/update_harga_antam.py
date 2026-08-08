#!/usr/bin/env python3
"""
Update HARGA ANTAM (sheet ANTAM) — TAMBAH blok harian baru.
Pola: salin blok harga terakhir, tambah 1 baris kosong, update harga naik/turun
sesuai --step, ganti tanggal, pertahankan format sel MERAH = KOSONG.

Jalankan:
    python update_harga_antam.py "TEMPLATE HARGA MAS2.xlsx" --date "8 AGUSTUS 2026" --step -40
    python update_harga_antam.py "TEMPLATE HARGA MAS2.xlsx" --date "8 AGUSTUS 2026" --step -40 --dry-run
"""

import argparse
import os
import shutil
import sys
from datetime import datetime

try:
    from openpyxl import load_workbook
except ImportError:
    print("ERROR: openpyxl tidak terinstal. Jalankan: pip install openpyxl")
    sys.exit(1)

def parse_args():
    parser = argparse.ArgumentParser(description="Update harga ANTAM (tambah blok baru)")
    parser.add_argument("file", help="Path ke file Excel TEMPLATE HARGA MAS2.xlsx")
    parser.add_argument("--date", required=True, help="Tanggal baru, contoh: '8 AGUSTUS 2026'")
    parser.add_argument("--step", type=int, default=50, help="Kenaikan/turun per gram (default 50)")
    parser.add_argument("--dry-run", action="store_true", help="Tampilkan preview tanpa menulis")
    parser.add_argument("--no-backup", action="store_true", help="Nonaktifkan backup otomatis")
    return parser.parse_args()


def backup_file(filepath):
    """Buat backup file dengan timestamp."""
    if not os.path.exists(filepath):
        return None
    base, ext = os.path.splitext(filepath)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{base}_BACKUP_{timestamp}{ext}"
    shutil.copy2(filepath, backup_path)
    print(f"[BACKUP] {filepath} -> {backup_path}")
    return backup_path

def find_last_block(ws):
    """Cari blok terakhir di sheet ANTAM (era baru)."""
    max_row = ws.max_row
    for r in range(max_row, 1, -1):
        cell_a = ws.cell(row=r, column=1).value
        if cell_a and isinstance(cell_a, str) and "HARGA ANTAM" in cell_a and "-" in cell_a:
            metadata_row = r + 1
            data_start = r + 2
            data_end = data_start
            for dr in range(data_start, max_row + 1):
                val_a = ws.cell(row=dr, column=1).value
                val_g = ws.cell(row=dr, column=7).value
                if val_a is None and val_g is None:
                    break
                if val_a == "MERAH = KOSONG" or val_g == "MERAH = KOSONG":
                    next_a = ws.cell(row=dr + 1, column=1).value if dr + 1 <= max_row else None
                    next_g = ws.cell(row=dr + 1, column=7).value if dr + 1 <= max_row else None
                    if next_a is None and next_g is None:
                        data_end = dr
                        break
                data_end = dr
            return r, data_start, data_end, metadata_row
    raise ValueError("Tidak dapat menemukan blok HARGA ANTAM terakhir di sheet.")

def copy_block(ws, header_row, data_start, data_end, metadata_row, new_date, step, dry_run=False):
    """Salin blok terakhir ke bawah, update harga dan tanggal."""
    new_header_row = data_end + 2
    new_metadata_row = new_header_row + 1
    new_data_start = new_metadata_row + 1

    new_header_a = f"HARGA ANTAM - {new_date}"
    new_header_g = "HARGA Galeri24"
    new_header_k = "HARGA UBS BATIK"

    if dry_run:
        print(f"[DRY-RUN] Akan menambahkan blok baru di baris {new_header_row}")
    else:
        ws.cell(row=new_header_row, column=1, value=new_header_a)
        ws.cell(row=new_header_row, column=7, value=new_header_g)
        ws.cell(row=new_header_row, column=11, value=new_header_k)

    for col in range(1, 17):
        old_val = ws.cell(row=metadata_row, column=col).value
        if old_val is not None:
            if dry_run:
                print(f"[DRY-RUN] Metadata {new_metadata_row} kol {col}: salin '{old_val}'")
            else:
                ws.cell(row=new_metadata_row, column=col, value=old_val)

    if dry_run:
        print(f"[DRY-RUN] Metadata G{new_metadata_row}: update tanggal -> '{new_date}'")
        print(f"[DRY-RUN] Metadata K{new_metadata_row}: update tanggal -> '{new_date}'")
    else:
        ws.cell(row=new_metadata_row, column=7, value=new_date)
        ws.cell(row=new_metadata_row, column=11, value=new_date)

    print(f"[INFO] Menyalin data baris {data_start}-{data_end}, step {step:+d}")
    rows_copied = 0
    for i, old_row in enumerate(range(data_start, data_end + 1)):
        new_row = new_data_start + i
        row_data = {col: ws.cell(row=old_row, column=col).value for col in range(1, 17)}
        is_merah = row_data.get(1) == "MERAH = KOSONG" or row_data.get(7) == "MERAH = KOSONG"

        if is_merah:
            merah_col = 7 if row_data.get(7) == "MERAH = KOSONG" else 1
            if dry_run:
                print(f"[DRY-RUN] Baris {new_row}: pertahankan MERAH = KOSONG")
            else:
                ws.cell(row=new_row, column=merah_col, value="MERAH = KOSONG")
            continue

        price_cols = {2: "B", 3: "C", 4: "D", 5: "E", 8: "H", 9: "I", 12: "L"}
        for col, desc in price_cols.items():
            val = row_data.get(col)
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                new_val = val + step
                if dry_run:
                    print(f"[DRY-RUN] {new_row} {desc}: {val:,.0f} -> {new_val:,.0f}")
                else:
                    ws.cell(row=new_row, column=col, value=new_val)

        for col in [1, 6, 7, 10, 11, 13, 14, 15, 16]:
            val = row_data.get(col)
            if val is not None:
                if dry_run:
                    print(f"[DRY-RUN] {new_row} kol {col}: salin '{val}'")
                else:
                    ws.cell(row=new_row, column=col, value=val)
        rows_copied += 1

    print(f"[INFO] Total baris disalin: {rows_copied}")
    return new_header_row


def main():
    args = parse_args()
    filepath = os.path.abspath(args.file)

    if not os.path.exists(filepath):
        print(f"ERROR: File tidak ditemukan: {filepath}")
        sys.exit(1)

    print(f"[INFO] File: {filepath}")
    print(f"[INFO] Tanggal: {args.date}, Step: {args.step:+d}, Dry-run: {args.dry_run}")

    wb = load_workbook(filepath)
    if "ANTAM" not in wb.sheetnames:
        print("ERROR: Sheet 'ANTAM' tidak ditemukan.")
        sys.exit(1)
    ws = wb["ANTAM"]

    header_row, data_start, data_end, metadata_row = find_last_block(ws)
    print(f"[INFO] Blok terakhir: header={header_row}, metadata={metadata_row}, data={data_start}-{data_end}")

    if args.dry_run:
        print("\n" + "=" * 60)
        print("DRY-RUN MODE — Tidak ada yang ditulis ke file")
        print("=" * 60)
        copy_block(ws, header_row, data_start, data_end, metadata_row, args.date, args.step, dry_run=True)
        print("\n[INFO] Jika sesuai, jalankan lagi tanpa --dry-run")
        return

    if not args.no_backup:
        backup_file(filepath)

    new_header_row = copy_block(ws, header_row, data_start, data_end, metadata_row, args.date, args.step, dry_run=False)
    wb.save(filepath)
    print(f"\n[SUCCESS] {filepath} diupdate, blok baru di baris {new_header_row}")


if __name__ == "__main__":
    main()

