# -*- coding: utf-8 -*-
"""
SKILL: update-harga-perak
=========================
Mengelola TEMPLATE HARGA PERAK.xlsx (sheet Sheet1).

MODUS 'add' (default): TAMBAH blok harian baru dengan MENYALIN blok terakhir
secara PLEK KETIPLEK (style sel, merged cells, tinggi baris), mengganti label
TANGGAL, lalu MENURUNKAN semua harga sesuai --step (rb/gram).

Struktur blok PERAK:
  - Baris = "TANGGAL <tanggal>" (kolom A), lalu "HARGA LOGAM MULIA PERAK".
  - Merek kiri  (STAR SILVER / ANTAM / LOTUS / LM PERAK) uji harga di kolom B(2),
    pasangan berat di kolom A(1).
  - Merek kanan (MT / SIMBA / EURO) harga di kolom E(5), pasangan berat di kolom D(4).
  - Format harga: "71rb/gr", "54,5rb/gr", "25jt" (harga total), "-" / kosong.
  - Pemisah antar blok: 4 baris kosong setelah baris "LM PERAK".

CONTOH:
  python update_harga_perak.py "<file>" --date "13 AGUSTUS 2026" --step 1.5 --dry-run
  python update_harga_perak.py "<file>" --date "13 AGUSTUS 2026" --step 1.5
"""
import argparse
import copy
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime

from openpyxl import load_workbook


def copy_style(src, dst):
    """Salin style sel persis (font/fill/border/alignment/number_format)."""
    if src.has_style:
        dst._style = copy.copy(src._style)


def fmt_num(v):
    """Float -> string, koma desimal, tanpa trailing .0."""
    if abs(v - round(v)) < 1e-9:
        return str(int(round(v)))
    s = ('%.3f' % v).rstrip('0').rstrip('.')
    return s.replace('.', ',')


def parse_weight(text):
    if text is None:
        return None
    t = str(text).strip().lower()
    m = re.match(r'^(\d+(?:[.,]\d+)?)\s*gr$', t)
    if m:
        return float(m.group(1).replace(',', '.'))
    m = re.match(r'^(\d+(?:[.,]\d+)?)\s*kg$', t)
    if m:
        return float(m.group(1).replace(',', '.')) * 1000.0
    return None


def parse_price(text):
    """None bila bukan sel harga. Lainnya -> (angka, unit)."""
    if text is None:
        return None
    t = str(text).strip()
    if t in ('-', ''):
        return None
    m = re.match(r'^([\d.,]+)\s*(rb/gr|jt)\s*$', t, re.I)
    if not m:
        return None
    num = float(m.group(1).replace('.', '').replace(',', '.'))
    return num, m.group(2).lower()


def weight_col_of(c):
    # kolom berat pasangan utk kolom harga B(2) -> A(1), E(5) -> D(4)
    return 1 if c == 2 else (4 if c == 5 else None)


def apply_thick_borders(ws, start):
    """Terapkan border THICK pada blok yang diawali baris `start` (baris TANGGAL).

    - FRAME luar  : baris start..start+13, kolom 1-5.
    - GRUP MEREK  : kiri  (A-B): STAR SILVER 103-107, ANTAM 108-109, LOTUS 110-114
                    kanan (D-E): MT 103-108, SIMBA 110-111, EURO 113-114
      (relatif thd start).
    Catatan penting: writer openpyxl men-collapse border sel merged (non-anchor)
    ke style ANCHOR saat save. Karena itu border tebal untuk sel yang ada di dalam
    merged range WAJIB ditebalkan lewat sel ANCHOR-nya agar render tebal di Excel.
    """
    from openpyxl.styles import Border as _Border, Side as _Side
    THICK = _Side(style='thick')

    # peta (baris,kolom) -> anchor utk semua merged range
    anchor_map = {}
    for m in ws.merged_cells.ranges:
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                anchor_map[(r, c)] = (m.min_row, m.min_col)

    def set_thick(r, c, attr):
        if (r, c) in anchor_map and (r, c) != anchor_map[(r, c)]:
            r, c = anchor_map[(r, c)]
        cell = ws.cell(r, c)
        b = cell.border
        kw = {'left': b.left, 'right': b.right, 'top': b.top, 'bottom': b.bottom}
        kw[attr] = THICK
        cell.border = _Border(**kw)

    def box(r1r, r2r, lidx, ridx):
        for off in range(r1r, r2r + 1):
            r = start + off
            if off == r1r:
                set_thick(r, lidx, 'top')
                set_thick(r, ridx, 'top')
            if off == r2r:
                set_thick(r, lidx, 'bottom')
                set_thick(r, ridx, 'bottom')
            set_thick(r, lidx, 'left')
            set_thick(r, ridx, 'right')

    # frame luar
    for c in range(1, 6):
        set_thick(start, c, 'top')
        set_thick(start + 13, c, 'bottom')
    for r in range(start, start + 14):
        set_thick(r, 1, 'left')
        set_thick(r, 5, 'right')

    # grup merek
    box(2, 6, 1, 2)    # STAR SILVER
    box(7, 8, 1, 2)    # ANTAM
    box(9, 13, 1, 2)   # LOTUS
    box(2, 7, 4, 5)    # MT
    box(9, 10, 4, 5)   # SIMBA
    box(12, 13, 4, 5)  # EURO



def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('file')
    ap.add_argument('--date', required=True,
                    help='label tanggal baru, mis. "13 AGUSTUS 2026"')
    ap.add_argument('--step', type=float, required=True,
                    help='penurunan harga per gram (rb/gr, positif = turun)')
    ap.add_argument('--dry-run', action='store_true')
    ap.add_argument('--backup', action='store_true',
                    help='Buat file *_BACKUP_*.xlsx. Default off; revert lewat git.')
    ap.add_argument('--no-backup', action='store_true', help=argparse.SUPPRESS)
    ap.add_argument('--send-discord', action='store_true',
                    help='Kirim foto tabel ke channel Discord setelah update.')
    ap.add_argument('--discord-channel', default='harga-lm-perak',
                    help='Nama channel di discord_config.json.')
    args = ap.parse_args()

    wb = load_workbook(args.file)
    ws = wb.active

    # ---- cari blok terakhir ----
    tang = [r for r in range(1, ws.max_row + 1)
            if isinstance(ws.cell(r, 1).value, str)
            and ws.cell(r, 1).value.strip().upper().startswith('TANGGAL ')]
    if not tang:
        raise SystemExit('[err] tidak ada baris TANGGAL di kolom A.')
    src_start = tang[-1]
    old_header = ws.cell(src_start, 1).value

    lm = None
    for r in range(src_start, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip().upper() == 'LM PERAK':
            lm = r
    if lm is None:  # fallback: blok berakhir di baris non-kosong terakhir
        end = ws.max_row
        while end > src_start and all(ws.cell(end, c).value is None
                                      for c in range(1, ws.max_column + 1)):
            end -= 1
        lm = end

    block = list(range(src_start, lm + 1))
    n_row = len(block)
    new_start = lm + 5          # 4 baris kosong pemisah + 1
    new_block = [new_start + i for i in range(n_row)]
    spacing = new_start - src_start
    new_header = 'TANGGAL ' + args.date
    max_col = ws.max_column

    def transform(row, col, val):
        pr = parse_price(val)
        if pr is None:
            return None
        num, unit = pr
        if unit == 'rb/gr':
            return fmt_num(num - args.step) + 'rb/gr'
        # unit 'jt' = harga total; kurangi step*berat/1000 jt
        wcol = weight_col_of(col)
        wt = parse_weight(ws.cell(row, wcol).value) if wcol else None
        if not wt:
            print('[warn] harga total tanpa berat terurai r%d c%d, dilewati'
                  % (row, col))
            return None
        return fmt_num(num - args.step * wt / 1000.0) + 'jt'

    # ---- laporan perubahan ----
    changes = []
    for r in block:
        for c in range(1, max_col + 1):
            sv = ws.cell(r, c).value
            nv = transform(r, c, sv)
            if nv is not None and nv != sv:
                changes.append((r, c, sv, nv))

    print('[info] blok sumber %d-%d : %s' % (src_start, lm, old_header))
    print('[info] blok baru   %d-%d : %s' % (new_block[0], new_block[-1], new_header))
    print('[info] step %g rb/gr | %d sel harga akan diubah' % (args.step, len(changes)))

    if args.dry_run:
        print('[dry-run] perubahan:')
        for r, c, sv, nv in changes:
            print('  r%d c%d : %r -> %r' % (r, c, sv, nv))
        print('[dry-run] selesai (belum disimpan).')
        return

    # ---- backup (opt-in; default off — revert lewat git) ----
    if args.backup:
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup = args.file.replace('.xlsx', '_BACKUP_%s.xlsx' % ts)
        shutil.copy2(args.file, backup)
        print('[ok] backup -> %s' % backup)

    # ---- salin nilai + style, ubah harga ----
    for di, sr in enumerate(block):
        dr = new_block[di]
        for c in range(1, max_col + 1):
            s = ws.cell(sr, c)
            d = ws.cell(dr, c)
            nv = transform(sr, c, s.value)
            d.value = nv if nv is not None else s.value
            copy_style(s, d)

    # tinggi baris
    for di, sr in enumerate(block):
        dr = new_block[di]
        if sr in ws.row_dimensions and ws.row_dimensions[sr].height is not None:
            ws.row_dimensions[dr].height = ws.row_dimensions[sr].height

    # duplikasi merged cells blok sumber (offset baris)
    for m in list(ws.merged_cells.ranges):
        if src_start <= m.min_row <= lm:
            ws.merge_cells(start_row=m.min_row + spacing, start_column=m.min_col,
                           end_row=m.max_row + spacing, end_column=m.max_col)

    # tanggal baru
    ws.cell(new_start, 1).value = new_header

    # border THICK: frame luar + tiap grup merek (judul+gramasi)
    apply_thick_borders(ws, new_start)

    wb.save(args.file)
    print('[ok] %s -> blok baru %d-%d (%s) | %d harga diubah'
          % (args.file, new_block[0], new_block[-1], new_header, len(changes)))

    if args.send_discord:
        _send_discord(args.file, new_header, args.discord_channel)


def _send_discord(xlsx_path, caption, channel):
    """Render sheet aktif jadi PNG lalu kirim ke channel Discord (webhook)."""
    here = os.path.dirname(os.path.abspath(__file__))
    tmp = os.path.join(tempfile.gettempdir(), "perak_snap_%s.png" % datetime.now().strftime("%H%M%S"))
    try:
        sys.path.insert(0, here)
        from screenshot_harga import render
        render(xlsx_path, "Sheet1", None, tmp)
    except Exception as e:
        print("[warn] gagal screenshot:", e)
        return
    try:
        subprocess.run(
            [sys.executable, os.path.join(here, "send_discord.py"),
             tmp, "--channel", channel, "--caption", caption],
            check=True)
    except subprocess.CalledProcessError as e:
        print("[warn] gagal kirim Discord:", e)
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass


if __name__ == '__main__':
    main()
