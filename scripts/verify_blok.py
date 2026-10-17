#!/usr/bin/env python3
"""verify_blok.py - verifikasi blok ANTAM terakhir vs blok sumber (atau vs HEAD).

Dipakai SETELAH update_harga.py (mode add / update) untuk membuktikan angka,
tanggal dan format benar - bukan sekadar membaca ringkasan "[ok]" dari script.

Yang dicek:
  1. tanggal header blok terakhir (A/G/K) == --date (kalau diberikan)
  2. nilai: blok terakhir == blok sumber + step x gramasi, per merk
     (SOLD / '-' dianggap sama; pembulatan 0.5 gr searah step)
  3. merged cells blok terakhir == blok sumber (offset baris)
  4. number_format blok terakhir == blok sumber
  5. urutan ANTAM wajib 2026 > 2025 > RANDOM

Contoh:
  # blok baru: ANTAM +10/g, Galeri24 +10/g, UBS disalin tetap
  python scripts/verify_blok.py "TEMPLATE HARGA MAS2.xlsx" --date "8 OKTOBER 2026" \
      --antam 10 --galeri24 10 --ubs 0

  # revisi di tempat: hanya Galeri24 yang berubah
  python scripts/verify_blok.py "TEMPLATE HARGA MAS2.xlsx" --galeri24 -15

  # bandingkan blok terakhir file kerja vs versi HEAD (butuh git LFS)
  python scripts/verify_blok.py "TEMPLATE HARGA MAS2.xlsx" --vs-head

Exit 0 = semua cocok; exit 1 = ada selisih (detail dicetak).
"""
import argparse
import math
import os
import re
import subprocess
import sys

from openpyxl import load_workbook

MARK = "HARGA ANTAM"
SOLD = {"SOLD", "-", "KOSONG"}
BRANDS = [
    ("RETRO", 2, 1),
    ("RANDOM", 3, 1),
    ("2025", 4, 1),
    ("2026", 5, 1),
    ("GALERI24", 8, 7),
    ("UBS", 12, 11),
]


def header_rows(ws):
    out = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip().upper().startswith(MARK):
            out.append(r)
    return out


def cell(ws, r, c):
    v = ws.cell(r, c).value
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip().upper()
        return "SOLD" if s in SOLD else v.strip()
    return round(float(v), 1)


def snap5(x, up):
    """Bulatkan ke kelipatan 5 searah step (naik -> ceil, turun -> floor)."""
    return int(math.ceil(x / 5.0) * 5) if up else int(math.floor(x / 5.0) * 5)


def shift_merge(m, off):
    return re.sub(r"\d+", lambda x: str(int(x.group(0)) + off), m)


def smudge_head(xlsx, out):
    """Tulis isi file di HEAD (blob LFS di-smudge) ke path `out`."""
    repo = os.path.dirname(os.path.abspath(xlsx))
    name = os.path.basename(xlsx)
    blob = subprocess.run(["git", "-C", repo, "show", "HEAD:%s" % name],
                          capture_output=True, check=True).stdout
    with open(out, "wb") as f:
        subprocess.run(["git", "-C", repo, "lfs", "smudge"], input=blob,
                       stdout=f, check=True)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--date", help="tanggal yang diharapkan di header blok terakhir")
    ap.add_argument("--antam", type=float, default=0.0,
                    help="step per gram utk RETRO/RANDOM/2025/2026")
    ap.add_argument("--galeri24", type=float, default=0.0)
    ap.add_argument("--ubs", type=float, default=0.0)
    ap.add_argument("--vs-head", action="store_true",
                    help="bandingkan blok terakhir vs versi HEAD (butuh git LFS)")
    a = ap.parse_args()

    tmp = None
    src_path = a.xlsx
    if a.vs_head:
        tmp = os.path.join(os.path.dirname(os.path.abspath(a.xlsx)), ".verify_head.xlsx")
        smudge_head(a.xlsx, tmp)
        src_path = tmp

    ws_src = load_workbook(src_path, data_only=True)["ANTAM"]
    wf_src = load_workbook(src_path)["ANTAM"]
    ws = load_workbook(a.xlsx, data_only=True)["ANTAM"]
    wf = load_workbook(a.xlsx)["ANTAM"]

    hs = header_rows(ws)
    if not hs:
        sys.exit("tidak ada blok %s di %s" % (MARK, a.xlsx))
    dst = hs[-1]

    if a.vs_head:
        src, off = dst, 0
    else:
        if len(hs) < 2:
            sys.exit("butuh minimal 2 blok untuk membandingkan")
        src, off = hs[-2], dst - hs[-2]

    bad = 0
    steps = {"RETRO": a.antam, "RANDOM": a.antam, "2025": a.antam, "2026": a.antam,
             "GALERI24": a.galeri24, "UBS": a.ubs}

    if a.date:
        for c in (1, 7, 11):
            got = str(ws.cell(dst + 1, c).value or "").strip()
            if got.upper() != a.date.strip().upper():
                bad += 1
                print("TANGGAL beda kolom %d: %r (harus %r)" % (c, got, a.date))
    print("tanggal blok terakhir:", ws.cell(dst + 1, 1).value)

    checked = 0
    for name, col, gcol in BRANDS:
        step = steps[name]
        for r_src in range(src + 2, src + 16):
            r_dst = r_src + off
            g = ws.cell(r_dst, gcol).value
            g = round(float(g), 1) if isinstance(g, (int, float)) else None
            if g is None:
                continue
            old, new = cell(ws_src, r_src, col), cell(ws, r_dst, col)
            if not isinstance(old, (int, float)) or step == 0:
                want = old
            else:
                want = int(old + snap5(step * g, up=(step > 0)))
            checked += 1
            if new != want:
                bad += 1
                print("NILAI beda %-9s %-5sg: %r (harus %r; sumber %r)"
                      % (name, g, new, want, old))

    # Blok sumber berakhir tepat sebelum baris pemisah + header blok berikutnya.
    # Jangan pakai "src + 13" statis: blok sekarang 12 baris, sehingga jendela
    # statis ikut menyeret merge header blok TARGET sendiri (false positive).
    next_after_src = [x - 2 for x in header_rows(wf_src) if x > src]
    src_end = next_after_src[0] if next_after_src else src + 13
    src_merges = [str(m) for m in wf_src.merged_cells.ranges
                  if m.min_row >= src and m.max_row <= src_end]
    dst_merges = [str(m) for m in wf.merged_cells.ranges]
    for m in src_merges:
        want = shift_merge(m, off)
        if want not in dst_merges:
            bad += 1
            print("MERGE hilang:", m, "->", want)

    nf = 0
    for r in range(src, src + 14):
        for c in range(1, 14):
            if wf_src.cell(r, c).number_format != wf.cell(r + off, c).number_format:
                nf += 1
    if nf:
        bad += nf
        print("NUMBER_FORMAT beda: %d sel" % nf)

    for r in range(dst + 2, dst + 16):
        v25, v26, vr = cell(ws, r, 4), cell(ws, r, 5), cell(ws, r, 3)
        if all(isinstance(x, (int, float)) for x in (v25, v26, vr)):
            if not (v26 > v25 > vr):
                bad += 1
                print("ANOMALI urutan row %d: 2026=%s 2025=%s RANDOM=%s" % (r, v26, v25, vr))

    if tmp and os.path.exists(tmp):
        os.remove(tmp)

    print("dicek %d sel | merged sumber %d | masalah %d" % (checked, len(src_merges), bad))
    print("HASIL:", "OK - blok terakhir sesuai harapan" if bad == 0 else "ADA SELISIH")
    sys.exit(0 if bad == 0 else 1)


if __name__ == "__main__":
    main()
