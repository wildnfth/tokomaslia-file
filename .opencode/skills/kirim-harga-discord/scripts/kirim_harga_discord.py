# -*- coding: utf-8 -*-
"""
kirim_harga_discord.py — SATU pintu kirim foto harga ke Discord.

Jangan tulis ulang screenshot/webhook. Script ini:
  1. foto BLOK TERAKHIR saja (bukan seluruh riwayat)
  2. tolak PNG kosong
  3. kirim ke channel yang benar
  4. hapus file temp

Pemakaian (dari D:\\TOKO MAS LIA\\DOKUMEN):
  python scripts/kirim_harga_discord.py --channel harga-lm-emas
  python scripts/kirim_harga_discord.py --channel harga-lm-perak
  python scripts/kirim_harga_discord.py --channel harga-perhiasan
  python scripts/kirim_harga_discord.py --channel all
"""
import argparse
import os
import sys
import tempfile
from datetime import datetime

# Path repo dikunci — salinan di folder skill jangan sampai baca xlsx yang salah.
REPO = r"D:\TOKO MAS LIA\DOKUMEN"
HERE = os.path.join(REPO, "scripts")
sys.path.insert(0, HERE)

from screenshot_harga import render  # noqa: E402

CHANNEL_MAP = {
    "harga-lm-emas": {
        "file": "TEMPLATE HARGA MAS2.xlsx",
        "sheet": "ANTAM",
        "prefix": "Harga LM",
    },
    "harga-lm-perak": {
        "file": "TEMPLATE HARGA PERAK.xlsx",
        "sheet": "Sheet1",
        "prefix": "Harga LM Perak",
    },
    "harga-perhiasan": {
        "file": "TEMPLATE HARGA MAS2.xlsx",
        "sheet": "EMAS",
        "prefix": "Harga Perhiasan",
    },
}


def latest_date(xlsx_path, sheet):
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet]
    date = None
    if sheet == "EMAS":
        date = ws.cell(1, 1).value
    else:
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if not isinstance(v, str):
                continue
            u = v.strip()
            up = u.upper()
            if up.startswith("TANGGAL"):
                parts = u.split(None, 1)
                date = parts[1] if len(parts) > 1 else u
            elif up.startswith("HARGA ANTAM"):
                nxt = ws.cell(r + 1, 1).value
                if isinstance(nxt, str) and nxt.strip() and not nxt.upper().startswith("HARGA"):
                    date = nxt.strip()
                elif " - " in u:
                    date = u.split(" - ", 1)[1].strip()
    wb.close()
    if date is None:
        return None
    return str(date).strip()


def send_one(channel, caption_override=None):
    spec = CHANNEL_MAP[channel]
    xlsx = os.path.join(REPO, spec["file"])
    if not os.path.exists(xlsx):
        raise SystemExit("File tidak ditemukan: %s" % xlsx)

    date = latest_date(xlsx, spec["sheet"])
    caption = caption_override or (
        "%s %s" % (spec["prefix"], date) if date else spec["prefix"]
    )
    tmp = os.path.join(
        tempfile.gettempdir(),
        "kirim_%s_%s.png" % (channel, datetime.now().strftime("%H%M%S")),
    )
    try:
        try:
            out = render(xlsx, spec["sheet"], None, tmp)
        except Exception as e:
            print("[kirim] screenshot gagal (%s), retry 1x instance Excel baru" % e)
            out = render(xlsx, spec["sheet"], None, tmp)
        size = os.path.getsize(out)
        print("[kirim] %s | %s | %d bytes | %s" % (channel, spec["sheet"], size, caption))
        send_py = os.path.join(HERE, "send_discord.py")
        import subprocess
        subprocess.run(
            [sys.executable, send_py, out, "--channel", channel, "--caption", caption],
            check=True,
        )
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass


def main():
    ap = argparse.ArgumentParser(description="Kirim blok harga terakhir ke Discord.")
    ap.add_argument(
        "--channel",
        required=True,
        choices=list(CHANNEL_MAP.keys()) + ["all"],
        help="Channel tujuan, atau 'all' untuk ketiga channel.",
    )
    ap.add_argument("--caption", default=None, help="Override teks pesan (opsional).")
    args = ap.parse_args()

    os.chdir(REPO)
    channels = list(CHANNEL_MAP.keys()) if args.channel == "all" else [args.channel]
    for ch in channels:
        send_one(ch, args.caption)
    print("[kirim] selesai: %s" % ", ".join(channels))


if __name__ == "__main__":
    main()
