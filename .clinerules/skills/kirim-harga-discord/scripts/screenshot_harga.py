# -*- coding: utf-8 -*-
"""
screenshot_harga.py - Render range tabel Excel (.xlsx) menjadi PNG via Excel COM.

Membuka file di instance Excel baru (background), CopyPicture sebagai bitmap,
lalu simpan PNG dari clipboard. Akurat sesuai tampilan asli
(font/border/warna/merged cells).

Pemakaian:
  screenshot_harga.py <file.xlsx> --sheet ANTAM --range "A1:M40" --out out.png
  screenshot_harga.py <file.xlsx> --sheet ANTAM --out out.png
      # range kosong: blok harga terakhir (bukan seluruh riwayat)

Catatan: butuh Microsoft Excel + pywin32. Pillow (PIL) disarankan.
"""
import argparse
import os
import sys
import time

try:
    import win32com.client
except Exception as e:  # pragma: no cover
    sys.exit("Butuh pywin32: pip install pywin32 (%s)" % e)

XL_SCREEN = 1
XL_BITMAP = 2
# Sheet panjang (riwayat harian) jangan di-capture utuh — Discord jadi
# gambar raksasa / terlihat kosong. Ambang: pakai blok terakhir.
USED_RANGE_MAX_ROWS = 80
BLOCK_MAX_COL = 16


def _cell_text(value):
    return value.strip().upper() if isinstance(value, str) else ""


def _last_nonempty_row(ws, start_row, end_row, last_col):
    r = end_row
    while r > start_row:
        for c in range(1, last_col + 1):
            if ws.Cells(r, c).Value is not None:
                return r
        r -= 1
    return start_row


def _last_used_col(ws, start_row, end_row, last_col):
    used = 1
    for r in range(start_row, end_row + 1):
        c = last_col
        while c > used:
            if ws.Cells(r, c).Value is not None:
                used = c
                break
            c -= 1
    # Merged header/harga (mis. UBS K:M) tidak punya value di kolom kanan.
    for r in range(start_row, min(end_row, start_row + 4) + 1):
        for c in range(1, used + 1):
            cell = ws.Cells(r, c)
            try:
                if cell.MergeCells:
                    area = cell.MergeArea
                    used = max(used, int(area.Column + area.Columns.Count - 1))
            except Exception:
                pass
    return min(used, last_col)


def _last_block_range(ws):
    """Blok harga terakhir, atau UsedRange jika sheet-nya pendek."""
    used = ws.UsedRange
    first_row = int(used.Row)
    n_rows = int(used.Rows.Count)
    n_cols = int(used.Columns.Count)
    last_row = first_row + n_rows - 1
    last_col = min(int(used.Column) + n_cols - 1, BLOCK_MAX_COL)

    # UsedRange Excel sering kepanjangan (baris style kosong).
    end = _last_nonempty_row(ws, first_row, last_row, last_col)
    content_rows = end - first_row + 1
    if content_rows <= USED_RANGE_MAX_ROWS:
        col = _last_used_col(ws, first_row, end, last_col)
        rng = ws.Range(ws.Cells(first_row, 1), ws.Cells(end, col))
        return rng, rng.Address

    header_row = None
    for r in range(last_row, first_row - 1, -1):
        u = _cell_text(ws.Cells(r, 1).Value)
        if u.startswith("TANGGAL"):
            header_row = r
            break
        if u.startswith("HARGA"):
            header_row = r
            prev = _cell_text(ws.Cells(r - 1, 1).Value) if r > 1 else ""
            if prev.startswith("TANGGAL"):
                header_row = r - 1
            break
    if header_row is None:
        return used, used.Address

    end = _last_nonempty_row(ws, header_row, last_row, last_col)
    col = _last_used_col(ws, header_row, end, last_col)
    rng = ws.Range(ws.Cells(header_row, 1), ws.Cells(end, col))
    return rng, rng.Address


def _grab_clipboard_image(retries=3, delay=0.35):
    try:
        from PIL import ImageGrab
    except Exception:
        return None
    img = None
    for i in range(retries):
        time.sleep(delay if i == 0 else delay * 2)
        img = ImageGrab.grabclipboard()
        if img is not None and hasattr(img, "save"):
            return img
    return None


def _export_via_clipboard(rng, out_path):
    rng.CopyPicture(XL_SCREEN, XL_BITMAP)
    img = _grab_clipboard_image()
    if img is None:
        return False
    img.save(os.path.abspath(out_path), "PNG")
    return True


def _export_via_chart(ws, rng, out_path):
    """Cadangan jika Pillow tidak ada / clipboard gagal."""
    rng.CopyPicture(XL_SCREEN, XL_BITMAP)
    time.sleep(0.4)
    chart = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
    try:
        chart.Activate()
        chart.Chart.Paste()
        chart.Chart.Export(os.path.abspath(out_path), "PNG")
    finally:
        try:
            chart.Delete()
        except Exception:
            pass


def _png_is_blank(path):
    try:
        from PIL import Image
    except Exception:
        return os.path.getsize(path) < 5000
    im = Image.open(path).convert("RGB")
    # sampel kecil: gambar tabel punya banyak warna (border/isi)
    sample = im.resize((48, 24))
    colors = sample.getcolors(48 * 24) or []
    if len(colors) <= 2:
        return True
    # hampir semua piksel putih/abu sangat muda
    near_white = 0
    total = 0
    for n, (r, g, b) in colors:
        total += n
        if r > 245 and g > 245 and b > 245:
            near_white += n
    return total > 0 and (near_white / float(total)) > 0.97


def render(xlsx_path, sheet, range_str, out_path):
    xl = None
    wb = None
    try:
        # Instance baru — jangan nempel ke Excel user, jangan Quit punya orang.
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AskToUpdateLinks = False
        try:
            xl.EnableEvents = False
        except Exception:
            pass
        wb = xl.Workbooks.Open(
            os.path.abspath(xlsx_path),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
        )
        ws = wb.Worksheets(sheet)
        ws.Activate()
        try:
            xl.ActiveWindow.Zoom = 100
        except Exception:
            pass

        if range_str:
            rng = ws.Range(range_str)
            addr = rng.Address
        else:
            rng, addr = _last_block_range(ws)
        print("[screenshot] range %s!%s" % (sheet, addr))

        ok = _export_via_clipboard(rng, out_path)
        if not ok:
            print("[screenshot] clipboard gagal, fallback Chart.Export")
            _export_via_chart(ws, rng, out_path)

        out_abs = os.path.abspath(out_path)
        if not os.path.exists(out_abs) or os.path.getsize(out_abs) < 1000:
            raise RuntimeError("PNG tidak tertulis atau terlalu kecil: %s" % out_abs)
        if _png_is_blank(out_abs):
            raise RuntimeError("PNG terlihat kosong (hampir polos). Range: %s" % addr)
        return out_abs
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        # lepas referensi COM biar proses Excel benar-benar mati
        ws = None
        rng = None
        wb = None
        xl = None


def main():
    ap = argparse.ArgumentParser(description="Screenshot tabel Excel ke PNG (via Excel COM).")
    ap.add_argument("file", help="Path file .xlsx.")
    ap.add_argument("--sheet", required=True, help="Nama sheet, mis. ANTAM / EMAS / Sheet1.")
    ap.add_argument("--out", required=True, help="Path output PNG.")
    ap.add_argument("--range", default=None, help="Range A1:M40. Jika kosong: blok terakhir.")
    args = ap.parse_args()

    if not os.path.exists(args.file):
        raise SystemExit("File tidak ditemukan: %s" % args.file)

    out = render(args.file, args.sheet, args.range or None, args.out)
    print("[screenshot] %s -> %s (%d bytes)" % (args.sheet, out, os.path.getsize(out)))


if __name__ == "__main__":
    main()
