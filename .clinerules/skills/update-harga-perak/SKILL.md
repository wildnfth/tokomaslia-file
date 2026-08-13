---
name: update-harga-perak
description: "Update harga logam mulia perak di file TEMPLATE HARGA PERAK.xlsx (sheet Sheet1) dengan menambah blok harian baru yang menyalin blok terakhir secara plek ketiplek (style, merged cells, tinggi baris), mengganti tanggal, dan menyesuaikan harga (umumnya turun/naik sekian rb/gr). Gunakan saat user minta update harga perak, harga logam mulia perak, LM perak, atau menambah blok harga perak."
---

# Update Harga Perak (LM Perak)

Skill ini mengelola `TEMPLATE HARGA PERAK.xlsx` (sheet **Sheet1**) — harga logam
mulia perak berbagai merek. Berbeda dengan harga emas, file perak **tidak** punya
sheet ANTAM/EMAS; strukturnya satu kolom blok per tanggal dengan merek kiri & kanan.

Script: `scripts/update_harga_perak.py`

## Struktur file `TEMPLATE HARGA PERAK.xlsx`

Sheet `Sheet1` berisi beberapa **blok tanggal**, masing-masing diawali baris
`TANGGAL <tanggal>` (kolom A), lalu judul "HARGA LOGAM MULIA PERAK", lalu daftar
merek dengan harga:

| Posisi | Kolom | Isi |
|--------|-------|-----|
| Merek kiri | A | STAR SILVER / ANTAM / LOTUS / **LM PERAK** |
| Harga kiri | **B** (2) | `"71rb/gr"`, `"54,5rb/gr"`, `"25jt"` (total), atau `-` |
| Berat kiri | A (baris harga kiri) | `5 gr`, `500 gr`, `1 kg` |
| Merek kanan | D | MT / SIMBA / EURO |
| Harga kanan | **E** (5) | `"66rb/gr"`, dst. |
| Berat kanan | D (baris harga kanan) | `5 gr`, `1 kg` |

- Format harga bisa: **`rb/gr`** (per gram), **`jt`** (harga total, mis. ANTAM
  500 gr = `25jt`), atau **`-`** (tidak ada harga).
- Desimal memakai koma: `54,5rb/gr`.
- Setiap blok diakhiri baris label **`LM PERAK`** (bisa kosong isinya).
- **Pemisah antar blok = 4 baris kosong** setelah baris `LM PERAK`.

## Cara pakai script

Script bekerja **mode `add`** (default): menyalin **blok terakhir** secara
**plek ketiplek** (style font/border/fill/number_format, merged cells, tinggi
baris), mengganti label `TANGGAL`, lalu menyesuaikan harga sesuai `--step`.

```powershell
# selalu mulai dengan --dry-run dulu
python scripts/update_harga_perak.py "TEMPLATE HARGA PERAK.xlsx" --date "13 AGUSTUS 2026" --step 1.5 --dry-run

# jalankan sungguhan (backup otomatis AKTIF)
python scripts/update_harga_perak.py "TEMPLATE HARGA PERAK.xlsx" --date "13 AGUSTUS 2026" --step 1.5
```

| Argumen | Wajib | Fungsi |
|---------|-------|--------|
| `<file>` | ya | Path ke `TEMPLATE HARGA PERAK.xlsx` |
| `--date` | ya | Label tanggal baru, mis. `"13 AGUSTUS 2026"` |
| `--step` | ya | Penyesuaian harga per gram dalam **rb/gr**. Positif = **turun** (mis. `1.5`), negatif = naik (mis. `-2`) |
| `--dry-run` | opsional | Tampilkan perubahan tanpa menyimpan |
| `--no-backup` | opsional | Matikan backup otomatis (**hindari** tanpa alasan) |

> **Catatan `--step`:** di file harga emas "naik = +step"; di skill perak untuk
> "turunkan 1,5 rb/gr" gunakan `--step 1.5` (positif = turun). Sesuaikan tanda
> dengan maksud pengguna dan konfirmasikan bila ambigu.

Script menyimpannya sebagai string persis (mis. `71rb/gr` → `69,5rb/gr`;
`57,5rb/gr` → `56rb/gr`; `-` tetap `-`). Sel harga total `jt` dikurangi
`step × berat_gr ÷ 1000` (mis. ANTAM 500 gr, step 1.5 → turun 0,75 jt).

## Alur kerja yang disarankan

1. Konfirmasi ke pengguna: **tanggal baru** + arah/nominal penyesuaian (naik/turun, rb/gr).
2. Jalankan `--dry-run` → periksa daftar perubahan (semua merek kena, tanggal benar).
3. Jalankan tanpa `--dry-run` → backup otomatis dibuat (`*_BACKUP_*.xlsx`).
4. **Verifikasi wajib**: buka/cek blok baru — nilai benar, merged cells tersalin,
   tinggi baris & style sama dengan blok sumber (bukan hanya nilainya).
5. Sinkronkan database: `python sync_perak_excel_to_db.py` (tutup Excel dulu).
6. Bersihkan temp + git commit & push (lihat di bawah).

> **⚠️ Pelajaran format (jangan terulang):** openpyxl TIDAK menyalin style
> otomatis saat mengisi `.value`. Blok hasil harus **plek ketiplek** — script ini
> memakai `copy_style()` (font/border/fill/number_format/alignment) + duplikasi
> **merged cells** (offset baris) + **tinggi baris**. Verifikasi Wajib: bandingkan
> border & merged cells blok baru vs blok sebelumnya — bukan hanya nilainya. Ciri
> blok yang salah = sel tanpa border / `number_format General` / tidak ada merged
> cells / tinggi baris beda.

## Aturan keselamatan

- Selalu mulai dengan `--dry-run`.
- Backup otomatis DIHIDUPKAN (jangan `--no-backup` tanpa alasan).
- Konfirmasi tanggal & nominal ke pengguna bila tidak jelas.
- Jangan ubah format (font/size/warna/background/merge) selain nilai harga/tanggal.
- JANGAN menebak tabel; kalau tabel pengguna ambigu, tanyakan.

## Setelah selesai (pembersihan WAJIB)

- Hapus file sementara (JSON temp, `TEST_*`, salinan uji).
- **JANGAN hapus `*_BACKUP_*.xlsx` tanpa izin pengguna** — itu jaring pengaman.
  Ingatkan pengguna tentang file backup yang baru dibuat.
- Prinsip "selesai = bersih": setelah selesai hanya boleh tersisa hasil update
  (+backup bila dipertahankan) dan file skill/script.

## Git (commit & push setelah setiap update)

`git` ada di `C:\Program Files\Git\cmd\git.exe` (tidak di PATH).

```powershell
cd "D:\TOKO MAS LIA\DOKUMEN"
$git = "C:\Program Files\Git\cmd\git.exe"
& $git add "TEMPLATE HARGA PERAK.xlsx" scripts/update_harga_perak.py sync_perak_excel_to_db.py harga_perhiasan.db
& $git commit -m "update harga perak <tanggal> - <ringkas, mis. semua merk turun 1.5rb/gr>"
& $git push origin main
```

- Commit hanya file yang terkait perak untuk menghindari tercampur dengan
  perubahan lain yang bukan bagian tugas.
- Jika push ditolak karena repo privat/autentikasi (`Repository not found`),
  push dilakukan manual oleh pengguna (Git Credential Manager), mis. lewat
  `save_and_push.bat`.

## Tips eksekusi

- JANGAN jalankan verifikasi lewat `python -c "..."` di PowerShell — tanda kutip
  dalam string ikut di-parse shell dan memunculkan `NameError`/`SyntaxError`.
  Tulis selalu file `.py` sementara, jalankan `python file.py`, lalu hapus.
- Siapkan tanggal persis sesuai pola file (`TANGGAL 13 AGUSTUS 2026`).
- Update perak TIDAK berhubungan dengan `TEMPLATE HARGA MAS2.xlsx` (emas) —
  jangan keliru target.

