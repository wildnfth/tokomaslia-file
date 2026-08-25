---
name: update-harga-mas
description: "Use when updating Toko Mas Lia gold prices in TEMPLATE HARGA MAS2.xlsx — LM/logam mulia on sheet ANTAM (new daily block) or jewelry/perhiasan on sheet EMAS (replace numbers and date). Triggers: update harga, naikkan harga, tambah blok, ganti angka emas, LM naik/turun per gram."
---

# Update Harga Mas (LM & Perhiasan)

Skill ini mengelola `TEMPLATE HARGA MAS2.xlsx` (sheet **ANTAM** = logam mulia/LM,
sheet **EMAS** = perhiasan). Ada DUA jenis update yang berbeda.

**Jalan script HANYA dari** `D:\TOKO MAS LIA\DOKUMEN\scripts\` — jangan pakai
salinan di folder skill (bisa versi lama).

| Jenis | Sheet | Sifat | Script |
|-------|-------|-------|--------|
| LM (logam mulia) | ANTAM | TAMBAH blok harian baru | `scripts/update_harga.py` |
| Perhiasan | EMAS | GANTI angka+tanggal di blok yang ada | `scripts/update_harga.py` |

> **⚠️ SCRIPT BARU (24 AGUSTUS 2026) — `update_harga.py` menggantikan 2 script lama.**
> Satu script menangani SEMUA: ANTAM `add`, `update` (merk/gramasi tunggal),
> dan perhiasan EMAS (`--json`). Otomatis: backup, buka Excel, bulatkan kelipatan 5,
> commit & push git. PELAJARAN sebelumnya tentang "LM = semua merk", harga kelipatan
> 5, dan format blok tetap berlaku (lihat catatan di bawah).

Penting: **jangan keliru target**.
- ANTAM = LM, langganan hariannya di Tambah Blok Baru.
- EMAS = perhiasan, TIDAK menambah blok (hanya perbarui angka & tanggal yang ada).

> **⚠️ FORMAT BLOK ANTAM BARU (sejak 24 AGUSTUS 2026) — header 2 baris.**
> Blok ANTAM sekarang punya header DUA baris (berbeda dari blok lama 1 baris):
>   - Baris judul : `HARGA ANTAM` (kolom A), `HARGA Galeri24` (G), `HARGA UBS BATIK` (K)
>   - Baris tanggal: `25 AGUSTUS 2026` di kolom A, G, K
>   - Baris label: `RETRO / RANDOM / 2025 / 2026` (A–E) + Galeri24 gr1 + UBS gr0.5
>   - Baris berikutnya: harga ANTAM mulai (gram di kolom A), Galeri24 & UBS lanjut.
> Script `update_harga.py --mode add` MENYALIN UTUH blok terakhir (nilai + style +
> merged cells + tinggi baris) lalu ganti tanggal, sehingga hasil PLEK KETIPLEK.

> **⚠️ PELAJARAN (15 AGUSTUS 2026) — "LM" = SEMUA MERK, bukan cuma ANTAM.**
> Saat user bilang "harga LM naik X/g" atau "naikin harga logam mulia", yang
> dimaksud adalah **semua merk logam mulia**: ANTAM (RETRO/RANDOM/2025/2026) +
> **Galeri24** + **UBS Batik**. TARGET ANTAM SAJA SALAH.
> - `--targets antam` (default) HANYA kolom B/C/D/E — Galeri24 & UBS tertinggal.
> - JANGAN pakai `update_harga_antam.py` (script lama). Satu perintah dari
>   `D:\TOKO MAS LIA\DOKUMEN`:
>   `python scripts/update_harga.py "<file>" --date "<TANGGAL>" --step X --targets allam`

> **⚠️ PELAJARAN (15 AGUSTUS 2026) — harga WAJIB kelipatan 5.**
> Setelah menaikkan harga, selisih `step * gram` bisa menghasilkan angka tak bulat
> kelipatan 5 (mis. base 1297 + 5 = 1302, base 1382 + 5 = 1387). PELANGGAN TIDAK
> SUKA harga bukan kelipatan 5. Selalu **bulatkan ke kelipatan 5 terdekat**
> (`round(v/5)*5`) pada semua sel harga blok baru setelah update.

---

## Mode 1 — Update LM (sheet ANTAM)

Script `update_harga.py` punya **tiga mode** (`--mode`):

| Mode | Fungsi |
|------|--------|
| `add` (default) | TAMBAH blok harian baru: menyalin utuh blok terakhir (nilai+style+merge+tinggi baris), ganti tanggal, terapkan step. |
| `update` | MENGUBAH harga blok TERAKHIR yang SUDAH ADA (di tempat, tanpa blok baru) — untuk menyesuaikan kolom tertentu tanpa blok harian. |
| `json` | Update sheet EMAS (perhiasan) dari file JSON. Otomatis ke sheet EMAS. |

**Target selektif (`--targets`)** — kolom mana yang dinaikkan; selain target TIDAK diubah:
- `allam` / `all` / `semua` = SEMUA merk: ANTAM (RETRO/RANDOM/2025/2026) + Galeri24 + UBS
- `antam` (default) = semua kolom ANTAM (RETRO/RANDOM/2025/2026)
- `retro`, `random`, `2025`, `2026`, `galeri24`, `ubs`
- Bisa gabung koma: `--targets "2025,ubs"`

**Filter gramasi (`--grams`)** — hanya ubah baris khusus: `--grams "1,2,5"`.

Contoh (langkah pertama selalu `--dry-run`):
```powershell
# "LM naik X/gr SEMUA merk" — SATU perintah (ANTAM + Galeri24 + UBS):
python scripts/update_harga.py "<file.xlsx>" --date "25 AGUSTUS 2026" --step 10 --targets allam
# ANTAM saja naik 50/gr:
python scripts/update_harga.py "<file.xlsx>" --date "25 AGUSTUS 2026" --step 50
# ubah blok terakhir DI TEMPAT: cuma UBS naik 30/gr:
python scripts/update_harga.py "<file.xlsx>" --mode update --step 30 --targets ubs
# cuma Galeri24 naik 25/gr:
python scripts/update_harga.py "<file.xlsx>" --mode update --step 25 --targets galeri24
# ANTAM 2025 gramasi 1 & 2 naik 40/gr:
python scripts/update_harga.py "<file.xlsx>" --mode update --step 40 --targets 2025 --grams "1,2"
```
- `--date` : label tanggal baru (WAJIB utk mode `add`; utk `update`/`json` abaikan).
- `--step`: kenaikan per gram (boleh negatif untuk turun; default 50).
- Struktur blok: HARGA ANTAM (A–E, gram di A), Galeri24 (G–H, gram di G),
  UBS Batik (K–L, gram di K). Baris dengan `MERAH = KOSONG` (kolom G) adalah
  penanda akhir bagian ANTAM/Galeri24.
- `--mode update` tidak menambah blok → **tidak perlu `--date`**, tidak mengubah header/tanggal.
- **Otomatis** setelah sukses (kecuali `--no-open` / `--no-git` / `--dry-run`):
  backup, buka file di Excel, commit & push ke git.

> **⚠️ FORMAT BARU (sejak blok 16 AGUSTUS 2026) — UBS memanjang melewati `MERAH = KOSONG`.**
> Blok yang sekarang (16 AGUSTUS 2026) lebih tinggi 12 baris daripada blok-blok
> lama:
> - **UBS (kolom K–L)** kini punya gramasi `0.5, 1, 2, 3, 4, 5, 10, 25, 50, 100`
>   — ada baris **"4 gram"** (yang TIDAK dimiliki ANTAM/Galeri24), dan dua baris
>   bawahnya (**50 & 100 gram**) berada di **BAWAH** baris `MERAH = KOSONG`.
> - **ANTAM & Galeri24** tetap `0.5, 1, 2, 3, 5, 10, 25, 50, 100` (tanpa "4 g").
> Akibatnya jangan berasumsi "baris terakhir blok = baris `MERAH = KOSONG`".
> Script sudah disesuaikan agar:
>   1. **`add`** menempatkan header blok baru di **`baris_terakhir_blok_sumber + 2`**
>      (1 baris kosong pemisah), bukan dari selisih antar-header yang bisa lebih
>      kecil dari tinggi blok.
>   2. **`update`** memproses **seluruh baris data blok terakhir** (termasuk UBS
>      50/100 g di bawah `MERAH = KOSONG`), baris tanpa nilai target otomatis
>      dilewati.
> Verifikasi wajib setelah `add`: bandingkan **merged cells** blok baru vs blok
> sumber (offset baris harus sama untuk semua, termasuk `L:M` baris UBS 50/100)
> dan borders/number_format — ciri blok salah = sel tanpa border / `General` /
> merged cells hilang.

> **⚠️ PELAJARAN (jangan terulang 2x) — Format blok ANTAM WAJIB ikut tersalin.**
> `openpyxl` TIDAK menyalin style otomatis saat menulis `.value`. Di mode `add`,
> blok hasil harus **plek ketiplek** dengan blok sumber, jadi WAJIB:
> 1. Menyalin style sel lewat `copy_style(src, dst)` (font/border/fill/
>    number_format/alignment) — bukan hanya nilainya.
> 2. Menduplikasi **merged cells** blok sumber ke blok baru (offset baris).
> 3. Menyalin **tinggi baris** (`row_dimensions.height`).
>
> **Penyebab bug yang pernah terjadi:** ada versi lama script yang hanya menyalin
> nilai sehingga blok copas kehilangan border + format angka (`General`) + merged
> cells + tinggi baris (terlihat polos/kosong). Setelah update, VERIFIKASI WAJIB:
> bandingkan border & merged cells blok baru vs blok sebelumnya — bukan hanya
> nilainya. Ciri blok yang salah = sel tanpa border, `number_format = General`,
> tidak ada merged cells.

## Setelah update berhasil — kirim Discord

**REQUIRED SUB-SKILL:** `kirim-harga-discord`

Jangan tulis screenshot/webhook sendiri. Setelah nilai + format terverifikasi:

```powershell
cd "D:\TOKO MAS LIA\DOKUMEN"
python scripts/kirim_harga_discord.py --channel harga-lm-emas      # setelah update LM
python scripts/kirim_harga_discord.py --channel harga-perhiasan    # setelah update EMAS
```

Jangan kirim saat `--dry-run`. Jangan buka Excel sebelum kirim.
Kalau `CopyPicture method of Range class failed`: jalankan perintah resmi
yang sama 1x lagi (lihat `kirim-harga-discord`). Jangan tulis screenshot sendiri.

---

## Mode 2 — Update Perhiasan (sheet EMAS): ganti angka & tanggal

TIDAK menambah blok. Minta pengguna nilai terbaru (biasanya tabel), susun
ke file JSON, lalu jalankan script. Format sel dipertahankan (angka tunggal
atau string `"X / Y"`).

Langkah:
1. Konfirmasi ke pengguna: tanggal baru + angka per kadar.
2. Susun file JSON (contoh `perhiasan_baru.json`):
```json
{
  "date": "6 AGUSTUS 2026",
  "kuning": { "300/6K": "825 / 850", "450/10K": 1275 },
  "putih":  { "300/6K": 850, "18K": 2125 }
}
```
   - `kuning` -> kolom B & F (dua bagian kembar di blok).
   - `putih`  -> kolom D & H (hanya kadar yang ada, biasanya 6K–18K).
   - Kadar = label di kolom A (`300/6K`, `9C PABRIK`, `9A/916/21K`, ...).
   - Konversi rupiah ke ribuan: `Rp825.000` -> `825`.

> **Struktur empat kolom sheet EMAS (pelajaran penting):**
> - **B & F** = kolom KUNING/ROSE GOLD; bernilai **string `"X / Y"`** bila kadar
>   punya ROSE (mis. `"820 / 840"`), atau **angka tunggal** bila tidak (10K, 18K,
>   20K ke atas = `1265`, `2060`, `2190`, ...).
> - **D & H** = kolom PUTIH; selalu **angka tunggal**, dan hanya terisi untuk
>   6K–18K (baris 3–9). 20K ke atas TIDAK punya sel PUTIH — JANGAN isi.
> Jadi saat pengguna memberi tabel berkolom "KUNING" + "ROSE GOLD", satukan ke
> string `"KUNING / ROSE"` untuk kolom B/F. Jangan mengisi PUTIH untuk kadar
> 20K ke atas.
>
> **Aturan harga PUTIH vs ROSE GOLD (penting):** umumnya harga PUTIH **≥ ROSE
> GOLD** (putih sama atau lebih mahal dari rose gold). Bila nilai PUTIH lebih
> kecil dari ROSE GOLD untuk kadar yang sama, TANYAKAN dulu — KECUALI pada
> **16K PUTIH**: ini pengecualian resmi. **16K PUTIH tidak ada yang dari pabrik**
> (aslinya emas 16K biasa yang dilapisi/di-*rhodium plating* jadi putih),
> sehingga wajar harganya **lebih murah daripada 16K ROSE GOLD pabrik**. Jadi
> untuk 16K jangan lagi mempertanyakan/bertanya; isi mengikuti angka pengguna.
>
> **Verifikasi label sebelum update (agar tak ada MISSING):**
> `--dry-run` update_harga_emas HANYA menampilkan jumlah kadar, BUKAN label yang
> cocok/tidak. Sebelum menjalankan, cek label kolom A persis cocok dengan kunci
> JSON (loop & bandingkan), baru lanjut ke update sungguhan.

3. Jalankan:
```powershell
python scripts/update_harga.py "<file.xlsx>" --json "perhiasan_baru.json"
python scripts/update_harga.py "<file.xlsx>" --json "perhiasan_baru.json" --dry-run
```
   - Tanggal ditulis di A1 & E1; **POT tidak diubah** (tidak ada di daftar).

---

## Pembersihan setelah selesai (WAJIB)

Setelah update BERHASIL dan tervalidasi (file terbuka normal, nilai benar),
lakukan hal berikut SEBELUM menutup tugas:

1. **Hapus semua file sementara** yang dibuat selama proses:
   - file JSON nilai sementara di folder temp (mis. `perhiasan_baru.json`),
   - file salinan uji apa pun (`TEST_*`, `*_test*`),
   - script/temp satu kali pakai yang sempat dibuat di folder kerja.
   JANGAN tinggalkan sampah di folder kerja pengguna.
2. **JANGAN hapus file `*_BACKUP_*.xlsx`** tanpa izin pengguna — itu jaring
   pengaman. 
3. **Ingatkan pengguna** dengan menyebutkan:
   - file mana yang diubah,
   - nama file backup yang baru dibuat (dan minta konfirmasi apakah backup lama
     mau dihapus/ dipertahankan).

Gunakan prinsip "selesai = bersih": begitu selesai, tidak boleh ada file
baru yang tertinggal selain hasil update dan backup-nya.

---

## Git (commit & push setelah setiap update)

Setelah update berhasil, validasi OK, dan pembersihan temp selesai, jalankan
git untuk mencatat perubahan ke GitHub. Di environment ini `git` terpasang di
`C:\Program Files\Git\cmd\git.exe` (tidak ada di PATH), jadi gunakan path
eksplisit.

```powershell
cd "D:\TOKO MAS LIA\DOKUMEN"
$git = "C:\Program Files\Git\cmd\git.exe"
& $git add -A
& $git commit -m "update harga <tanggal> - <ringkas perubahan, mis. UBS +30, EMAS 7 AGUSTUS 2026>"
& $git push origin main
```

- `git add -A` menambah semua perubahan file (Excel + skill jika ada yang diubah).
- Pesan commit harus **singkat dan jelas**.
- Jika push ditolak karena branch behind, tarik dulu (`git pull --rebase`) lalu push ulang.

---

## Tips eksekusi (pembelajaran agar lebih cepat)
- **JANGAN jalankan kode inspeksi/verifikasi lewat `python -c`.** Di PowerShell,
  tanda kutip di dalam string (mis. `f = "TEMPLATE HARGA MAS2.xlsx"`) ikut
  di-parse oleh shell dan memunculkan `NameError`. Tulis selalu sebagai file
  `.py` sementara, jalankan `python file.py`, lalu hapus.
- Tulis script inspeksi SATU kali yang mencetak: sheet, baris header, label
  kolom A, nilai kolom B/D/F/H saat ini — untuk ANTAM dan EMAS — agar langsung
  tahu struktur & label tanpa perulangan.
- Konfirmasi ke pengguna bila tabelnya ambigu (mis. "hanya 16K/17K yang berubah"
  padahal tabel menurunkan semua kadar). Jangan menebak — tanyakan.
- Siapkan JSON langsung dengan label persis tertera di kolom A file (bukan
  penamaan karat "6K" jika di file tertulis "300/6K").
- **Mode `add` ANTAM = copas FORMAT lengkap, bukan sekadar nilai.** Pastikan
  script yang benar-benar DIEKSEKUSI memuat `copy_style(src, dst)` + penyalinan
  merged cells & tinggi baris. Jangan pakai script yang TIDAK memiliki
  `copy_style` (versi lama) — hasil bloknya kehilangan border/angka/merge.
- **Script yang dijalankan WAJIB** `D:\TOKO MAS LIA\DOKUMEN\scripts\update_harga.py`.
  Jangan `update_harga_antam.py` dan jangan salinan di folder skill/`.clinerules`.

## Aturan keselamatan
- Selalu mulai dengan `--dry-run`, lalu jalankan tanpa flag itu.
- Backup otomatis DIHIDUPKAN (jangan `--no-backup` tanpa alasan).
- Konfirmasi tanggal & nominal ke pengguna bila tidak jelas.
- Jangan ubah format (font/size/warna/background/merge) selain nilai yang diminta.
- **SETELAH selesai edit, BUKA filenya di Excel laptop pengguna**
  (`Start-Process "<path.xlsx>"`) supaya user langsung bisa cek hasilnya.
- **Harga wajib kelipatan 5.** Setelah update harga apa pun, periksa semua sel
  harga; yang bukan kelipatan 5 dibulatkan ke kelipatan 5 terdekat.
- **Setelah verifikasi, kirim Discord** lewat `kirim_harga_discord.py` (lihat
  skill `kirim-harga-discord`). Jangan `CopyPicture`/`Chart.Paste` manual.

---

## Git & Backup (Update Penting)

Sistem Git + Git LFS sudah terpasang dan dikonfigurasi untuk melacat semua file dokumen di folder ini. Ini berarti:

### Kemampuan Sistem
- **Riwayat penuh**: Setiap perubahan pada file terlacat otomatis
- **Audit trail**: Dapat melihat siapa mengubah apa kapan saja
- **Restore fleksibel**: Dapat kembali ke versi file lama kapan saja
- **Backup otomatis**: Semua perubahan di-push ke GitHub (private repo)

### Repository Status
- **GitHub**: https://github.com/liagoldeditor/tokomaslia-file
- **Lokasi lokal**: D:\\TOKO MAS LIA\\DOKUMEN\\
- **Branch**: main
- **File**: Semua file .docx, .xlsx, .pptx dilacat dengan Git LFS

### Alur Simpan yang Direkomendasikan

Setelah memodifikasi file, gunakan salah satu cara ini:

**Cara 1 - Manual (PowerShell):**
```powershell
cd "D:\TOKO MAS LIA\DOKUMEN"
$git = "C:\Program Files\Git\cmd\git.exe"
& $git add -A
& $git commit -m "update harga [tanggal] - [deskripsi singkat]"
& $git push origin main
```

**Cara 2 - Script otomatis (lebih mudah):**
```
Klik dobel "D:\TOKO MAS LIA\DOKUMEN\save_and_push.bat"
```

### Melihat Riwayat Perubahan
```bash
# Lihat semua commit
git log --oneline --graph --all

# Lihat riwayat file tertentu
git log --oneline --follow "TEMPLATE HARGA MAS2.xlsx"

# Lihat detail perubahan dalam commit tertentu
git show [commit-hash] --stat
```

### Mengembalikan File ke Versi Lama
```bash
# Cari commit yang diinginkan
git log --oneline --follow "TEMPLATE HARGA MAS2.xlsx"

# Kembalikan file ke versi tersebut
git checkout [commit-hash] -- "TEMPLATE HARGA MAS2.xlsx"
```

Atau gunakan script interaktif:
```
Klik dobel "D:\TOKO MAS LIA\DOKUMEN\dokumen-tools.bat"
Pilih menu 2 untuk riwayat, menu 3 untuk restore
```

### File Pendukung
| File | Fungsi |
|------|--------|
| `dokumen-tools.bat` | Menu utama semua operasi Git |
| `save_and_push.bat` | Simpan & push cepat |
| `view_history.bat` | Lihat riwayat perubahan |
| `restore_file.bat` | Kembalikan file lama |
| `PANDUAN_PENGGUNAAN.txt` | Panduan penggunaan |

### Catatan Penting
- `TEMPLATE HARGA MAS2.xlsx` sudah dilacat oleh Git LFS, perubahan tidak akan tersambar lagi
- File `*_BACKUP_*.xlsx` juga ikut dilacat Git secara otomatis
- Setiap ekseskusi skill ini, perubahan akan tersedia di riwayat GitHub
- Tidak perlu lagi khawatir kehilangan versi lama dokumen


