---
name: kirim-harga-discord
description: "Use when sending Toko Mas Lia gold, silver, or jewelry price tables to Discord; when a Discord harga screenshot is blank, white, or empty; when CopyPicture method of Range class failed / pywintypes.com_error / 0x800A03EC on screenshot; when posting harga LM emas, LM perak, or perhiasan after an update; or when adding or fixing a Discord webhook for those channels."
---

# Kirim Harga Discord

Satu pintu kirim foto tabel harga ke Discord. Agent lain WAJIB pakai script
ini — jangan menulis ulang screenshot/webhook.

**Repo:** `D:\TOKO MAS LIA\DOKUMEN`
**Script:** `python scripts/kirim_harga_discord.py`

## Channel

| Channel | File | Sheet | Isi |
|---------|------|-------|-----|
| `harga-lm-emas` | `TEMPLATE HARGA MAS2.xlsx` | ANTAM | blok LM terakhir (ANTAM+Galeri24+UBS) |
| `harga-lm-perak` | `TEMPLATE HARGA PERAK.xlsx` | Sheet1 | blok perak terakhir |
| `harga-perhiasan` | `TEMPLATE HARGA MAS2.xlsx` | EMAS | tabel perhiasan |

Webhook ada di `discord_config.json` (sudah di-`.gitignore`). **Jangan
cetak/commit URL webhook.**

## Perintah (hanya ini)

```powershell
cd "D:\TOKO MAS LIA\DOKUMEN"
python scripts/kirim_harga_discord.py --channel harga-lm-emas
python scripts/kirim_harga_discord.py --channel harga-lm-perak
python scripts/kirim_harga_discord.py --channel harga-perhiasan
python scripts/kirim_harga_discord.py --channel all
```

Sukses = baris `[screenshot] range ...` + `[discord] status 200 -> OK`.
PNG temp dihapus otomatis.

Setelah **update harga berhasil + terverifikasi**, kirim channel yang relevan
(LM emas / LM perak / perhiasan). Jangan kirim saat `--dry-run`.

## Larangan (penyebab PNG kosong — sudah terjadi)

| Jangan | Kenapa | Ganti dengan |
|--------|--------|--------------|
| Tulis `CopyPicture` / `Chart.Paste` / `ImageGrab` sendiri | `CopyPicture(1,1)` clipboard kosong → PNG putih | `kirim_harga_discord.py` |
| Foto `UsedRange` sheet ANTAM | 1400+ baris riwayat; Discord terlihat kosong | script ambil **blok terakhir** |
| `os.startfile` / buka Excel **sebelum** screenshot | file lock → PNG kosong | kirim dulu, baru buka Excel |
| Range ANTAM hanya sampai kolom L | header/harga UBS merge `K:M` terpotong | script pakai sampai kolom M |
| Pakai `update_harga_antam.py` / salinan di folder skill | script lama, drift | hanya `D:\TOKO MAS LIA\DOKUMEN\scripts\` |
| `python -c "..."` di PowerShell | kutip rusak, `NameError` | script `.py` lalu hapus |
| Anggap `CopyPicture` COM error = gagal total / tulis script baru | instance Excel pertama sering belum siap | perintah resmi yang **sama** 1x lagi |

> **⚠️ PELAJARAN (25 AGUSTUS 2026) — `CopyPicture method of Range class failed`.**
> Gejala: range sudah benar (`ANTAM!$A$1461:$M$1472`) lalu crash
> `pywintypes.com_error` / `-2146827284` / `0x800A03EC`. Bukan file rusak,
> bukan webhook, bukan range salah. Instance Excel COM pertama (`Visible=False`)
> kadang belum siap; retry perintah resmi yang sama langsung sukses (PNG valid).
> File `~$TEMPLATE…xlsx` berumur bulan = sampah, bukan lock aktif.
>
> Yang benar:
> 1. Script resmi sudah retry `CopyPicture` 3x + fallback Chart + 1x instance baru.
> 2. Kalau perintah masih gagal: jalankan **perintah resmi yang sama** sekali lagi.
> 3. Jangan tulis `CopyPicture` / `Chart.Paste` / `ImageGrab` sendiri.
> 4. Ini BEDA dari Discord HTTP bukan 200 — jangan bikin sender baru.

## Cek cepat sebelum bilang "sudah terkirim"

1. File `discord_config.json` punya ketiga channel (bukan `ISI_WEBHOOK_*`).
2. Output ada `range ANTAM!$A$…` (bukan ribuan baris dari atas).
3. Status Discord `200`. Kalau bukan 200: jangan ulangi dengan kode baru.
4. `CopyPicture` COM error: retry perintah resmi 1x dulu, baru laporkan gagal.
5. Jangan tinggalkan `*.png` / `TEST_*` / JSON temp di folder kerja.

## Tambah webhook

Isi `webhook_url` di `discord_config.json` untuk key channel di tabel atas.
Jangan commit file itu. Contoh bentuk: `discord_config.example.json`.
