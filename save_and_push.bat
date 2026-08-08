@echo off
title Dokumen Tools - Save & Push

:save
cls
echo ================================================================
echo   SIMPAN DAN PUSH PERUBAHAN KE GITHUB
echo   Repository: TOKO MAS LIA Dokumen
echo ================================================================
echo.

REM Cek status repository
echo [1/4] Mengecek status repository...
git -C "D:\TOKO MAS LIA\DOKUMEN" status --short
if "%errorlevel%" neq "0" (
    echo [ERROR] Gagal mengecek status. Pastikan folder ini adalah repository Git.
    pause
    exit /b 1
)

REM Membungkuk perubahan
echo [2/4] Menambahkan perubahan file...
git -C "D:\TOKO MAS LIA\DOKUMEN" add .

REM Commit perubahan
echo [3/4] Membuat commit...
set /p commit_msg="Masukkan pesan commit (mis: Update laporan penjualan Hari Rabu): "
if "%commit_msg%"=="" set commit_msg=Auto-save pada %date% %time%

git -C "D:\TOKO MAS LIA\DOKUMEN" commit -m "%commit_msg%" >nul 2>&1
if "%errorlevel%" equ "0" (
    echo Commit berhasil dibuat.
) else (
    if "%errorlevel%" equ "1" (
        echo Tidak ada perubahan untuk di-commit. Lanjut ke push...
    ) else (
        echo [ERROR] Gagal melakukan commit.
        pause
        exit /b 1
    )
)

REM Push ke GitHub
echo [4/4] Mengunggah ke GitHub...
git -C "D:\TOKO MAS LIA\DOKUMEN" push origin main
if "%errorlevel%" equ "0" (
    echo.
    echo [SELESAI] Semua perubahan berhasil disimpan dan di-upload ke GitHub!
    echo.
    
    REM Tampilkan ringkasan
    echo Ringkasan commit terbaru:
    git -C "D:\TOKO MAS LIA\DOKUMEN" log -1 --oneline
    echo.
    
    choice /c YN /m "Apakah ingin menyimpan dokumen lagi"
    if "%errorlevel%" equ "1" goto save
    echo.
    echo Terima kasih! Dokumen Anda aman di GitHub.
    timeout /t 3 >nul
) else (
    echo [ERROR] Gagal push ke GitHub. Periksa koneksi internet atau kredensial.
    echo Silakan jalankan perintah berikut secara manual untuk memecahkan masalah:
    echo git -C "D:\TOKO MAS LIA\DOKUMEN" push origin main
    pause
)

exit /b 0