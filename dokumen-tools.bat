@echo off
title Dokumen Tools - TOKO MAS LIA
setlocal enabledelayedexpansion

:main
cls
echo ================================================================
echo   MANAJEMEN DOKUMEN TOKO MAS LIA
echo   Git + Git LFS + GitHub
echo ================================================================
echo.
echo Pilih aksi:
echo   1. Simpan dan push perubahan ke GitHub
echo   2. Lihat riwayat perubahan
echo   3. Kembalikan file ke versi sebelumnya
echo   4. Sinkronisasi dari GitHub (pull)
echo   5. Buat backup lokal (ZIP)
echo   6. Cek git LFS
echo   7. Keluar
echo.
set /p choice="Pilihan (1-7): "

if "%choice%"=="1" goto save
if "%choice%"=="2" goto history
if "%choice%"=="3" goto restore
if "%choice%"=="4" goto pull
if "%choice%"=="5" goto backup
if "%choice%"=="6" goto lfs_check
if "%choice%"=="7" goto end

echo Pilihan tidak valid!
timeout /t 2 >nul
goto main

:save
cls
echo ================================================================
echo   SIMPAN DAN PUSH PERUBAHAN KE GITHUB
echo ================================================================
echo.

echo [1/4] Mengecek status...
git -C "D:\TOKO MAS LIA\DOKUMEN" status --short
echo.

echo [2/4] Menambahkan perubahan...
git -C "D:\TOKO MAS LIA\DOKUMEN" add .
echo.

echo [3/4] Membuat commit...
set /p commit_msg="Masukkan pesan commit (tekan Enter untuk auto): "
if "%commit_msg%"=="" (
    for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
    set "commit_msg=Auto-save %dt:~0,4%-%dt:~4,2%-%dt:~6,2% %dt:~8,2%:%dt:~10,2%"
)

git -C "D:\TOKO MAS LIA\DOKUMEN" commit -m "%commit_msg%" 2>nul
if "%errorlevel%" equ "0" (
    echo Commit berhasil: %commit_msg%
) else (
    if "%errorlevel%" equ "1" (
        echo Tidak ada perubahan baru.
    )
)
echo.

echo [4/4] Push ke GitHub...
git -C "D:\TOKO MAS LIA\DOKUMEN" push origin main
if "%errorlevel%" equ "0" (
    echo.
    echo [SELESAI] Perubahan berhasil disimpan dan di-upload!
    echo.
    git -C "D:\TOKO MAS LIA\DOKUMEN" log -1 --oneline
) else (
    echo [ERROR] Gagal push. Periksa koneksi.
)
echo.
pause
goto main

:history
cls
echo ================================================================
echo   RIWAYAT PERUBAHAN
echo ================================================================
echo.

echo -- 10 Commit Terakhir --
git -C "D:\TOKO MAS LIA\DOKUMEN" log --oneline --graph --all -10
echo.

echo -- File yang Dilacat --
git -C "D:\TOKO MAS LIA\DOKUMEN" ls-files | findstr /i "\.docx\|\.xlsx\|\.pptx"
echo.

echo -- Riwayat Spesifik File --
set /p filename="Nama file untuk riwayat spesifik (tekan Enter untuk kembali): "
if not "%filename%"=="" (
    echo.
    echo Riwayat untuk %filename%:
    git -C "D:\TOKO MAS LIA\DOKUMEN" log --oneline --follow -- "%filename%"
    echo.
)
pause
goto main

:restore
cls
echo ================================================================
echo   KEMBALIKAN FILE KE VERSI SEBELUMNYA
echo ================================================================
echo.

set /p filename="Nama file yang ingin dikembalikan: "
if "%filename%"=="" (
    echo Nama file kosong.
    pause
    goto main
)

echo.
echo Riwayat untuk %filename%:
git -C "D:\TOKO MAS LIA\DOKUMEN" log --oneline --follow -- "%filename%"
echo.

set /p commit_hash="Hash commit tujuan: "
if "%commit_hash%"=="" (
    echo Hash commit kosong.
    pause
    goto main
)

echo.
choice /c YN /m "Yakin? File akan dikembalikan ke versi commit ini"
if "%errorlevel%" equ "2" goto main

git -C "D:\TOKO MAS LIA\DOKUMEN" checkout "%commit_hash%" -- "%filename%"
if "%errorlevel%" equ "0" (
    echo.
    echo [BERHASIL] File dikembalikan! Jalankan save untuk menyimpannya.
) else (
    echo [ERROR] Gagal mengembalikan file.
)
echo.
pause
goto main

:pull
cls
echo ================================================================
echo   SINKRONISASI DARI GITHUB
echo ================================================================
echo.
git -C "D:\TOKO MAS LIA\DOKUMEN" pull origin main
if "%errorlevel%" equ "0" (
    echo.
    echo [SELESAI] Sinkronisasi berhasil!
) else (
    echo.
    echo [PERINGATAN] Mungkin ada konflik. Periksa manual.
)
echo.
pause
goto main

:backup
cls
echo ================================================================
echo   BACKUP LOKAL (ZIP)
echo ================================================================
echo.

if not exist "D:\TOKO MAS LIA\DOKUMEN\backup" mkdir "D:\TOKO MAS LIA\DOKUMEN\backup"

for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "timestamp=%dt:~0,4%-%dt:~4,2%-%dt:~6,2%_%dt:~8,2%-%dt:~10,2%"

set "backup_name=DOKUMEN_backup_%timestamp%.zip"

git -C "D:\TOKO MAS LIA\DOKUMEN" archive --format=zip --output="backup\%backup_name%"
if "%errorlevel%" equ "0" (
    echo.
    echo [BERHASIL] Backup dibuat: backup\%backup_name%
) else (
    echo.
    echo [ERROR] Gagal membuat backup.
)
echo.
pause
goto main

:lfs_check
cls
echo ================================================================
echo   STATUS GIT LFS
echo ================================================================
echo.
git -C "D:\TOKO MAS LIA\DOKUMEN" lfs ls-files
echo.
pause
goto main

:end
cls
echo.
echo Terima kasih! Semua dokumen Anda aman di GitHub.
echo Repository: https://github.com/liagoldeditor/tokomaslia-file
timeout /t 3 >nul
exit /b 0