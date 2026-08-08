@echo off
title Dokumen Tools - Kembalikan File
cls
echo ================================================================
echo   KEMBALIKAN FILE KE VERSI SEBELUMNYA
echo   Repository: TOKO MAS LIA Dokumen
echo ================================================================
echo.

set /p filename="Masukkan nama file yang ingin dikembalikan: "
if "%filename%"=="" (
    echo Nama file tidak boleh kosong.
    pause
    exit /b 1
)

echo.
echo -- Riwayat versi untuk %filename% --
git -C "D:\TOKO MAS LIA\DOKUMEN" log --oneline --follow -- "%filename%"
echo.

set /p commit_hash="Masukkan hash commit yang ingin dikembalikan: "
if "%commit_hash%"=="" (
    echo Hash commit tidak boleh kosong.
    pause
    exit /b 1
)

echo.
echo PERINGATAN: File %filename% akan dikembalikan ke versi commit %commit_hash%
echo Perubahan saat ini yang belum di-commit akan hilang!
choice /c YN /m "Apakah yakin ingin melanjutkan"
if "%errorlevel%" equ "2" exit /b 0

git -C "D:\TOKO MAS LIA\DOKUMEN" checkout "%commit_hash%" -- "%filename%"
if "%errorlevel%" equ "0" (
    echo.
    echo [BERHASIL] File berhasil dikembalikan!
    echo Jalankan save_and_push.bat untuk menyimpan perubahan ini ke GitHub.
) else (
    echo.
    echo [ERROR] Gagal mengembalikan file. Periksa hash commit dan nama file.
)

echo.
pause