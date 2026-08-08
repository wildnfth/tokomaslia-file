@echo off
title Dokumen Tools - Lihat Riwayat
cls
echo ================================================================
echo   RIWAYAT PERUBAHAN DOKUMEN
echo   Repository: TOKO MAS LIA Dokumen
echo ================================================================
echo.

echo -- Riwayat commit terakhir --
git -C "D:\TOKO MAS LIA\DOKUMEN" log --oneline --graph --all -10
echo.

echo -- File yang sedang dilacat --
git -C "D:\TOKO MAS LIA\DOKUMEN" ls-files | findstr /i "\.docx\|\.xlsx\|\.pptx"
echo.

echo -- Riwayat perubahan untuk file spesifik --
set /p filename="Nama file (mis: ABSENSI TOKO.xlsx) atau kosongkan untuk kembali: "
if "%filename%"=="" exit /b 0

echo.
echo Riwayat untuk %filename%:
git -C "D:\TOKO MAS LIA\DOKUMEN" log --oneline --follow -- "%filename%"
echo.

pause