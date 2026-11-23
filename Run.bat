@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ====================================================================
echo               MENDELEY CITATION AUTOMATION (UNIVERSAL)
echo ====================================================================
echo.
echo Petunjuk:
echo   1. Pastikan Anda sudah meletakkan 1 file .docx dan 1 file .xlsx
echo      di dalam folder ini.
echo   2. Script akan otomatis mendeteksi format placeholder sitasi,
echo      mencocokkan ke Mendeley, dan menghasilkan:
echo        - "manuscript - Mendeley.docx"
echo        - "laporan_sitasi.xlsx"
echo.
echo ====================================================================
echo.

set "PY="
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set "PY=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
) else (
    where py >nul 2>&1
    if not errorlevel 1 (
        set "PY=py"
    ) else (
        set "PY=python"
    )
)

"%PY%" mendeley_cite.py %*
if errorlevel 1 (
    echo.
    echo [PERINGATAN] Terjadi kesalahan saat memproses dokumen.
    echo Silakan periksa pesan di atas.
)

echo.
pause
