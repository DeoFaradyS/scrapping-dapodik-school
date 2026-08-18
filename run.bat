@echo off
cd /d "%~dp0"
python scrape_api.py
echo.
echo ================================
echo Selesai / berhenti. Tekan tombol apa aja buat nutup jendela ini.
pause >nul
