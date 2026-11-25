@echo off
setlocal
cd /d "%~dp0"

if not exist "%~dp0install.ps1" (
    echo.
    echo  File install.ps1 tidak ditemukan di folder ini:
    echo  %~dp0
    echo.
    echo  Ekstrak SELURUH isi ZIP Dockit dulu ^(klik kanan file ZIP, pilih "Extract All"^),
    echo  lalu jalankan Install.bat dari folder hasil ekstrak.
    echo.
    echo  Tidak punya Python? Download versi siap pakai ^(Dockit.exe^):
    echo  https://github.com/Augustian17/litek-kelomok-5/releases/latest
    echo.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
pause
