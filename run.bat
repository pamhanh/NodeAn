@echo off
setlocal
cd /d "%~dp0"

set PYCMD=

where python >nul 2>nul
if not errorlevel 1 (
    set PYCMD=python
) else (
    where py >nul 2>nul
    if not errorlevel 1 (
        set PYCMD=py -3
    )
)

if "%PYCMD%"=="" (
    echo [Notes Overlay] Khong tim thay Python tren may nay.
    echo Vui long cai Python 3.8+ tu https://python.org
    echo Khi cai, nho tick vao o "Add python.exe to PATH", roi chay lai file run.bat nay.
    echo.
    pause
    exit /b 1
)

echo [Notes Overlay] Dang dung: %PYCMD%
echo [Notes Overlay] Dang kiem tra/cai thu vien can thiet ^(keyboard, Pillow^)...
%PYCMD% -m pip install -r requirements.txt --quiet --disable-pip-version-check
if errorlevel 1 (
    echo.
    echo [Notes Overlay] Cai thu vien that bai. Thu chay lenh sau bang tay de xem loi chi tiet:
    echo     %PYCMD% -m pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

echo [Notes Overlay] Dang khoi dong...
%PYCMD% notes_overlay.py 2>run_error.log
if errorlevel 1 (
    echo.
    echo [Notes Overlay] Chuong trinh bi loi khi chay. Noi dung loi:
    echo ----------------------------------------------------------
    type run_error.log
    echo ----------------------------------------------------------
    echo Loi cung da duoc luu vao file run_error.log trong thu muc nay.
    echo.
    pause
    exit /b 1
)

exit /b 0
