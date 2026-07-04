@echo off
REM One-click build for GW Order Tool.exe.
REM Run this ONCE on a Windows PC that has Python installed. It creates the
REM finished, self-contained GW Order Tool.exe in the "dist" folder. Copy that
REM single file to any other Windows computer (with Google Chrome installed) --
REM no Python, no setup, no other files needed there.

setlocal
cd /d "%~dp0"

echo ============================================
echo  GW Order Tool - build script
echo ============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python was not found on this PC.
    echo Install Python 3.10 or newer from https://www.python.org/downloads/
    echo ^(tick "Add python.exe to PATH" during install^), then re-run this script.
    echo.
    pause
    exit /b 1
)

echo Creating a private build environment in .build_venv ...
python -m venv .build_venv
if errorlevel 1 (
    echo [ERROR] Failed to create the virtual environment.
    pause
    exit /b 1
)

call .build_venv\Scripts\activate.bat

echo.
echo Installing build requirements (openpyxl, selenium, pyinstaller) ...
python -m pip install --upgrade pip >nul
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install requirements. Check your internet connection and try again.
    pause
    exit /b 1
)

echo.
echo Removing any previous build output ...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo Building GW Order Tool.exe ...
pyinstaller "GW Order Tool.spec" --noconfirm
if errorlevel 1 (
    echo [ERROR] Build failed - see the messages above.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Build finished successfully.
echo  Your app is at: dist\GW Order Tool.exe
echo.
echo  Copy that ONE file anywhere you like (Desktop, a shared drive, another
echo  computer). Just make sure:
echo    1. It's kept in a normal, writable folder (not Program Files) - it
echo       saves its own data files (saved_products.json, the Neto login,
echo       etc.) right next to itself.
echo    2. Google Chrome is installed on whichever PC runs it.
echo  No Python and no other files from this folder are needed anywhere else.
echo ============================================
echo.
pause
