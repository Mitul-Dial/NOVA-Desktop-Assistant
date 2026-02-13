@echo off
title NOVA - Build Script
echo.
echo  ============================================
echo     N O V A  -  Desktop Assistant Builder
echo  ============================================
echo.

REM ── Check Python ──
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python is not installed or not in PATH.
    echo  Please install Python 3.10+ from https://python.org
    echo.
    pause
    exit /b 1
)

echo  [1/4] Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo  [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)
echo.

echo  [2/4] Generating app icon...
python generate_icon.py
echo.

echo  [3/4] Building NOVA.exe (this may take a minute)...
pyinstaller --noconfirm ^
    --onefile ^
    --windowed ^
    --name "NOVA" ^
    --icon "nova.ico" ^
    --add-data "nova.ico;." ^
    --hidden-import "pyttsx3.drivers" ^
    --hidden-import "pyttsx3.drivers.sapi5" ^
    --hidden-import "win32com.client" ^
    --hidden-import "win32api" ^
    --hidden-import "win32gui" ^
    --hidden-import "win32con" ^
    --hidden-import "pythoncom" ^
    --hidden-import "comtypes" ^
    --hidden-import "customtkinter" ^
    --hidden-import "speech_recognition" ^
    "NOVA Desktop Assistant.py"

if errorlevel 1 (
    echo  [ERROR] Build failed.
    pause
    exit /b 1
)
echo.

echo  [4/4] Creating Desktop shortcut...
powershell -Command ^
    "$ws = New-Object -ComObject WScript.Shell; " ^
    "$sc = $ws.CreateShortcut([System.IO.Path]::Combine($env:USERPROFILE, 'Desktop', 'NOVA.lnk')); " ^
    "$sc.TargetPath = (Resolve-Path 'dist\NOVA.exe').Path; " ^
    "$sc.IconLocation = (Resolve-Path 'nova.ico').Path; " ^
    "$sc.WorkingDirectory = (Resolve-Path 'dist').Path; " ^
    "$sc.Description = 'NOVA - Voice Assistant'; " ^
    "$sc.Save()"
echo.

echo  ============================================
echo     BUILD COMPLETE!
echo  ============================================
echo.
echo   NOVA.exe is at:  dist\NOVA.exe
echo   Shortcut placed on your Desktop.
echo.
echo   Double-click "NOVA" on your Desktop to start!
echo.
pause
