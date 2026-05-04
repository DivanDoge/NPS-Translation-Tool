@echo off
REM Build NPS Translator into a standalone Windows .exe using PyInstaller
REM Requirements: Python 3.x and PyInstaller installed (`pip install pyinstaller`)
REM
REM NOTE: PyInstaller does NOT support cross-compilation.
REM   - For Linux: run  build_linux.sh  on a Linux machine
REM   - For macOS: run  build_mac.sh    on a macOS machine

set SCRIPT_NAME=NPSTranslationTool.py
set EXE_NAME=NPSTranslationTool

echo Closing running instance (if any)...
taskkill /IM "%EXE_NAME%.exe" /F >nul 2>&1
timeout /t 1 /nobreak >nul

echo Building %EXE_NAME%.exe for Windows...

py -3 -m PyInstaller --noconfirm --onefile --windowed --name "%EXE_NAME%" --icon "Locus-logo.ico" --add-data "Locus-logo.png;." --add-data "saya-pic.png;." "%SCRIPT_NAME%"

echo.
echo Done. The executable is in the "dist" folder as %EXE_NAME%.exe
pause

