@echo off
REM Build NPS Translator into a standalone Windows .exe using PyInstaller
REM Requirements: Python 3.x and PyInstaller installed (`pip install pyinstaller`)

set SCRIPT_NAME=NPSTranslationTool.py
set EXE_NAME=NPSTranslationTool

echo Building %EXE_NAME%.exe ...

py -3 -m PyInstaller --noconfirm --onefile --windowed --name "%EXE_NAME%" --icon "Locus-logo.ico" --add-data "Locus-logo.png;." "%SCRIPT_NAME%"

echo.
echo Done. The executable will be in the "dist" folder as %EXE_NAME%.exe
pause

