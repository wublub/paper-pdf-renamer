@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Installing/updating dependencies...
python -m pip install --upgrade pypdf pyinstaller
echo.
echo Building PDF_Renamer.exe ...
pyinstaller --onefile --windowed --noconfirm --clean --name PDF_Renamer pdf_renamer.py
echo.
if exist "dist\PDF_Renamer.exe" (
    copy /Y "dist\PDF_Renamer.exe" "PDF_Renamer.exe" >nul
    echo Done. exe is at: %CD%\PDF_Renamer.exe
) else (
    echo Build failed.
)
pause
