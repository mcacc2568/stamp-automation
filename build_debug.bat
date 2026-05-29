@echo off
cd /d "%~dp0"

echo Building debug version...
python -m PyInstaller --onefile --console --name "StampPDF_debug" main.py
if %errorlevel% neq 0 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

echo Done! Run dist\StampPDF_debug.exe to see errors
pause
