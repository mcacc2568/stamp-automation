@echo off
cd /d "%~dp0"

echo [1/2] Installing dependencies...
python -m pip install pyinstaller pymupdf pillow -q --disable-pip-version-check
if %errorlevel% neq 0 (
    echo ERROR: pip install failed
    pause
    exit /b 1
)

echo [2/2] Building .exe...
python -m PyInstaller --onefile --windowed --clean --noconfirm --name "StampPDF" --collect-all tkinter main.py
if %errorlevel% neq 0 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

echo Done! File is at: dist\StampPDF.exe
pause
