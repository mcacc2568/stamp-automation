@echo off
cd /d "%~dp0"

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found. Please install from python.org
    pause
    exit /b 1
)

python -m pip install pymupdf pillow -q --disable-pip-version-check
start "" pythonw main.py
