@echo off
cd /d "%~dp0"
echo Installing Python packages from requirements.txt...
echo.

where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python not found. Please install Python 3.6+ and add it to PATH.
    pause
    exit /b 1
)

if not exist requirements.txt (
    echo ERROR: requirements.txt not found in current directory.
    pause
    exit /b 1
)

python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Installation failed. See messages above.
    pause
    exit /b 1
)

echo.
echo Installation completed successfully.
pause
