@echo off
echo Installing Python packages from requirements.txt...

rem Check if requirements.txt exists
if not exist requirements.txt (
    echo ERROR: requirements.txt not found in current directory.
    pause
    exit /b 1
)

rem Run pip install
python -m pip install -r requirements.txt

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Installation failed. See messages above.
    pause
    exit /b 1
)

echo Installation completed successfully.
pause
