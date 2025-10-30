@echo off
cd /d %~dp0

echo [1] Register context menu
echo [2] Unregister context menu
echo [3] Exit
set /p action="Enter your choice: "

if "%action%"=="1" (
    python register-context.py
    goto:eof
) else if "%action%"=="2" (
    python register-context.py --uninstall
    goto:eof
) else (
    echo Exiting.
    goto:eof
)
