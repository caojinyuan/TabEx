@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"
echo ======================================
echo  TabExplorer - install dependencies
echo ======================================
echo.

"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\project_python.ps1" -Action Install
if errorlevel 1 (
    echo.
    echo [ERROR] Project Python or dependency installation failed. See the messages above.
    pause
    exit /b 1
)

echo.
echo Python and dependencies installed into .venv. Double-click 1_run_TabEx.bat to start TabExplorer.
pause
exit /b 0


