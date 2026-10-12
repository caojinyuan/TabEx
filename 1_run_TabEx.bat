@echo off
setlocal
cd /d "%~dp0"
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\project_python.ps1" -Action Run
if errorlevel 1 (
    pause
    exit /b 1
)
exit /b 0



