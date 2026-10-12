@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

REM Version comes from APP_VERSION in TabEx.py (single source)
set VERSION=
for /f "tokens=3 delims= " %%v in ('findstr /b /c:"APP_VERSION = " TabEx.py') do set VERSION=%%v
set VERSION=%VERSION:"=%
if "%VERSION%"=="" set VERSION=unknown

echo ======================================
echo  TabExplorer v%VERSION% - build EXE
echo ======================================
echo.

echo Existing EXE is kept until all checks pass.
echo Close all TabEx windows and stop VS Code F5 sessions before building.
echo Native validation briefly opens an isolated test window.
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\project_python.ps1" -Action Build
if errorlevel 1 (
    echo.
    echo [ERROR] Build or validation failed. The previous EXE was not replaced.
    echo Check artifacts\release-runtime.json when available.
    pause
    exit /b 1
)

echo.
echo Validated output: TabExplorer.exe
echo Runtime report: artifacts\release-runtime.json
echo Close the old application before starting this release.
pause
exit /b 0
