@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

REM Prefer the project virtual environment so the build uses the same packages as running from source
set "PY=python"
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"

REM Version comes from APP_VERSION in TabEx.py (single source)
set VERSION=
for /f "tokens=3 delims= " %%v in ('findstr /b /c:"APP_VERSION = " TabEx.py') do set VERSION=%%v
set VERSION=%VERSION:"=%
if "%VERSION%"=="" set VERSION=unknown

echo ======================================
echo  TabExplorer v%VERSION% - build EXE
echo ======================================
echo.

"%PY%" --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found.
    echo Install Python 3.9 or later, or run 0_install_requirements.bat first.
    pause
    exit /b 1
)

REM Install runtime dependencies so PyInstaller can collect PyQt5 and the other packages
echo [Step 1/5] Installing requirements.txt with %PY% ...
"%PY%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install requirements.txt.
    pause
    exit /b 1
)

"%PY%" -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [Step 2/5] Installing PyInstaller...
    "%PY%" -m pip install pyinstaller
    if errorlevel 1 (
        echo [ERROR] Failed to install PyInstaller.
        pause
        exit /b 1
    )
) else (
    echo [Step 2/5] PyInstaller is already installed
)

echo.
echo [Step 3/5] Removing old build output...
if exist TabExplorer.exe (
    echo Deleting old TabExplorer.exe
    del /q TabExplorer.exe
)
if exist build (
    echo Deleting old build folder
    rmdir /s /q build
)
if exist dist (
    echo Deleting old dist folder
    rmdir /s /q dist
)
if exist *.spec (
    echo Deleting old spec files
    del /q *.spec
)

echo.
echo [Step 4/5] Building, this can take a few minutes...
echo.

if exist "icons\TabExplorer.ico" (
    echo Using icons\TabExplorer.ico
    set ICON_PARAM=--icon="icons\TabExplorer.ico"
) else (
    echo icons\TabExplorer.ico not found, using the default icon
    set ICON_PARAM=
)

REM Single-file EXE written to this folder
"%PY%" -m PyInstaller --onefile --windowed --name TabExplorer %ICON_PARAM% --add-data "icons;icons" --distpath . TabEx.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed, see the messages above.
    pause
    exit /b 1
)

echo.
echo [Step 5/5] Cleaning up temporary files...
if exist build rmdir /s /q build
if exist TabExplorer.spec del /q TabExplorer.spec

echo.
echo ======================================
echo  Build finished: v%VERSION%
echo ======================================
echo.
echo Output: TabExplorer.exe
if exist TabExplorer.exe (
    for %%A in (TabExplorer.exe) do echo Size: %%~zA bytes
)
echo.
echo Notes:
echo 1. config.json and bookmarks.json are created next to the EXE on first run
echo 2. Ship README.md together with the EXE when publishing
echo.
echo Next steps:
echo - Test: double-click TabExplorer.exe
echo - Publish: upload to GitHub Releases
echo.
pause
