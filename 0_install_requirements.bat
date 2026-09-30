@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"
echo ======================================
echo  TabExplorer - install dependencies
echo ======================================
echo.

set "VENV_PY=.venv\Scripts\python.exe"
if exist "%VENV_PY%" goto install

set "BASE_PY="
py -3 --version >nul 2>&1 && set "BASE_PY=py -3"
if not defined BASE_PY (
    python --version >nul 2>&1 && set "BASE_PY=python"
)
if not defined BASE_PY (
    echo [ERROR] Python not found. Please install Python 3.9 or later.
    pause
    exit /b 1
)

echo Creating the project virtual environment .venv ...
%BASE_PY% -m venv .venv
if errorlevel 1 (
    echo [ERROR] Failed to create the virtual environment, see the messages above.
    pause
    exit /b 1
)

:install
"%VENV_PY%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install dependencies, see the messages above.
    pause
    exit /b 1
)

echo.
echo Dependencies installed into .venv. Double-click 1_run_TabEx.bat to start TabExplorer.
pause


