@echo off
cd /d "%~dp0"
REM Prefer the project virtual environment (all dependencies installed); fall back to the system Python.
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" -OO TabEx.py
) else (
    start "" pythonw.exe -OO TabEx.py
)



