@echo off
chcp 936 >nul
cd /d "%~dp0"

set PYTHON_EXE=python
where python >nul 2>nul
if errorlevel 1 (
    where python3 >nul 2>nul
    if not errorlevel 1 set PYTHON_EXE=python3
)

start "" %PYTHON_EXE% gui_analyzer.py
