@echo off
cd /d "%~dp0"
title Markdown Studio

:: Check Python installation
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python 3.10+ from https://python.org
    pause
    exit /b 1
)

:: Try launch with pythonw (no console window)
start pythonw run_gui.py
if %ERRORLEVEL% neq 0 (
    python run_gui.py
)
