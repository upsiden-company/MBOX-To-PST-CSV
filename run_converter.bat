@echo off
REM MBOX Converter - Windows Batch Runner
REM This script ensures Python is available and runs the converter

setlocal enabledelayedexpansion

REM Check if Python is available
where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Python not found. Downloading installer...
    set INST=%TEMP%\python-installer.exe
    curl -L -o %INST% https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe
    echo Installing Python...
    %INST% /quiet InstallAllUsers=1 PrependPath=1
    del %INST%
    echo Python installed. Please restart the command prompt and run again.
    pause
    exit /b 1
)

REM Check if dependencies are installed
python -c "import tqdm" >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Installing dependencies...
    pip install -r requirements.txt
)

REM Run the converter with all passed arguments
python mbox_converter.py %*

endlocal
