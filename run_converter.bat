@echo off
setlocal
where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Python not found. Downloading installer...
    set INST=%TEMP%\python-installer.exe
    curl -L -o %INST% https://www.python.org/ftp/python/3.10.12/python-3.10.12-amd64.exe
    %INST% /quiet InstallAllUsers=1 PrependPath=1
    del %INST%
)
python mbox_converter.py %*
