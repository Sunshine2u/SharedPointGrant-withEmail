@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo Starting SharePoint Permission Manager...
set "PYTHON_CMD="
call :find_python

if not defined PYTHON_CMD (
    echo Python was not found. Installing Python 3.12...
    where winget >nul 2>&1
    if errorlevel 1 (
        echo ERROR: winget is required to install Python automatically.
        echo Install Python from https://www.python.org/downloads/ and run this file again.
        pause
        exit /b 1
    )

    winget install --id Python.Python.3.12 --exact --source winget --accept-source-agreements --accept-package-agreements
    if errorlevel 1 (
        echo ERROR: Python installation failed.
        pause
        exit /b 1
    )

    call :find_python
)

if not defined PYTHON_CMD (
    echo ERROR: Python was installed but could not be located.
    echo Close this window, open a new Command Prompt, and run this file again.
    pause
    exit /b 1
)

echo Using %PYTHON_CMD%
"%PYTHON_CMD%" %PYTHON_ARGS% -m pip --version >nul 2>&1
if errorlevel 1 (
    echo pip was not found. Enabling pip...
    "%PYTHON_CMD%" %PYTHON_ARGS% -m ensurepip --upgrade
    if errorlevel 1 (
        echo ERROR: Could not enable pip.
        pause
        exit /b 1
    )
)

"%PYTHON_CMD%" %PYTHON_ARGS% -m pip install -r config\requirements.txt
if errorlevel 1 (
    echo ERROR: Python dependencies could not be installed.
    pause
    exit /b 1
)

echo Starting Python Backend Server...
start "SharePoint Permission Manager" "%PYTHON_CMD%" %PYTHON_ARGS% app\main.py
echo Waiting for server to start...
timeout /t 3 /nobreak >nul
start http://127.0.0.1:5000
exit /b 0

:find_python
python --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    set "PYTHON_ARGS="
    exit /b 0
)

py -3 --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py"
    set "PYTHON_ARGS=-3"
    exit /b 0
)

for /f "delims=" %%P in ('where /r "%LocalAppData%\Programs\Python" python.exe 2^>nul') do if not defined PYTHON_CMD set "PYTHON_CMD=%%P"
if defined PYTHON_CMD exit /b 0

for /d %%D in ("%ProgramFiles%\Python*") do for /f "delims=" %%P in ('where /r "%%~D" python.exe 2^>nul') do if not defined PYTHON_CMD set "PYTHON_CMD=%%P"
exit /b 0