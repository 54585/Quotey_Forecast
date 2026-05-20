@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM ==============================================================
REM One-click Quotey Forecast Generator
REM Double-click this file to create a clean forecast workbook.
REM ==============================================================

cd /d "%~dp0"

set "APP_DIR=%~dp0"
set "RAW_DIR=%APP_DIR%Raw_Exports"
set "OUTPUT_DIR=%APP_DIR%Outputs"
set "LOG_FILE=%APP_DIR%run_log.txt"
set "VENV_DIR=%APP_DIR%.venv"
set "VENV_PYTHON=%VENV_DIR%\Scripts\python.exe"
set "CODEX_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "AZURECLI_PYTHON=C:\Program Files\Microsoft SDKs\Azure\CLI2\python.exe"

if not exist "%RAW_DIR%" mkdir "%RAW_DIR%"
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

> "%LOG_FILE%" echo Run started %date% %time%

cls
echo ==============================================================
echo Quotey Forecast Generator
echo ==============================================================
echo.
echo Folder: %APP_DIR%
echo.

REM Locate a real Python interpreter.
set "PYTHON_CMD="
if exist "%VENV_PYTHON%" set "PYTHON_CMD=%VENV_PYTHON%"
if not defined PYTHON_CMD (
    where py >nul 2>nul
    if %ERRORLEVEL%==0 set "PYTHON_CMD=py -3"
)
if not defined PYTHON_CMD (
    where python >nul 2>nul
    if %ERRORLEVEL%==0 for /f "delims=" %%P in ('where python') do (
        echo %%P | find /i "WindowsApps" >nul
        if errorlevel 1 (
            set "PYTHON_CMD=%%P"
            goto :python_found
        )
    )
)
if not defined PYTHON_CMD if exist "%CODEX_PYTHON%" set "PYTHON_CMD=%CODEX_PYTHON%"
if not defined PYTHON_CMD if exist "%AZURECLI_PYTHON%" set "PYTHON_CMD=%AZURECLI_PYTHON%"

:python_found

if not defined PYTHON_CMD (
    echo ERROR: Python was not found.
    echo.
    echo Install Python 3.10 or newer, and select "Add Python to PATH" during setup.
    echo Then double-click this file again.
    echo.
    pause
    exit /b 1
)

if not exist "%VENV_PYTHON%" (
    echo Creating local Python environment...
    "%PYTHON_CMD%" -m venv "%VENV_DIR%" > "%LOG_FILE%" 2>&1
    if errorlevel 1 (
        echo ERROR: Could not create the local Python environment.
        echo See this log file for details:
        echo %LOG_FILE%
        echo.
        pause
        exit /b 1
    )
)

set "PYTHON_CMD=%VENV_PYTHON%"

echo Checking Python packages...
"%PYTHON_CMD%" -m pip install --upgrade pip >> "%LOG_FILE%" 2>&1
"%PYTHON_CMD%" -m pip install -r "%APP_DIR%requirements.txt" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo ERROR: Could not install required Python packages.
    echo See this log file for details:
    echo %LOG_FILE%
    echo.
    pause
    exit /b 1
)

REM Find newest .xlsx in Raw_Exports first, then this folder.
set "INPUT_FILE="
for /f "delims=" %%F in ('dir /b /a-d /o-d "%RAW_DIR%\*.xlsx" 2^>nul') do (
    set "INPUT_FILE=%RAW_DIR%\%%F"
    goto :found_input
)
for /f "delims=" %%F in ('dir /b /a-d /o-d "%APP_DIR%*.xlsx" 2^>nul ^| findstr /v /i "Quotey_Forecast_Clean.xlsx"') do (
    set "INPUT_FILE=%APP_DIR%%%F"
    goto :found_input
)

:found_input
if not defined INPUT_FILE (
    echo ERROR: No raw .xlsx export was found.
    echo.
    echo Put your latest Quotey export into this folder:
    echo %RAW_DIR%
    echo.
    echo The workbook must contain a tab named Export.
    echo.
    pause
    exit /b 1
)

for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd"') do set "TODAY=%%I"
set "OUTPUT_FILE=%OUTPUT_DIR%\Quotey_Forecast_Clean_%TODAY%.xlsx"

echo Input file:
echo %INPUT_FILE%
echo.
echo Output file:
echo %OUTPUT_FILE%
echo.
echo Running forecast processor...

"%PYTHON_CMD%" "%APP_DIR%forecast_processor.py" "%INPUT_FILE%" --output "%OUTPUT_FILE%" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo ERROR: Forecast generation failed.
    echo See this log file for details:
    echo %LOG_FILE%
    echo.
    pause
    exit /b 1
)

echo.
echo SUCCESS: Forecast workbook created.
echo.
echo Opening output workbook...
start "" "%OUTPUT_FILE%"
echo.
pause
exit /b 0
