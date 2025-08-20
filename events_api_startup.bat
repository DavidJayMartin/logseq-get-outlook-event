@echo off
REM Batch script to activate Python virtual environment and run Python program
REM Configure the paths below according to your setup

REM Enable delayed variable expansion for better error handling
setlocal enabledelayedexpansion

REM Set your paths here
set VENV_PATH=C:\Users\martin.1537\OneDrive - The Ohio State University\Digital Mind\logseq-get-outlook-event\.venv
set PYTHON_SCRIPT=C:\Users\martin.1537\OneDrive - The Ohio State University\Digital Mind\logseq-get-outlook-event\outlook_events_api.py
set LOG_FILE=C:\Users\martin.1537\OneDrive - The Ohio State University\Digital Mind\logseq-get-outlook-event\logfile.log

REM Validate paths exist
echo %date% %time% - Starting script execution >> "%LOG_FILE%" 2>&1
if not exist "%VENV_PATH%" (
    echo %date% %time% - ERROR: Virtual environment path not found: %VENV_PATH% >> "%LOG_FILE%" 2>&1
    exit /b 2
)
if not exist "%PYTHON_SCRIPT%" (
    echo %date% %time% - ERROR: Python script not found: %PYTHON_SCRIPT% >> "%LOG_FILE%" 2>&1
    exit /b 2
)

REM Calculate date 7 days ago
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YYYY=%dt:~0,4%"
set "MM=%dt:~4,2%"
set "DD=%dt:~6,2%"

REM Create PowerShell command to clean old log entries
powershell -Command "& { if (Test-Path '%LOG_FILE%') { $cutoffDate = (Get-Date).AddDays(-7); $content = Get-Content '%LOG_FILE%' | Where-Object { if ($_ -match '^(\d{2}\/\d{2}\/\d{4})') { try { $logDate = [DateTime]::ParseExact($matches[1], 'MM/dd/yyyy', $null); $logDate -ge $cutoffDate } catch { $true } } else { $true } }; $content | Out-File '%LOG_FILE%' -Encoding UTF8 } }"

REM Change to the virtual environment directory
cd /d "%VENV_PATH%"

REM Activate the virtual environment
call "%VENV_PATH%\Scripts\activate.bat"

REM Check if activation was successful
if %ERRORLEVEL% neq 0 (
    echo %date% %time% - ERROR: Failed to activate virtual environment >> "%LOG_FILE%"
    exit /b 1
)

echo %date% %time% - Virtual environment activated successfully >> "%LOG_FILE%"

REM Run your Python script
echo %date% %time% - Starting Python script: %PYTHON_SCRIPT% >> "%LOG_FILE%"
python "%PYTHON_SCRIPT%" >> "%LOG_FILE%" 2>&1

REM Check if Python script ran successfully
if %ERRORLEVEL% neq 0 (
    echo %date% %time% - ERROR: Python script failed with error code %ERRORLEVEL% >> "%LOG_FILE%"
    exit /b %ERRORLEVEL%
) else (
    echo %date% %time% - Python script completed successfully >> "%LOG_FILE%"
)

REM Deactivate virtual environment (optional, as the script will end anyway)
deactivate

echo %date% %time% - Script execution completed >> "%LOG_FILE%"