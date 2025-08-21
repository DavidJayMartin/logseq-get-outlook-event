@echo off
REM Debug version with console output
setlocal enabledelayedexpansion

REM Set your paths here
set VENV_PATH=C:\Users\martin.1537\OneDrive - The Ohio State University\Digital Mind\logseq-get-outlook-event\.venv
set PYTHON_SCRIPT=C:\Users\martin.1537\OneDrive - The Ohio State University\Digital Mind\logseq-get-outlook-event\outlook_events_api.py
set LOG_FILE=C:\Users\martin.1537\OneDrive - The Ohio State University\Digital Mind\logseq-get-outlook-event\logfile.log

echo Starting script execution...
echo Logging to: %LOG_FILE%

REM Validate paths exist
echo Checking virtual environment path...
if not exist "%VENV_PATH%" (
    echo ERROR: Virtual environment path not found: %VENV_PATH%
    echo %date% %time% - ERROR: Virtual environment path not found: %VENV_PATH% >> "%LOG_FILE%" 2>&1
    pause
    exit /b 2
)
echo Virtual environment path found: %VENV_PATH%

echo Checking Python script path...
if not exist "%PYTHON_SCRIPT%" (
    echo ERROR: Python script not found: %PYTHON_SCRIPT%
    echo %date% %time% - ERROR: Python script not found: %PYTHON_SCRIPT% >> "%LOG_FILE%" 2>&1
    pause
    exit /b 2
)
echo Python script found: %PYTHON_SCRIPT%

REM Clean old log entries (keeping the PowerShell command from original)
echo Cleaning old log entries...
powershell -Command "& { if (Test-Path '%LOG_FILE%') { $cutoffDate = (Get-Date).AddDays(-7); $content = Get-Content '%LOG_FILE%' | Where-Object { if ($_ -match '^(\d{2}\/\d{2}\/\d{4})') { try { $logDate = [DateTime]::ParseExact($matches[1], 'MM/dd/yyyy', $null); $logDate -ge $cutoffDate } catch { $true } } else { $true } }; $content | Out-File '%LOG_FILE%' -Encoding UTF8 } }"

REM Change to the virtual environment directory
echo Changing to virtual environment directory...
cd /d "%VENV_PATH%"

REM Activate the virtual environment
echo Activating virtual environment...
call "%VENV_PATH%\Scripts\activate.bat"

REM Check if activation was successful
if %ERRORLEVEL% neq 0 (
    echo ERROR: Failed to activate virtual environment
    echo %date% %time% - ERROR: Failed to activate virtual environment >> "%LOG_FILE%"
    pause
    exit /b 1
)

echo Virtual environment activated successfully
echo %date% %time% - Virtual environment activated successfully >> "%LOG_FILE%"

REM Show which Python is being used
echo Using Python: 
python --version
python -c "import sys; print('Python path:', sys.executable)"

REM Run your Python script
echo Starting Python script: %PYTHON_SCRIPT%
echo %date% %time% - Starting Python script: %PYTHON_SCRIPT% >> "%LOG_FILE%"

REM Run Python script with output to console (log file output handled separately)
python "%PYTHON_SCRIPT%"

REM Check if Python script ran successfully
if %ERRORLEVEL% neq 0 (
    echo ERROR: Python script failed with error code %ERRORLEVEL%
    echo %date% %time% - ERROR: Python script failed with error code %ERRORLEVEL% >> "%LOG_FILE%"
    pause
    exit /b %ERRORLEVEL%
) else (
    echo Python script completed successfully
    echo %date% %time% - Python script completed successfully >> "%LOG_FILE%"
)

REM Note: If this is a Flask API, it should keep running, not exit
echo Script execution completed
echo %date% %time% - Script execution completed >> "%LOG_FILE%"

echo Press any key to exit...
pause