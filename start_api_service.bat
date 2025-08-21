@echo off
REM Simple background service starter
setlocal enabledelayedexpansion

REM Set relative paths
set SCRIPT_DIR=%~dp0
set VENV_PATH=%SCRIPT_DIR%.venv
set PYTHON_SCRIPT=%SCRIPT_DIR%outlook_events_api.py
set LOG_FILE=%SCRIPT_DIR%logfile.log
set PID_FILE=%SCRIPT_DIR%api_service.pid

echo Starting Outlook Events API in background...

REM Validate paths exist
if not exist "%VENV_PATH%" (
    echo ERROR: Virtual environment not found: %VENV_PATH%
    pause
    exit /b 2
)

if not exist "%PYTHON_SCRIPT%" (
    echo ERROR: Python script not found: %PYTHON_SCRIPT%
    pause
    exit /b 2
)

REM Check if service is already running
if exist "%PID_FILE%" (
    echo Checking if service is already running...
    set /p EXISTING_PID=<"%PID_FILE%"
    tasklist /fi "pid eq !EXISTING_PID!" 2>nul | find "!EXISTING_PID!" >nul
    if !errorlevel! equ 0 (
        echo API service is already running with PID !EXISTING_PID!
        echo Use stop_api_service.bat to stop it first.
        pause
        exit /b 1
    ) else (
        echo Removing stale PID file...
        del "%PID_FILE%" 2>nul
    )
)

REM Clean old log entries (skip if file is locked)
if exist "%LOG_FILE%" (
    echo Cleaning old log entries...
    powershell -Command "try { if (Test-Path '%LOG_FILE%') { $content = Get-Content '%LOG_FILE%' | Select-Object -Last 100; $content | Out-File '%LOG_FILE%' -Encoding UTF8 } } catch { Write-Host 'Log cleanup skipped - file may be open in another program' }"
)

REM Log startup (handle file locking)
powershell -Command "try { '%date% %time% - Starting API service in background' | Add-Content '%LOG_FILE%' } catch { Write-Host 'Could not write to log file - may be open in another program' }"

REM Change to script directory
cd /d "%SCRIPT_DIR%"

REM Start the service using a VBS script to run truly hidden
echo Set WshShell = CreateObject("WScript.Shell") > "%SCRIPT_DIR%\start_hidden.vbs"
echo Set FSO = CreateObject("Scripting.FileSystemObject") >> "%SCRIPT_DIR%\start_hidden.vbs"
echo Set logFile = FSO.OpenTextFile("%LOG_FILE%", 8, True) >> "%SCRIPT_DIR%\start_hidden.vbs"
echo logFile.WriteLine Now ^& " - Launching Python API service" >> "%SCRIPT_DIR%\start_hidden.vbs"
echo WshShell.CurrentDirectory = "%SCRIPT_DIR%" >> "%SCRIPT_DIR%\start_hidden.vbs"
echo cmd = """%VENV_PATH%\Scripts\python.exe"" ""%PYTHON_SCRIPT%""" >> "%SCRIPT_DIR%\start_hidden.vbs"
echo logFile.WriteLine Now ^& " - Command: " ^& cmd >> "%SCRIPT_DIR%\start_hidden.vbs"
echo WshShell.Run cmd, 0, False >> "%SCRIPT_DIR%\start_hidden.vbs"
echo logFile.WriteLine Now ^& " - Command sent to shell" >> "%SCRIPT_DIR%\start_hidden.vbs"
echo logFile.Close >> "%SCRIPT_DIR%\start_hidden.vbs"

cscript //nologo "%SCRIPT_DIR%start_hidden.vbs"

REM Wait a moment for the process to start
timeout /t 3 /nobreak >nul

REM Find the Python process (more reliable method)
set SERVICE_PID=
echo Looking for the Python service process...

REM Wait a bit longer for process to fully start
timeout /t 5 /nobreak >nul

REM Method 1: Look for python processes and check their command line
for /f "skip=1 tokens=2" %%i in ('wmic process where "name='python.exe'" get ProcessId /format:table 2^>nul') do (
    if "%%i" neq "" if "%%i" neq "ProcessId" (
        for /f "tokens=*" %%j in ('wmic process where "ProcessId=%%i" get CommandLine /format:value 2^>nul ^| find "outlook_events_api.py"') do (
            set SERVICE_PID=%%i
            echo Found service process with PID: %%i
            goto :found_pid
        )
    )
)

REM Method 2: If Method 1 fails, try simpler approach
if not defined SERVICE_PID (
    echo Method 1 failed, trying alternative detection...
    for /f "skip=3 tokens=2" %%i in ('tasklist /fi "imagename eq python.exe" 2^>nul') do (
        if not defined SERVICE_PID (
            REM Check if this python process is in our directory
            for /f %%j in ('powershell -Command "try { $p = Get-Process -Id %%i -ErrorAction Stop; $p.MainModule.FileName } catch { 'NONE' }"') do (
                echo Checking process %%i: %%j
                echo %%j | find /i ".venv" >nul && (
                    set SERVICE_PID=%%i
                    echo Found service process with PID: %%i
                    goto :found_pid
                )
            )
        )
    )
)

REM Method 3: Last resort - just grab the newest python process
if not defined SERVICE_PID (
    echo Method 2 failed, using newest python process...
    for /f "skip=3 tokens=2" %%i in ('tasklist /fi "imagename eq python.exe" 2^>nul') do (
        set SERVICE_PID=%%i
        echo Assuming PID: %%i (newest python process)
        goto :found_pid
    )
)

:found_pid

REM Clean up the VBS file
del "%SCRIPT_DIR%start_hidden.vbs" 2>nul

if defined SERVICE_PID (
    echo !SERVICE_PID! > "%PID_FILE%"
    echo API service started successfully in background
    echo PID: !SERVICE_PID!
    echo Log file: %LOG_FILE%
    powershell -Command "try { '%date% %time% - API service started with PID !SERVICE_PID!' | Add-Content '%LOG_FILE%' } catch { }"
    
    REM Test common Flask ports
    echo.
    echo Testing service availability on common ports...
    timeout /t 2 /nobreak >nul
    
    set SERVICE_FOUND=0
    for %%p in (5000 8000 8080 3000) do (
        if !SERVICE_FOUND! equ 0 (
            powershell -Command "try { $response = Invoke-WebRequest -Uri 'http://localhost:%%p' -TimeoutSec 2 -ErrorAction Stop; Write-Host 'Service is responding on http://localhost:%%p'; exit 1 } catch { }" && set SERVICE_FOUND=1
        )
    )
    
    if !SERVICE_FOUND! equ 0 (
        echo Service started but not responding on common ports (5000, 8000, 8080, 3000)
        echo Check your Flask app configuration or the log file for the actual port
    )
    
) else (
    echo WARNING: Could not determine service PID
    echo The service may have started but PID detection failed.
    echo Check the log file: %LOG_FILE%
    echo Or use Task Manager to verify if python.exe is running.
)

echo.
echo Management commands:
echo - To stop the service: stop_api_service.bat
echo - To check status: check_api_service.bat
echo - To view recent logs: type "%LOG_FILE%"
echo.
pause