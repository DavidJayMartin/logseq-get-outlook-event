@echo off
REM Stop the Flask API background service
setlocal enabledelayedexpansion

set SCRIPT_DIR=%~dp0
set PID_FILE=%SCRIPT_DIR%api_service.pid
set LOG_FILE=%SCRIPT_DIR%logfile.log

echo Stopping Outlook Events API service...

if not exist "%PID_FILE%" (
    echo No PID file found. Service may not be running.
    echo Attempting to stop any Python processes in the virtual environment...
    
    REM Try to find and stop Python processes from our venv
    echo Attempting to stop any Python processes in the virtual environment...
    
    set STOPPED_ANY=0
    for /f "skip=3 tokens=2" %%i in ('tasklist /fi "imagename eq python.exe" 2^>nul') do (
        powershell -Command "try { $proc = Get-Process -Id %%i -ErrorAction Stop; if ($proc.Path -like '*\.venv\*') { Write-Host 'Found and stopping process %%i'; Stop-Process -Id %%i -Force; exit 1 } } catch { exit 0 }" && set STOPPED_ANY=1
    )
    
    if !STOPPED_ANY! equ 1 (
        echo Stopped orphaned virtual environment processes.
    ) else (
        echo No related processes found.
    )
    pause
    exit /b 1
)

REM Read PID from file
for /f %%i in ('type "%PID_FILE%"') do set SERVICE_PID=%%i

echo Found service PID: !SERVICE_PID!

REM Check if process is still running
tasklist /fi "pid eq !SERVICE_PID!" 2>nul | find "!SERVICE_PID!" >nul
if !errorlevel! neq 0 (
    echo Process !SERVICE_PID! is not running.
    del "%PID_FILE%" 2>nul
    echo Cleaned up stale PID file.
    pause
    exit /b 0
)

REM Stop the process (try multiple methods)
echo Stopping process !SERVICE_PID!...

REM First try graceful termination
taskkill /pid !SERVICE_PID! >nul 2>&1
set KILL_RESULT=!errorlevel!

REM If that fails, try forceful termination
if !KILL_RESULT! neq 0 (
    echo Graceful termination failed, trying forceful stop...
    taskkill /pid !SERVICE_PID! /f >nul 2>&1
    set KILL_RESULT=!errorlevel!
)

REM If still failing, try PowerShell method
if !KILL_RESULT! neq 0 (
    echo Trying PowerShell method...
    powershell -Command "try { Stop-Process -Id !SERVICE_PID! -Force -ErrorAction Stop; exit 0 } catch { exit 1 }"
    set KILL_RESULT=!errorlevel!
)

if !KILL_RESULT! equ 0 (
    echo API service stopped successfully.
    
    REM Try to log, but don't fail if we can't
    powershell -Command "try { '%date% %time% - API service stopped (PID !SERVICE_PID!)' | Add-Content '%LOG_FILE%' } catch { }"
    
    del "%PID_FILE%" 2>nul
    
    REM Wait a moment and verify it's really stopped
    timeout /t 2 /nobreak >nul
    tasklist /fi "pid eq !SERVICE_PID!" 2>nul | find "!SERVICE_PID!" >nul
    if !errorlevel! neq 0 (
        echo Process !SERVICE_PID! has been successfully terminated.
    ) else (
        echo Warning: Process !SERVICE_PID! may still be running.
    )
) else (
    echo Failed to stop process !SERVICE_PID!
    echo.
    echo This could happen if:
    echo 1. The process is already stopped
    echo 2. The process has elevated privileges
    echo 3. The process is hung or unresponsive
    echo.
    echo You can try:
    echo 1. Run this script as Administrator
    echo 2. Use Task Manager to end the process manually
    echo 3. Restart your computer if the process is completely stuck
)

pause