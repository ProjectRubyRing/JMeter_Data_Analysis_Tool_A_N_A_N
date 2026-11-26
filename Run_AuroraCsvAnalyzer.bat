@echo off
setlocal EnableExtensions

rem ============================================================================
rem  Aurora Serverless v2 (MySQL 8.4.7) A5:SQL Mk-2 Bulk CSV Export &
rem  Performance Testing Data Analysis Tool Launcher
rem  Windows 11 / PowerShell 5.1 (STA Mode)
rem ============================================================================

set "SCRIPT_DIR=%~dp0"
set "PS1_FILE=%SCRIPT_DIR%AuroraCsvAnalyzer.ps1"

if not exist "%PS1_FILE%" (
    echo [ERROR] PowerShell script not found: "%PS1_FILE%"
    pause
    exit /b 1
)

set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS_EXE%" (
    set "PS_EXE=powershell.exe"
)

if "%~1"=="" (
    "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -STA -File "%PS1_FILE%"
) else (
    "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -STA -File "%PS1_FILE%" %*
)

set "EXIT_CODE=%ERRORLEVEL%"
if "%EXIT_CODE%"=="130" exit /b 130
if not "%EXIT_CODE%"=="0" (
    echo.
    echo [ERROR] Execution failed with exit code %EXIT_CODE%.
    if "%~1"=="" pause
)

exit /b %EXIT_CODE%
