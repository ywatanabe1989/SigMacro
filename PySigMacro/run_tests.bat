@echo off
setlocal

REM Get the directory of this batch file
set "SCRIPT_DIR=%~dp0"

REM Convert to local drive if running from UNC path
if "%SCRIPT_DIR:~0,2%"=="\\" (
    REM Create a temporary mapped drive
    for /f "tokens=2" %%d in ('net use * /delete ^| findstr "deleted"') do set "DRIVE=%%d"
    if not defined DRIVE (
        echo Unable to find available drive letter
        exit /b 1
    )

    net use %DRIVE% "%SCRIPT_DIR%" >nul
    if errorlevel 1 (
        echo Failed to map network drive
        exit /b 1
    )

    %DRIVE%
    pushd %DRIVE%\
) else (
    pushd "%SCRIPT_DIR%"
)

REM Execute the PowerShell script with execution policy bypass
powershell.exe -ExecutionPolicy Bypass -File "run_tests.ps1" %*

REM Capture exit code
set ERRORLEVEL_BACKUP=%ERRORLEVEL%

REM Clean up mapped drive if created
if "%SCRIPT_DIR:~0,2%"=="\\" (
    popd
    net use %DRIVE% /delete >nul
)

exit /b %ERRORLEVEL_BACKUP%
