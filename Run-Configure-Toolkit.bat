@echo off
setlocal
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Configure-Toolkit.ps1"
if errorlevel 1 (
    echo.
    echo EndpointGuy Toolkit Configuration exited with an error.
    pause
)
endlocal
