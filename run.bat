@echo off
setlocal

set /p PHASE=Phase number to process the whole guild (leave empty to paste a single report URL instead):

if "%PHASE%"=="" goto :single_report

echo "%PHASE%" | findstr /C:"://" >nul
if not errorlevel 1 (
    set "WCL_URL=%PHASE%"
    goto :run_single
)

"%~dp0.venv\Scripts\python.exe" -m log_analyzer.cli --phase %PHASE%
goto :end

:single_report
set /p WCL_URL=WCL Report URL:

:run_single
"%~dp0.venv\Scripts\python.exe" -m log_analyzer.cli "%WCL_URL%"

:end
pause
