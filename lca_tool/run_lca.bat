@echo off
REM Double-click: reads elements.txt and writes TSCT-filled.xlsm
cd /d "%~dp0"

set WB=TSCT-v3-01.xlsm
if not exist "%WB%" (
  echo Could not find %WB% in this folder:
  echo %~dp0
  echo Copy your TSCT workbook here, named %WB%.
  pause
  exit /b 1
)

set PY=
where py >nul 2>nul && set PY=py
if "%PY%"=="" where python >nul 2>nul && set PY=python
if "%PY%"=="" (
  echo Python was not found. Install it from https://www.python.org/downloads/
  echo and tick "Add python.exe to PATH" during installation.
  pause
  exit /b 1
)

%PY% tsct_insert.py elements.txt "%WB%" TSCT-filled.xlsm
if errorlevel 1 (
  echo.
  echo Something went wrong - see the message above.
) else (
  start "" "TSCT-filled.xlsm"
)
pause
