@echo off
setlocal enabledelayedexpansion
title Gold Procurement Report Converter

echo =====================================================
echo   Gold Procurement TXT to Excel Converter
echo =====================================================
echo.

REM Script lives in gold-procurement\ but .txt files are in the parent folder
set "SCRIPT_DIR=%~dp0"
set "DATA_DIR=%~dp0..\"

cd /d "%SCRIPT_DIR%"

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.x and try again.
    pause
    exit /b 1
)

REM If a file is dragged onto the batch, process just that file
if not "%~1"=="" (
    echo Processing: %~1
    python parse_gold_report.py "%~1"
    goto done
)

REM Otherwise find all .txt files in the PARENT folder (where reports live)
set "FOUND_ANY="

for %%F in ("%DATA_DIR%*.txt") do (
    echo Found: %%F
    set "FOUND_ANY=1"
)

if not defined FOUND_ANY (
    echo.
    echo No .txt files found in:
    echo %DATA_DIR%
    echo.
    echo Please place the monthly report .txt file there and run again.
    pause
    exit /b 1
)

echo.
echo ─────────────────────────────────────────────────────
echo Processing all .txt files in parent folder
echo ─────────────────────────────────────────────────────
echo.

REM Pass each .txt file explicitly so the script resolves paths correctly
for %%F in ("%DATA_DIR%*.txt") do (
    python parse_gold_report.py "%%F"
)

:done
echo.
if errorlevel 1 (
    echo =====================================================
    echo   COMPLETED WITH ERRORS  -  check output above
    echo =====================================================
) else (
    echo =====================================================
    echo   SUCCESS  -  Excel file(s) saved in same folder
    echo =====================================================
)
echo.
pause
