@echo off
setlocal enabledelayedexpansion
title TimeOff Calendar - Quick Import

REM ================================================================
REM   Drag & Drop Quick Import
REM   Drop an Excel file onto this batch file to create a calendar
REM ================================================================

REM Check if a file was dropped
if "%~1"=="" (
    echo ================================================================
    echo            Employee Time-Off Calendar Importer
    echo                    DRAG ^& DROP MODE
    echo ================================================================
    echo.
    echo No file was dropped!
    echo.
    echo To use this tool:
    echo   1. Find your Excel file with time-off data
    echo   2. Drag and drop it onto this batch file
    echo   3. A calendar file will be created automatically
    echo.
    echo For advanced options, use TimeOffCalendar.bat instead.
    echo.
    pause
    exit /b 1
)

REM Check if exe exists
if not exist "%~dp0TimeOffCalendar.exe" (
    echo ERROR: TimeOffCalendar.exe not found in this folder!
    echo.
    pause
    exit /b 1
)

REM Get the dropped file path
set EXCEL_FILE=%~1

REM Check if file exists
if not exist "!EXCEL_FILE!" (
    echo ERROR: File not found: !EXCEL_FILE!
    echo.
    pause
    exit /b 1
)

REM Get the base name of the Excel file (without extension)
set EXCEL_NAME=%~n1

REM Create output filename based on input file
set OUTPUT_FILE=!EXCEL_NAME!_calendar.ics

echo ================================================================
echo            Employee Time-Off Calendar Importer
echo                    DRAG ^& DROP MODE
echo ================================================================
echo.
echo Input File:  !EXCEL_FILE!
echo Output File: !OUTPUT_FILE!
echo Calendar:    Employee Time Off
echo.
echo Creating calendar file...
echo.

REM Run the converter with default settings
"%~dp0TimeOffCalendar.exe" "!EXCEL_FILE!" --output "!OUTPUT_FILE!" --name "Employee Time Off"

if !ERRORLEVEL! EQU 0 (
    echo.
    echo ================================================================
    echo                   SUCCESS!
    echo ================================================================
    echo.
    echo Calendar file created: !OUTPUT_FILE!
    echo.
    echo You can now:
    echo   - Double-click the .ics file to open it
    echo   - Import it into Outlook, Google Calendar, etc.
    echo.
) else (
    echo.
    echo ================================================================
    echo                   ERROR OCCURRED
    echo ================================================================
    echo.
    echo The import failed. Check the error messages above.
    echo.
)

pause
