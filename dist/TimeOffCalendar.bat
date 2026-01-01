@echo off
setlocal enabledelayedexpansion
title TimeOff Calendar Importer

echo ================================================================
echo            Employee Time-Off Calendar Importer
echo ================================================================
echo.

REM Check if exe exists
if not exist "%~dp0TimeOffCalendar.exe" (
    echo ERROR: TimeOffCalendar.exe not found in this folder!
    echo.
    pause
    exit /b 1
)

echo Step 1: Excel File
echo ------------------
echo Drag and drop your Excel file here, or type the path:
echo (Press Enter to use 'timeoff.xlsx' in the current folder)
echo.
set /p EXCEL_FILE="Excel file: "

REM Default to timeoff.xlsx if empty
if "!EXCEL_FILE!"=="" set EXCEL_FILE=timeoff.xlsx

REM Remove quotes if user added them
set EXCEL_FILE=!EXCEL_FILE:"=!

REM Check if file exists
if not exist "!EXCEL_FILE!" (
    echo.
    echo ERROR: File '!EXCEL_FILE!' not found!
    echo.
    pause
    exit /b 1
)

echo.
echo Step 2: Output Mode
echo ---------------------
echo How do you want to save the calendar?
echo   1. Save as file (.ics) - Works with any calendar app (RECOMMENDED)
echo   2. Import directly to Outlook (legacy mode)
echo.
set /p OUTPUT_MODE="Choose mode (1 or 2, default=1): "

REM Default to file mode if empty or invalid
if "!OUTPUT_MODE!"=="" set OUTPUT_MODE=1
if not "!OUTPUT_MODE!"=="1" if not "!OUTPUT_MODE!"=="2" set OUTPUT_MODE=1

set OUTLOOK_FLAG=
set OUTPUT_FILE=timeoff_calendar.ics

if "!OUTPUT_MODE!"=="2" goto outlook_mode

REM File mode
echo.
echo Step 3: Output Filename
echo -----------------------
echo What should the calendar file be called?
echo (Press Enter for default: "timeoff_calendar.ics")
echo.
set /p OUTPUT_FILE="Filename: "
if "!OUTPUT_FILE!"=="" set OUTPUT_FILE=timeoff_calendar.ics
REM Ensure .ics extension
if not "!OUTPUT_FILE:~-4!"==".ics" set OUTPUT_FILE=!OUTPUT_FILE!.ics
goto after_mode_selection

:outlook_mode
set OUTLOOK_FLAG=--outlook
echo.
echo Using legacy Outlook import mode...

:after_mode_selection

echo.
echo Step 4: Calendar Name
echo ---------------------
echo What should the calendar be called?
echo (Press Enter for default: "Employee Time Off")
echo.
set /p CALENDAR_NAME="Calendar name: "

REM Default if empty
if "!CALENDAR_NAME!"=="" set CALENDAR_NAME=Employee Time Off

set CLEAR_FLAG=

REM Only ask about clearing in Outlook mode
if "!OUTPUT_MODE!"=="2" goto ask_clear
set STEP_NUM=5
goto after_clear

:ask_clear
echo.
echo Step 5: Clear Existing Calendar
echo --------------------------------
echo Do you want to clear the existing calendar before importing?
echo (Type 'y' for yes, or press Enter to skip)
echo.
set /p CLEAR_CHOICE="Clear calendar? (y/N): "

if /i "!CLEAR_CHOICE!"=="y" set CLEAR_FLAG=--clear
if /i "!CLEAR_CHOICE!"=="yes" set CLEAR_FLAG=--clear
set STEP_NUM=6

:after_clear

echo.
echo Step !STEP_NUM!: Date Range Filter (Optional)
echo -------------------------------------
echo Do you want to import only specific dates?
echo (Type 'y' for yes, or press Enter to import all dates)
echo.
set /p DATE_FILTER="Filter by date range? (y/N): "

set DATE_RANGE=
if /i "!DATE_FILTER!"=="y" (
    echo.
    echo Enter date range in format: MM-DD-YYYY
    echo.
    set /p START_DATE="Start date (e.g., 01-01-2025): "
    set /p END_DATE="End date   (e.g., 12-31-2025): "
    
    if not "!START_DATE!"=="" if not "!END_DATE!"=="" (
        set DATE_RANGE=--range "!START_DATE!" "!END_DATE!"
    )
)

echo.
echo ================================================================
echo                    Starting Import...
echo ================================================================
echo.

REM Build and run the command
if "!OUTPUT_MODE!"=="1" (
    "%~dp0TimeOffCalendar.exe" "!EXCEL_FILE!" --output "!OUTPUT_FILE!" --name "!CALENDAR_NAME!" !DATE_RANGE!
) else (
    "%~dp0TimeOffCalendar.exe" "!EXCEL_FILE!" --outlook --name "!CALENDAR_NAME!" !CLEAR_FLAG! !DATE_RANGE!
)

echo.
echo ================================================================
echo                        Complete!
echo ================================================================
echo.
echo Press any key to close this window...
pause >nul
