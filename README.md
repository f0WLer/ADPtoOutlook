## Features

- **Generates calendar files (.ics)** that work with any calendar application
- Compatible with **new Outlook**, old Outlook, Google Calendar, Apple Calendar, and more
- Imports only approved time off requests
- Supports all-day, partial-day, and full-day events
- CLI options for verbose titles, date range filtering, and custom calendar names
- Shows duration, reason, and policy information
- Easy to view all employee time off at a glance
- Optional legacy mode for direct Outlook import (old behavior)

## Requirements

- Windows 10 or later
- Python 3.8+ (for script usage) or use the provided .exe
- Excel file with ADP time off data
- *Optional*: Microsoft Outlook (only needed for `--outlook` legacy mode)

## Installation

Install the required packages:

```powershell
pip install -r requirements.txt
```

The required packages are:
- openpyxl (Excel file reading)
- icalendar (calendar file generation)
- pywin32 (optional, only for `--outlook` mode)

## Usage

### Quick Start

```powershell
# Generate calendar file (default mode)
python src/excel_to_outlook.py example.xlsx
```

This creates a `timeoff_calendar.ics` file that you can import into any calendar application.

### Command-Line Options

```powershell
# Basic usage - creates timeoff_calendar.ics
python src/excel_to_outlook.py example.xlsx

# Specify output filename
python src/excel_to_outlook.py example.xlsx --output vacation_2026.ics

# Verbose event titles (include reason)
python src/excel_to_outlook.py example.xlsx --verbose

# Custom calendar name
python src/excel_to_outlook.py example.xlsx --name "Staff Vacation Calendar"

# Import only a date range
python src/excel_to_outlook.py example.xlsx --range 02-01-2026 02-14-2026

# Combine options
python src/excel_to_outlook.py example.xlsx --verbose --range 01-01-2026 12-31-2026 --name "Q1 Time Off" --output q1.ics

# Legacy mode: Import directly to Outlook (old behavior)
python src/excel_to_outlook.py example.xlsx --outlook
python src/excel_to_outlook.py example.xlsx --outlook --clear
```

**Available Options:**
- `--output FILE`: Specify output .ics filename (default: timeoff_calendar.ics)
- `--outlook`: Import directly to Outlook instead of creating a file (legacy mode)
- `--clear`: Delete all events from calendar before importing (only with `--outlook`)
- `--verbose`: Include reason code in event titles (e.g., "John Doe - PTO")
- `--name "Name"`: Specify a custom base name for the calendar (default: "Employee Time Off")
- `--range START END`: Only import events within date range (format: MM-DD-YYYY MM-DD-YYYY)

### Importing the Calendar File

**Into Outlook (New or Old):**
1. Open Outlook
2. Go to **File > Open & Export > Import/Export**
3. Select **"Import an iCalendar (.ics) or vCalendar file (.vcs)"**
4. Browse to the generated `.ics` file
5. Click **Import** to add to your calendar

**Quick Method:**
- Simply double-click the `.ics` file and it will open in your default calendar app

**Into Google Calendar:**
1. Open Google Calendar
2. Click the **+** next to "Other calendars"
3. Select **Import**
4. Choose the `.ics` file and select which calendar to add it to

### What the Program Does

1. Loads the Excel file and parses all requests
2. Filters for approved, valid, and in-range requests
3. Generates a calendar name based on the date range (or uses custom name if specified)
4. Creates a standard `.ics` calendar file with all events
5. Prints a summary and import instructions
6. You import the file into your calendar application of choice

### Excel File Format

The script expects the following columns:
- NAME: Employee name
- TIME OFF REQUEST DATE: Start date of time off
- DURATION: Number of days
- REQUEST STATUS: Status (only "Approved" requests are imported)
- REASON CODE: Type of time off (vacation, sick, etc.)
- POLICY NAME: Time off policy