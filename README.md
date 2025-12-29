<!-- MOVED TO docs/README.md -->

## Features

- Creates a dedicated "Employee Time Off" calendar in Outlook (with date range in name)
- Imports only approved time off requests
- Supports all-day, partial-day, and full-day events
- CLI options for clearing calendar, verbose titles, and date range filtering
- Shows duration, reason, and policy information
- Easy to view all employee time off at a glance

## Requirements

- Windows 10 or later
- Microsoft Outlook installed and configured
- Python 3.8+ (for script usage) or use the provided .exe
- Excel file with ADP time off data

## Installation

The required packages are listed in `requirements.txt`:
- pandas
- openpyxl
- pywin32

## Usage

### Command-Line Options

```powershell
# Basic usage
python excel_to_outlook.py timeoff.xlsx

# Clear calendar before import
python excel_to_outlook.py timeoff.xlsx --clear

# Verbose event titles (include reason)
python excel_to_outlook.py timeoff.xlsx --verbose

# Import only a date range
python excel_to_outlook.py timeoff.xlsx --range 02-01-2026 02-14-2026

# Combine options
python excel_to_outlook.py timeoff.xlsx --clear --verbose --range 01-01-2026 12-31-2026
```

### What the Program Does

1. Loads the Excel file and parses all requests
2. Filters for approved, valid, and in-range requests
3. Generates a calendar name based on the date range
4. (Optional) Clears the calendar if `--clear` is used
5. Connects to Outlook and creates/fetches the calendar
6. Creates events for each approved request
7. Prints a summary (events created/skipped)
8. User opens Outlook to view the new calendar

### Excel File Format

The script expects the following columns:
- NAME: Employee name
- TIME OFF REQUEST DATE: Start date of time off
- DURATION: Number of days
- REQUEST STATUS: Status (only "Approved" requests are imported)
- REASON CODE: Type of time off (vacation, sick, etc.)
- POLICY NAME: Time off policy

## Customization

- Change the calendar name (default: "Employee Time Off")
- Filter by different status values
- Change how events are displayed
- Add custom categories or colors

## Support

Contact your IT administrator for assistance or see DEVNOTES.md for developer documentation.
