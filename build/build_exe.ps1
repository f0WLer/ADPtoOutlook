# Build script for creating standalone .exe with PyInstaller

Write-Host "TimeOffCalendar - Build Script" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green
Write-Host ""

# Check if PyInstaller is installed
Write-Host "Checking for PyInstaller..." -ForegroundColor Cyan
$pyInstallerCheck = python -m pip show pyinstaller 2>$null
if (-not $pyInstallerCheck) {
    Write-Host "PyInstaller not found. Installing..." -ForegroundColor Yellow
    python -m pip install pyinstaller
} else {
    Write-Host "PyInstaller found!" -ForegroundColor Green
}

# Clean previous builds
Write-Host "`nCleaning previous builds..." -ForegroundColor Cyan
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
if (Test-Path "*.spec") { Remove-Item -Force "*.spec" }

# Build the executable
Write-Host "`nBuilding executable..." -ForegroundColor Cyan
Write-Host "This may take a few minutes..." -ForegroundColor Yellow
pyinstaller --onefile `
    --console `
    --name "TimeOffCalendar" `
    --icon NONE `
    --add-data "..\\README.md;." `
    ..\src\excel_to_outlook.py

# Check if build was successful
if (Test-Path "dist\TimeOffCalendar.exe") {
    Write-Host "`n✓ Build successful!" -ForegroundColor Green
    
    # Create distribution folder
    $distFolder = "TimeOffCalendar_Distribution"
    Write-Host "`nCreating distribution package..." -ForegroundColor Cyan
    New-Item -ItemType Directory -Force -Path $distFolder | Out-Null
    
    # Copy files
    Copy-Item "dist\TimeOffCalendar.exe" "$distFolder\"
    Copy-Item "..\README.md" "$distFolder\" -ErrorAction SilentlyContinue
    
    # Create user instructions
    @"
# TimeOffCalendar - User Guide

## Quick Start

1. Place your Excel file (e.g., timeoff.xlsx) in this folder
2. Double-click TimeOffCalendar.exe
3. The program will process your file and create Outlook calendar events

## Requirements

- Windows 10 or later
- Microsoft Outlook installed and configured
- Excel file with employee time-off data

## What This Does

- Reads approved time-off requests from Excel
- Creates calendar events in Outlook
- Organizes events in a dedicated calendar folder
- Shows employee names and time blocks

## Troubleshooting

**Executable won't run:**
- Windows may show a SmartScreen warning (click "More info" → "Run anyway")
- Some antivirus software may flag it - add an exception if needed

**"Could not connect to Outlook":**
- Ensure Outlook is installed
- Open Outlook manually first to verify it works

**Excel file not found:**
- Make sure your Excel file is in the same folder as the .exe
- Edit the script or use command line: TimeOffCalendar.exe path\to\file.xlsx

## File Size

The executable is ~30-50MB because it includes Python and all necessary libraries.

## Support

Contact your IT administrator for assistance.
"@ | Out-File -FilePath "$distFolder\USER_GUIDE.txt" -Encoding UTF8
    
    # Get file size
    $exeSize = (Get-Item "dist\TimeOffCalendar.exe").Length / 1MB
    
    Write-Host "`n================================" -ForegroundColor Green
    Write-Host "Distribution package ready!" -ForegroundColor Green
    Write-Host "================================" -ForegroundColor Green
    Write-Host "Location: $distFolder\" -ForegroundColor Yellow
    Write-Host "Executable size: $([math]::Round($exeSize, 2)) MB" -ForegroundColor Yellow
    Write-Host "`nYou can now distribute the '$distFolder' folder to users." -ForegroundColor Cyan
    
    # Clean up build artifacts
    Write-Host "`nCleaning up build artifacts..." -ForegroundColor Cyan
    Remove-Item -Recurse -Force "build" -ErrorAction SilentlyContinue
    Remove-Item -Force "*.spec" -ErrorAction SilentlyContinue
    
} else {
    Write-Host "`n✗ Build failed!" -ForegroundColor Red
    Write-Host "Check the output above for errors." -ForegroundColor Yellow
}

Write-Host "`nDone!" -ForegroundColor Green
