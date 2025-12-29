# Build script for creating standalone .exe with PyInstaller

Write-Host "TimeOffCalendar - Build Script" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green
Write-Host ""

# Always run from root directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir = Split-Path -Parent $scriptDir
Set-Location $rootDir

Write-Host "Running from: $rootDir" -ForegroundColor Cyan
Write-Host ""

# Check if PyInstaller is installed
Write-Host "Checking for PyInstaller..." -ForegroundColor Cyan
$pyInstallerCheck = python -m pip show pyinstaller 2>$null
if (-not $pyInstallerCheck) {
    Write-Host "PyInstaller not found. Installing version 5.13.2..." -ForegroundColor Yellow
    python -m pip install pyinstaller==5.13.2
} else {
    Write-Host "PyInstaller found! Installing version 5.13.2..." -ForegroundColor Yellow
    python -m pip install pyinstaller==5.13.2
}

# Clean previous builds (but NOT the dist folder or this script)
Write-Host "`nCleaning previous builds..." -ForegroundColor Cyan
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "dist\TimeOffCalendar.exe") { Remove-Item -Force "dist\TimeOffCalendar.exe" }
if (Test-Path "*.spec") { Remove-Item -Force "*.spec" }

# Build the executable
Write-Host "`nBuilding executable..." -ForegroundColor Cyan
Write-Host "This may take a few minutes..." -ForegroundColor Yellow

pyinstaller --onefile `
    --console `
    --name "TimeOffCalendar" `
    --icon NONE `
    --distpath dist `
    src\excel_to_outlook.py

# Check if build was successful
if (Test-Path "dist\TimeOffCalendar.exe") {
    Write-Host "`nBuild successful!" -ForegroundColor Green
    
    # Get file size
    $exeSize = (Get-Item "dist\TimeOffCalendar.exe").Length / 1MB
    
    Write-Host "`n================================" -ForegroundColor Green
    Write-Host "Build complete!" -ForegroundColor Green
    Write-Host "================================" -ForegroundColor Green
    Write-Host "Location: dist\TimeOffCalendar.exe" -ForegroundColor Yellow
    Write-Host "Executable size: $([math]::Round($exeSize, 2)) MB" -ForegroundColor Yellow
    
    # Clean up PyInstaller build artifacts (but NOT dist folder)
    Write-Host "`nCleaning up build artifacts..." -ForegroundColor Cyan
    if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
    if (Test-Path "*.spec") { Remove-Item -Force "*.spec" }
    
} else {
    Write-Host "`nX Build failed!" -ForegroundColor Red
    Write-Host "Check the output above for errors." -ForegroundColor Yellow
}

Write-Host "`nDone!" -ForegroundColor Green
