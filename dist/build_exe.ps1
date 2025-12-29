# Build script for creating standalone .exe with PyInstaller

Write-Host "TimeOffCalendar - Build Script" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green
Write-Host ""

# Detect if running from root or build directory based on current location
$currentDir = Get-Location
$currentDirName = Split-Path -Leaf $currentDir

# Set paths based on where we are running from
if ($currentDirName -eq "build") {
    # Running from build directory
    $srcPath = "..\src\excel_to_outlook.py"
    $readmePath = "..\README.md"
} else {
    # Running from root directory
    $srcPath = "src\excel_to_outlook.py"
    $readmePath = "README.md"
}

Write-Host "Running from: $currentDir" -ForegroundColor Cyan
Write-Host "Source file: $srcPath" -ForegroundColor Cyan
Write-Host ""

# Check if PyInstaller is installed
Write-Host "Checking for PyInstaller..." -ForegroundColor Cyan
$pyInstallerCheck = python -m pip show pyinstaller 2>$null
if (-not $pyInstallerCheck) {
    Write-Host "PyInstaller not found. Installing version 5.13.2..." -ForegroundColor Yellow
    python -m pip install pyinstaller==5.13.2
} else {
    Write-Host "PyInstaller found! Installing version 5.13.2 (stable with pandas)..." -ForegroundColor Yellow
    python -m pip install pyinstaller==5.13.2
}

# Clean previous builds (only PyInstaller temp files, not this build/ directory)
Write-Host "`nCleaning previous builds..." -ForegroundColor Cyan
if ($currentDirName -ne "build") {
    # Running from root, safe to delete build folder
    if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
}
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
if (Test-Path "*.spec") { Remove-Item -Force "*.spec" }

# Build the executable
Write-Host "`nBuilding executable..." -ForegroundColor Cyan
Write-Host "This may take a few minutes..." -ForegroundColor Yellow

pyinstaller --onefile `
    --console `
    --name "TimeOffCalendar" `
    --icon NONE `
    $srcPath

# Check if build was successful
if (Test-Path "dist\TimeOffCalendar.exe") {
    Write-Host "`nBuild successful!" -ForegroundColor Green
    
    # Create distribution folder
    $distFolder = "TimeOffCalendar_Distribution"
    Write-Host "`nCreating distribution package..." -ForegroundColor Cyan
    New-Item -ItemType Directory -Force -Path $distFolder | Out-Null
    
    # Copy files
    Copy-Item "dist\TimeOffCalendar.exe" "$distFolder\"
    Copy-Item "$readmePath" "$distFolder\" -ErrorAction SilentlyContinue
    
    # Get file size
    $exeSize = (Get-Item "dist\TimeOffCalendar.exe").Length / 1MB
    
    Write-Host "`n================================" -ForegroundColor Green
    Write-Host "Distribution package ready!" -ForegroundColor Green
    Write-Host "================================" -ForegroundColor Green
    Write-Host "Location: $distFolder\" -ForegroundColor Yellow
    Write-Host "Executable size: $([math]::Round($exeSize, 2)) MB" -ForegroundColor Yellow
    Write-Host "`nYou can now distribute the $distFolder folder to users." -ForegroundColor Cyan
    
    # Clean up PyInstaller build artifacts
    Write-Host "`nCleaning up build artifacts..." -ForegroundColor Cyan
    # Only delete the 'build' folder if we're in the build directory (PyInstaller temp files)
    if ($currentDirName -eq "build") {
        Remove-Item -Recurse -Force "build" -ErrorAction SilentlyContinue
    }
    Remove-Item -Force "*.spec" -ErrorAction SilentlyContinue
    
} else {
    Write-Host "`nX Build failed!" -ForegroundColor Red
    Write-Host "Check the output above for errors." -ForegroundColor Yellow
}

Write-Host "`nDone!" -ForegroundColor Green
