# TimeOffCalendar - Deployment Guide

## Building the Executable

### Prerequisites
- Python 3.8 or later installed
- All dependencies from requirements.txt installed

### Build Steps

1. **Install PyInstaller:**
   ```powershell
   pip install pyinstaller
   ```

2. **Run the build script:**
   ```powershell
   build\build_exe.ps1
   ```

3. **Find your distribution:**
   - Executable will be in `TimeOffCalendar_Distribution\`
   - Size will be approximately 30-50 MB

### CLI Options (for .exe or script)

The program now supports the following command-line options:

- `--clear` : Clear the calendar before importing
- `--verbose` : Include reason code in event titles
- `--range MM-DD-YYYY MM-DD-YYYY` : Only import events within a date range

See docs/README.md for usage examples.

## Distribution

### For Small Scale (< 10 users)
- Zip the `TimeOffCalendar_Distribution` folder
- Email to users or share via network drive
- Users extract and run

### For Enterprise Deployment

#### Option 1: Network Share
```
\\company-share\Tools\TimeOffCalendar\
├── TimeOffCalendar.exe
└── USER_GUIDE.txt
```
Users copy to their machine and run.

#### Option 2: SCCM/Intune
- Package the .exe as an application
- Deploy to specific user groups
- Set as available (user-initiated) or required

#### Option 3: Email Distribution
- Zip the folder (will be ~30-50 MB)
- Send via email with instructions
- Note: Some email systems block .exe files - rename to .ex_ if needed

## Security Considerations

### Windows SmartScreen
Unsigned executables will trigger SmartScreen warnings. Users will see:
- "Windows protected your PC"
- They must click "More info" → "Run anyway"

**Solutions:**
1. **Code Signing (Recommended for Enterprise):**
   - Purchase a code signing certificate (~$200-500/year)
   - Sign the executable using signtool.exe
   - Eliminates SmartScreen warnings
2. **Group Policy:**
   - IT can whitelist the executable via Group Policy
   - Add to trusted applications list
3. **User Training:**
   - Provide clear instructions on bypassing SmartScreen
   - Verify the executable came from trusted source

### Antivirus Software
Some AV may flag PyInstaller executables as suspicious. Solutions:
- Submit to AV vendors for whitelisting
- Add exception in enterprise AV policy
- Use code signing (significantly reduces false positives)

## Updating

### Version Control
In `src/excel_to_outlook.py`, add version info:
```python
__version__ = "1.0.0"
print(f"TimeOffCalendar v{__version__}")
```

### Distributing Updates
1. Build new executable with updated code
2. Increment version number
3. Distribute via same channels
4. Users replace old .exe with new one

## User Installation

Users need to:
1. Copy folder anywhere on their PC
2. Ensure Outlook is installed
3. Place Excel file in same folder (or specify path)
4. Run TimeOffCalendar.exe

**No admin rights required** - runs from user space.

## Troubleshooting

### Build Issues

**"PyInstaller not found":**
```powershell
pip install pyinstaller
```

**"Module not found during build":**
```powershell
pip install -r requirements.txt
pyinstaller --onefile --hidden-import=MODULE_NAME src/excel_to_outlook.py
```

**Build succeeds but .exe crashes:**
- Test on a clean Windows VM
- Check for missing dependencies
- Add hidden imports to build script

### Runtime Issues

**"Missing DLL":**
- PyInstaller should bundle everything, but some COM objects need registration
- Ensure Outlook is installed on target machine

**Large File Size:**
- Normal for PyInstaller (~30-50MB)
- Includes Python interpreter and all libraries
- Cannot be reduced significantly without compromising functionality

## Advanced: Custom Build Options

### Add an Icon
1. Create or obtain a .ico file
2. Modify build script:
   ```powershell
   pyinstaller --onefile --icon=icon.ico src/excel_to_outlook.py
   ```

### Console vs Windowed
- Current: `--console` (shows progress window)
- Alternative: `--windowed` (no console, runs silently)

### Add Data Files
```powershell
pyinstaller --onefile --add-data "config.json;." src/excel_to_outlook.py
```

### Create Spec File for Custom Build
```powershell
pyi-makespec --onefile src/excel_to_outlook.py
# Edit excel_to_outlook.spec
pyinstaller excel_to_outlook.spec
```

## Testing Checklist

Before distributing:
- [ ] Test on clean Windows 10 VM
- [ ] Test on clean Windows 11 VM
- [ ] Verify Outlook interaction works
- [ ] Test with sample Excel file
- [ ] Check SmartScreen behavior
- [ ] Verify file size is reasonable
- [ ] Test without Python installed
- [ ] Document any warnings/prompts users will see

## Support

For build issues or questions, contact the development team.
