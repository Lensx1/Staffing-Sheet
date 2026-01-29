# Building the Staffing Sheet Generator Executable

This guide explains how to build the `StaffingSheetGenerator.exe` (Windows) or `StaffingSheetGenerator` (Linux/Mac) executable from source code.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Building on Windows](#building-on-windows)
3. [Building on Linux/Mac](#building-on-linuxmac)
4. [Using the Executable](#using-the-executable)
5. [Distribution](#distribution)
6. [Troubleshooting](#troubleshooting)

---

## Prerequisites

### System Requirements

- **Python**: 3.8 or higher
- **Disk Space**: ~500 MB for build process (final executable is 30-50 MB)
- **Memory**: 2 GB RAM minimum
- **Operating System**:
  - Windows: Windows 10/11 (64-bit)
  - Linux: Ubuntu 20.04+ or equivalent
  - macOS: macOS 10.15+

### Required Python Packages

All packages will be automatically installed by the build script, but you can also install them manually:

```bash
pip install pyinstaller>=5.0
pip install -r requirements.txt
```

The `requirements.txt` includes:
- pandas>=2.0.0
- openpyxl>=3.1.0
- xlrd>=2.0.0

---

## Building on Windows

### Quick Build

1. Open Command Prompt or PowerShell in the project directory
2. Run the build script:
   ```cmd
   build_exe.bat
   ```
3. Wait for the build to complete (typically 2-5 minutes)
4. The executable will be in `dist\StaffingSheetGenerator.exe`

### What the Build Script Does

The `build_exe.bat` script automatically:
1. ✅ Checks if Python is installed
2. ✅ Installs PyInstaller if not already present
3. ✅ Installs required dependencies (pandas, openpyxl, xlrd)
4. ✅ Cleans previous build artifacts
5. ✅ Runs PyInstaller with the configured spec file
6. ✅ Reports success/failure with clear messages

### Manual Build (Advanced)

If you prefer to build manually:

```cmd
# Install dependencies
pip install pyinstaller>=5.0
pip install -r requirements.txt

# Clean previous builds
rmdir /s /q build dist

# Build with spec file
pyinstaller build.spec
```

---

## Building on Linux/Mac

### Quick Build

1. Open Terminal in the project directory
2. Make the build script executable (first time only):
   ```bash
   chmod +x build_exe.sh
   ```
3. Run the build script:
   ```bash
   ./build_exe.sh
   ```
4. Wait for the build to complete (typically 2-5 minutes)
5. The executable will be in `dist/StaffingSheetGenerator`

### What the Build Script Does

The `build_exe.sh` script automatically:
1. ✅ Checks if Python 3 is installed
2. ✅ Installs PyInstaller if not already present
3. ✅ Installs required dependencies (pandas, openpyxl, xlrd)
4. ✅ Cleans previous build artifacts
5. ✅ Runs PyInstaller with the configured spec file
6. ✅ Reports success/failure with clear messages

### Manual Build (Advanced)

If you prefer to build manually:

```bash
# Install dependencies
pip3 install pyinstaller>=5.0
pip3 install -r requirements.txt

# Clean previous builds
rm -rf build dist

# Build with spec file
pyinstaller build.spec
```

---

## Using the Executable

The executable works identically to the Python script but doesn't require Python to be installed.

### Basic Usage

**Windows:**
```cmd
# Place the .exe in a folder with your HR export files
StaffingSheetGenerator.exe
```

**Linux/Mac:**
```bash
# Make sure the executable has execute permissions
chmod +x StaffingSheetGenerator
./StaffingSheetGenerator
```

### With Command-Line Arguments

The executable supports all the same arguments as the Python script:

```bash
# Custom input files
StaffingSheetGenerator.exe --file1 myfile1.xlsx --file2 myfile2.xlsx

# Custom output file
StaffingSheetGenerator.exe --output MyReport.xlsx

# Complete example
StaffingSheetGenerator.exe \
  --file1 "HR Export 1.xlsx" \
  --file2 "HR Export 2.xlsx" \
  --output "Staffing_Report_2026.xlsx"
```

### Default Behavior

When run without arguments, the executable:
1. Looks for `Emailed - Staffing Emergency Employee List - no grouping.xls`
2. Looks for `EmployeeInformation-EmergencyEmployeeList-nogrouping.xlsx`
3. Creates `Staffing_Sheet_[DATE].xlsx` with current date

---

## Distribution

### What to Include

When sharing the executable with others, include:

1. **Required:**
   - `StaffingSheetGenerator.exe` (Windows) or `StaffingSheetGenerator` (Linux/Mac)
   - Brief instructions (see sample below)

2. **Optional but Recommended:**
   - Sample README with basic usage instructions
   - Example HR export files (without sensitive data)

### Sample Distribution README

Create a simple README.txt for distribution:

```
Staffing Sheet Generator v1.0
==============================

USAGE:
1. Place this executable in a folder with your HR export files
2. Double-click StaffingSheetGenerator.exe to run
3. The formatted staffing sheet will be created as Staffing_Sheet_[DATE].xlsx

REQUIREMENTS:
- Windows 10/11 (64-bit)
- No Python installation required
- No admin privileges needed

DEFAULT INPUT FILES:
- Emailed - Staffing Emergency Employee List - no grouping.xls
- EmployeeInformation-EmergencyEmployeeList-nogrouping.xlsx

CUSTOM FILES:
To use different input files, run from command prompt:
  StaffingSheetGenerator.exe --file1 yourfile1.xlsx --file2 yourfile2.xlsx

SUPPORT:
For issues or questions, contact [your contact info]
```

### Creating a Distribution Package

For easy distribution, create a ZIP file:

**Windows:**
```cmd
# Create a release folder
mkdir release
copy dist\StaffingSheetGenerator.exe release\
copy README_Distribution.txt release\README.txt

# Create ZIP (requires 7-zip or similar)
7z a StaffingSheetGenerator-v1.0-Windows.zip release\*
```

**Linux/Mac:**
```bash
# Create a release folder
mkdir release
cp dist/StaffingSheetGenerator release/
cp README_Distribution.txt release/README.txt

# Create tarball
tar -czf StaffingSheetGenerator-v1.0-$(uname -s).tar.gz release/
```

### File Size Note

The executable is approximately **30-50 MB** because it includes:
- Python interpreter
- pandas library
- openpyxl library
- xlrd library
- All their dependencies

This is normal for PyInstaller executables and allows the program to run on systems without Python installed.

---

## Troubleshooting

### Common Issues

#### 1. Antivirus Warnings

**Problem:** Antivirus software flags the executable as suspicious.

**Cause:** PyInstaller executables are sometimes flagged because they bundle a Python interpreter and decompress code at runtime - behavior that antivirus heuristics may flag.

**Solutions:**
- Add an exclusion for the executable in your antivirus software
- Build the executable yourself from trusted source code
- For corporate distribution, submit the executable to your IT department for whitelisting
- Some companies sign their executables with a code signing certificate to avoid this

**Note:** This is a common issue with PyInstaller and not indicative of malware. Building from source yourself ensures the executable contains only your code.

#### 2. "Failed to execute script" Error

**Problem:** Executable runs but shows "Failed to execute script" error.

**Solutions:**
- Ensure input files are in the same directory as the executable
- Check that input files are readable (not locked by Excel or another program)
- Try running from command prompt to see detailed error messages
- Verify file paths don't contain special characters that need escaping

#### 3. Missing Dependencies During Build

**Problem:** Build fails with "Module not found" errors.

**Solutions:**
```bash
# Reinstall all dependencies
pip install --force-reinstall -r requirements.txt

# Upgrade pip first
pip install --upgrade pip

# Then try building again
```

#### 4. Build Takes Very Long

**Problem:** Build process seems stuck or takes more than 10 minutes.

**Solutions:**
- This is sometimes normal on slower systems
- Check Task Manager/Activity Monitor - Python should be using CPU
- If truly stuck (no CPU usage for 5+ minutes), cancel (Ctrl+C) and try again
- Disable antivirus temporarily during build (it may be scanning files as they're created)

#### 5. Executable Won't Run on Target Machine

**Problem:** Executable works on build machine but not on another Windows machine.

**Possible Causes & Solutions:**

**A. Different Architecture:**
- Executable built on 64-bit Python will only run on 64-bit Windows
- Solution: Build with 32-bit Python for wider compatibility (though 64-bit is recommended for modern systems)

**B. Windows Defender SmartScreen:**
- Windows may show "Windows protected your PC" message
- Solution: Click "More info" → "Run anyway"
- For distribution: Code signing certificate prevents this

**C. Missing Visual C++ Redistributables:**
- Rare, but some dependencies may need Visual C++ runtime
- Solution: Install [Microsoft Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe)

**D. Antivirus Quarantine:**
- Antivirus may silently quarantine the executable
- Solution: Check antivirus quarantine, add exclusion

#### 6. Output File Not Created

**Problem:** Executable runs without error but no output file appears.

**Solutions:**
- Check for error messages in the console window
- Verify input files exist and are readable
- Ensure you have write permissions in the current directory
- Check if Excel has the file open from a previous run
- Try specifying full path with `--output`

#### 7. Executable Size Concerns

**Problem:** 40+ MB seems too large for a simple script.

**Explanation:** This is normal! The executable includes:
- Python interpreter (~15 MB)
- pandas library (~20 MB with numpy)
- openpyxl and xlrd
- All dependencies

**If size is critical:**
- Consider distributing the Python script instead
- Use a lighter alternative to pandas (would require code rewrite)
- Accept the size - it's the cost of standalone distribution

---

## Build Configuration

The build is controlled by `build.spec`, which configures:

- **One-file mode**: Single .exe contains everything
- **Console window**: Enabled so users see progress messages
- **Hidden imports**: Ensures pandas, openpyxl, xlrd are included
- **UPX compression**: Reduces file size by ~30%
- **No icon**: Add `icon='path/to/icon.ico'` if you have one

### Customizing the Build

Edit `build.spec` to customize:

```python
# Change executable name
name='YourCustomName',

# Add an icon (must be .ico format for Windows)
icon='path/to/icon.ico',

# Disable console window (not recommended - hides errors)
console=False,

# Add data files if needed
datas=[('config.json', '.')],
```

---

## Advanced Topics

### Code Signing (Corporate/Professional Distribution)

For professional distribution, consider code signing:

1. Obtain a code signing certificate (from Sectigo, DigiCert, etc.)
2. Sign the executable after building:
   ```cmd
   signtool sign /f certificate.pfx /p password /t http://timestamp.digicert.com dist\StaffingSheetGenerator.exe
   ```

Benefits:
- Eliminates antivirus false positives
- Removes Windows SmartScreen warnings
- Shows your company name when users run it

### Automated Builds

For regular releases, create a GitHub Actions workflow or similar CI/CD pipeline:

```yaml
name: Build Executable
on: [push]
jobs:
  build:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v2
      - uses: actions/setup-python@v2
      - run: pip install -r requirements.txt pyinstaller
      - run: pyinstaller build.spec
      - uses: actions/upload-artifact@v2
        with:
          name: StaffingSheetGenerator
          path: dist/
```

---

## Support

If you encounter issues not covered in this guide:

1. Check the PyInstaller documentation: https://pyinstaller.org/
2. Review the main README.md for application-specific help
3. Verify your Python environment is working: `python main.py`
4. Try building in a clean virtual environment

For application-specific issues (not build-related), see README.md.
