@echo off
REM Build script for creating StaffingSheetGenerator.exe on Windows
REM This script checks for PyInstaller and builds the executable

echo ========================================
echo Staffing Sheet Generator - Build Script
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo Python found:
python --version
echo.

REM Check if PyInstaller is installed
python -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo PyInstaller is not installed. Installing now...
    echo.
    python -m pip install pyinstaller>=5.0
    if errorlevel 1 (
        echo ERROR: Failed to install PyInstaller
        pause
        exit /b 1
    )
    echo.
    echo PyInstaller installed successfully!
    echo.
) else (
    echo PyInstaller is already installed
    echo.
)

REM Check if required packages are installed
echo Checking for required dependencies...
python -c "import pandas, openpyxl, xlrd" >nul 2>&1
if errorlevel 1 (
    echo Installing required dependencies...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies
        pause
        exit /b 1
    )
    echo.
)

REM Clean previous build artifacts
echo Cleaning previous build artifacts...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo.

REM Build the executable
echo Building StaffingSheetGenerator.exe...
echo This may take a few minutes...
echo.
python -m PyInstaller build.spec

if errorlevel 1 (
    echo.
    echo ========================================
    echo ERROR: Build failed!
    echo ========================================
    echo Please check the error messages above.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build completed successfully!
echo ========================================
echo.
echo The executable is located at:
echo   dist\StaffingSheetGenerator.exe
echo.
echo You can now:
echo 1. Copy StaffingSheetGenerator.exe to any folder
echo 2. Place your HR export files in the same folder
echo 3. Run StaffingSheetGenerator.exe (double-click or from command line)
echo.
echo Note: The .exe file is approximately 30-50 MB due to bundled libraries.
echo.
pause
