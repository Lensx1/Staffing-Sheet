#!/bin/bash
# Build script for creating StaffingSheetGenerator executable on Linux/Mac
# This script checks for PyInstaller and builds the executable

echo "========================================"
echo "Staffing Sheet Generator - Build Script"
echo "========================================"
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed or not in PATH"
    echo "Please install Python 3.8 or higher"
    echo "  Ubuntu/Debian: sudo apt install python3 python3-pip"
    echo "  macOS: brew install python3"
    exit 1
fi

echo "Python found:"
python3 --version
echo ""

# Check if pip is available
if ! python3 -m pip --version &> /dev/null; then
    echo "ERROR: pip is not available"
    echo "Please install pip for Python 3"
    exit 1
fi

# Check if PyInstaller is installed
if ! python3 -c "import PyInstaller" &> /dev/null; then
    echo "PyInstaller is not installed. Installing now..."
    echo ""
    python3 -m pip install pyinstaller>=5.0
    if [ $? -ne 0 ]; then
        echo "ERROR: Failed to install PyInstaller"
        exit 1
    fi
    echo ""
    echo "PyInstaller installed successfully!"
    echo ""
else
    echo "PyInstaller is already installed"
    echo ""
fi

# Check if required packages are installed
echo "Checking for required dependencies..."
if ! python3 -c "import pandas, openpyxl, xlrd" &> /dev/null; then
    echo "Installing required dependencies..."
    python3 -m pip install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "ERROR: Failed to install dependencies"
        exit 1
    fi
    echo ""
fi

# Clean previous build artifacts
echo "Cleaning previous build artifacts..."
rm -rf build dist
echo ""

# Build the executable
echo "Building StaffingSheetGenerator executable..."
echo "This may take a few minutes..."
echo ""
python3 -m PyInstaller build.spec

if [ $? -ne 0 ]; then
    echo ""
    echo "========================================"
    echo "ERROR: Build failed!"
    echo "========================================"
    echo "Please check the error messages above."
    exit 1
fi

echo ""
echo "========================================"
echo "Build completed successfully!"
echo "========================================"
echo ""
echo "The executable is located at:"
echo "  dist/StaffingSheetGenerator"
echo ""
echo "You can now:"
echo "1. Copy dist/StaffingSheetGenerator to any folder"
echo "2. Place your HR export files in the same folder"
echo "3. Run ./StaffingSheetGenerator (you may need to chmod +x it first)"
echo ""
echo "Note: The executable is approximately 30-50 MB due to bundled libraries."
echo ""
