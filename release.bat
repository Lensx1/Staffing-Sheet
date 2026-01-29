@echo off
REM Release script - Builds the executable and creates a distribution package
REM This creates a ready-to-share ZIP file with the executable and documentation

echo ========================================
echo Staffing Sheet Generator - Release Build
echo ========================================
echo.

REM First, build the executable
echo Step 1/4: Building executable...
echo.
call build_exe.bat
if errorlevel 1 (
    echo.
    echo ERROR: Build failed. Release aborted.
    pause
    exit /b 1
)

REM Create release directory
echo.
echo Step 2/4: Creating release directory...
if exist release rmdir /s /q release
mkdir release
echo Release directory created.

REM Copy executable
echo.
echo Step 3/4: Copying files to release directory...
copy dist\StaffingSheetGenerator.exe release\ >nul
if errorlevel 1 (
    echo ERROR: Failed to copy executable
    pause
    exit /b 1
)

REM Create distribution README
echo.
echo Creating distribution README...
(
echo Staffing Sheet Generator
echo ========================
echo.
echo QUICK START:
echo 1. Place this executable in a folder with your HR export files
echo 2. Double-click StaffingSheetGenerator.exe to run
echo 3. The formatted staffing sheet will be created as Staffing_Sheet_[DATE].xlsx
echo.
echo REQUIREMENTS:
echo - Windows 10/11 ^(64-bit^)
echo - No Python installation required
echo - No admin privileges needed
echo.
echo DEFAULT INPUT FILES:
echo The program looks for these files in the same folder:
echo   - Emailed - Staffing Emergency Employee List - no grouping.xls
echo   - EmployeeInformation-EmergencyEmployeeList-nogrouping.xlsx
echo.
echo CUSTOM FILES:
echo To use different input files, open Command Prompt in this folder and run:
echo   StaffingSheetGenerator.exe --file1 yourfile1.xlsx --file2 yourfile2.xlsx
echo.
echo CUSTOM OUTPUT:
echo To specify a custom output filename:
echo   StaffingSheetGenerator.exe --output MyReport.xlsx
echo.
echo COMMAND-LINE HELP:
echo For more options, run:
echo   StaffingSheetGenerator.exe --help
echo.
echo SUPPORT:
echo For issues or questions, please refer to the BUILD.md file in the
echo source repository or contact your IT department.
echo.
echo FILE SIZE NOTE:
echo This executable is 30-50 MB because it includes Python and all required
echo libraries. This allows it to run without any installation.
echo.
echo ANTIVIRUS NOTE:
echo Some antivirus software may flag PyInstaller executables as suspicious.
echo This is a false positive. The executable contains only Python and
echo standard data processing libraries. If needed, add an exclusion in
echo your antivirus software or contact your IT department.
echo.
echo Version: 1.0
echo Built: %date% %time%
) > release\README.txt

echo Distribution README created.

REM Create release package info
(
echo Release Package Contents
echo ========================
echo.
echo Files in this package:
echo   - StaffingSheetGenerator.exe   Main executable
echo   - README.txt                   User guide
echo.
echo Build Information:
echo   Built on: %date% %time%
echo   Built with: PyInstaller
) > release\RELEASE_INFO.txt

REM Add Python version to release info
for /f "delims=" %%v in ('python --version 2^>^&1') do echo   Python version used: %%v >> release\RELEASE_INFO.txt

REM Create a zip file (requires PowerShell)
echo.
echo Step 4/4: Creating ZIP archive...
for /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c-%%a-%%b)
set ZIPNAME=StaffingSheetGenerator-v1.0-Windows-%mydate%.zip

powershell -command "Compress-Archive -Path release\* -DestinationPath '%ZIPNAME%' -Force"
if errorlevel 1 (
    echo WARNING: Failed to create ZIP file. You can manually zip the release folder.
) else (
    echo ZIP archive created: %ZIPNAME%
)

echo.
echo ========================================
echo Release Package Created Successfully!
echo ========================================
echo.
echo Release contents:
echo   Folder: release\
echo     - StaffingSheetGenerator.exe
echo     - README.txt
echo     - RELEASE_INFO.txt
echo.
if exist %ZIPNAME% (
    echo ZIP file: %ZIPNAME%
    echo.
)
echo You can now:
echo 1. Test the executable in the release\ folder
echo 2. Share the release\ folder or the ZIP file
echo 3. Upload to file sharing service or internal network drive
echo.
echo TESTING RECOMMENDATION:
echo Before distributing, test the executable on a machine without
echo Python installed to ensure it works correctly.
echo.
pause
