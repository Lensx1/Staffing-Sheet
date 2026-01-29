# Staffing Sheet Generator

A Python application that converts HR export files into a formatted, printable staffing list optimized for 17x11 (ledger) paper.

## Overview

This tool replaces the Excel VBA-based staffing sheet generator with a clean Python implementation. It processes two HR export files, merges employee data, filters to active employees, and generates a professionally formatted Excel file ready for printing.

## Features

- ✅ Reads multiple HR export file formats (.xls, .xlsx, .csv)
- ✅ Automatically detects and handles files with extra header rows
- ✅ Merges and deduplicates employee data from multiple sources
- ✅ Filters to show only active employees
- ✅ Extracts and displays status codes:
  - **L** = Line lead
  - **S** = Safety Team
  - **P** = Part Time
  - **O** = Offsite employee
  - **T** = Trainee/Temp (Not In Payroll)
- ✅ Organizes employees by department into three major sections
- ✅ Generates print-ready Excel output for 17x11 landscape paper
- ✅ Includes generation date in output

## Requirements

- Python 3.8 or higher
- pandas
- openpyxl
- xlrd (for reading .xls files)

## Installation

1. Clone or download this repository

2. Install required Python packages:
   ```bash
   pip install -r requirements.txt
   ```

   Or install individually:
   ```bash
   pip install pandas openpyxl xlrd
   ```

## Usage

### Basic Usage

Run with default input files in the current directory:

```bash
python main.py
```

This will:
- Read `Emailed - Staffing Emergency Employee List - no grouping.xls`
- Read `EmployeeInformation-EmergencyEmployeeList-nogrouping.xlsx`
- Generate `Staffing_Sheet_[DATE].xlsx`

### Custom Input Files

Specify custom input file paths:

```bash
python main.py --file1 path/to/first_file.xlsx --file2 path/to/second_file.xlsx
```

### Custom Output File

Specify custom output filename:

```bash
python main.py --output MyStaffingSheet.xlsx
```

### Complete Example

```bash
python main.py \
  --file1 "HR Exports/Emailed - Staffing Emergency Employee List - no grouping.xls" \
  --file2 "HR Exports/EmployeeInformation-EmergencyEmployeeList-nogrouping.xlsx" \
  --output "Staffing_Sheet_January_2026.xlsx"
```

### Get Help

```bash
python main.py --help
```

## Input File Format

The application expects two CSV/Excel files with the following columns:

- Default Department
- Work Schedule
- Employee Id
- First Name
- Last Name
- Employee Status
- Jobs (HR)(1)
- Primary Status
- Secondary Status

Both files should have the same structure. The application will:
- Automatically detect header rows (even if there are extra rows before the actual headers)
- Merge data from both files
- Use Employee Id as the unique key for deduplication

## Output Format

The generated Excel file includes:

### Layout
- **3-column layout**: Office, Fabrication, Finishing sections side-by-side
- **Page setup**: 17x11 (tabloid/ledger) paper, landscape orientation
- **Print-ready**: Optimized margins and column widths

### Organization

#### Office Section (Left)
- Executives
- HR
- Customer Service
- Estimating
- Office
- Engineering
- Quality
- Pack/Ship
- Inventory
- Accounting
- Purchasing
- Sales

#### Fabrication Section (Middle)
- Cutting
- Bending
- Welding Manager
- Small Weld
- Med Weld
- Large Weld
- SS Weld (Stainless Steel)
- Maintenance

#### Finishing Section (Right)
- Finishing Manager
- Paint
- Assembly
- MFG Engineering
- Float
- Systems
- Scheduling
- Operations Manager

### Employee Information

For each employee:
- **Name**: First and Last name
- **Shift**: 1, 2, 3, or P (Part Time)
- **ID #**: Employee ID number
- **Title**: Simplified job title (without department prefix)
- **Codes**: Status codes (L, S, P, O, T)

## Department Mapping

The application maps department codes to sections:

| Department Code | Section | Subsection |
|----------------|---------|------------|
| EXEC | Office | Executives |
| HR | Office | HR |
| CS | Office | Cust Service |
| ESTIMATIONS | Office | Estimating |
| OFFICE | Office | Office |
| ENGR | Office | Engineering |
| QUALITY | Office | Quality |
| SHIP | Office | Pack/Ship |
| INVENTORY | Office | Inventory |
| FINC | Office | Accounting |
| CUT | Fabrication | Cutting |
| BEND | Fabrication | Bending |
| PCO | Fabrication | Welding Manager |
| WLDSM | Fabrication | Small Weld |
| WLDMD | Fabrication | Med Weld |
| WLDLG | Fabrication | Large Weld |
| WLDSS | Fabrication | SS Weld |
| MAINT | Fabrication | Maintenance |
| PAINT | Finishing | Paint |
| ASMBL | Finishing | Assembly |
| LPO | Finishing | MFG Engineering |
| GPO | Finishing | Float |

## Error Handling

The application handles common issues gracefully:

- **Missing files**: Clear error message with file path
- **Incorrect formats**: Supports .xls, .xlsx, and .csv
- **Extra header rows**: Automatically detects where the real data starts
- **Missing data**: Handles NaN values appropriately
- **Duplicates**: Uses Employee Id to remove duplicates

## Troubleshooting

### "File not found" error
- Check that the input files are in the current directory or provide full paths
- Verify file names match exactly (case-sensitive on Linux/Mac)

### "Module not found" error
- Install required packages: `pip install -r requirements.txt`

### Output looks incorrect
- Verify input files have the expected columns
- Check that Employee Status values are "Active" or "Not In Payroll"
- Ensure department codes match the expected values

### Excel file won't open
- Make sure you have sufficient disk space
- Try a different output filename
- Check that openpyxl is installed: `pip install openpyxl`

## Development

### Project Structure

```
Staffing-Sheet/
├── main.py                 # Main application script
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── .gitignore            # Git ignore rules
└── [Input files]         # HR export files (not in git)
```

### Making Changes

1. Edit `main.py` to modify logic
2. Update `requirements.txt` if adding dependencies
3. Test with sample data files
4. Update this README if changing usage

## License

This project is for internal use by Schaefer's Electrical Enclosures.

## Support

For issues or questions, contact the IT department or the repository maintainer.