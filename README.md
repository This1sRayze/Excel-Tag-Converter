# Excel Tag Converter

A professional GUI application for converting and organizing Excel tag data with automatic SCADA signal generation and formatting.

## Overview

Excel Tag Converter is a desktop application built with Python and tkinter that processes industrial tag data from Excel files. It allows you to:

- **Convert tag data** from input Excel files into organized area-based sheets
- **Map UDT types** to signal types using a mapping Excel file
- **Auto-generate SCADA signals** with intelligent categorization (ANALOG, DIGITAL, COMM, CALCULATED)
- **Apply professional formatting** with color-coding, borders, and styling
- **Handle array types** with automatic expansion and proper indexing
- **Export organized data** with multiple output sheets per area

## Features

- 🎨 **Professional Formatting**: Color-coded sheets with consistent styling and auto-fit columns
- 📊 **Area-based Organization**: Automatically creates separate sheets for each area
- 🔧 **UDT Type Mapping**: Map complex UDT types to signal types using a separate mapping file
- 🚀 **SCADA Signal Generation**: Automatically generates SCADA_SIGNAL sheet with formatted signal names
- 📈 **Array Type Support**: Handles ARRAY[n..m] OF TYPE declarations with automatic expansion
- 🎯 **Signal Categorization**: Intelligent categorization of signals (ANALOG, DIGITAL, COMM, CALCULATED)
- 📝 **Processing Log**: Real-time log display showing all processing steps
- 💾 **Custom Output**: Choose output filename and save location

## Requirements

See [requirements.txt](requirements.txt) for Python package dependencies.

**System Requirements:**
- Python 3.7 or higher
- Windows, macOS, or Linux
- 100MB free disk space

## Installation

### Option 1: Use the Executable (Easiest)

If you prefer not to install Python, simply download and run the `.exe` file:

1. Download `ExcelTagConverter.exe` from the releases
2. Double-click to run
3. No installation or dependencies required

**Note:** The `.exe` file includes Python and all dependencies bundled together.

### Option 2: Run from Python Source

For developers or those with Python installed:

1. Clone or download the project files:
```bash
git clone <repository-url>
cd Excel\ Tag\ Converter
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python ExcelTagConverter.py
```

## Usage

### Running the Application

**Using the `.exe` (Easiest):**
```bash
ExcelTagConverter.exe
```

**Using Python:**
```bash
python ExcelTagConverter.py
```

### Input Files

#### 1. **Input Excel File** (Required)
Your input data file with the following columns (column names are customizable):
- **Tag Name**: Identifier for the tag
- **Data Block**: The data block reference
- **UDT Type**: The User Defined Type (e.g., ANL, DIG_ALR, ARRAY[0..16] OF ANL)
- **Area**: Area classification (e.g., Engine Room, Bridge, Pump Room)
- **Description**: Tag description
- **Comments** (Optional): Additional comments
- **Origin** (Optional): Tag origin information

#### 2. **Mapping File** (Optional, Required for SCADA Generation)
An Excel file that maps UDT types to signal types with data type definitions:

| UDT Type | Signal Type | Data Type |
|----------|-------------|-----------|
| ANL | Status | REAL |
| ANL | HiAlarm | REAL |
| ANL | LoAlarm | REAL |
| DIG_ALR | Status | BOOL |
| DIG_ALR | Alarm | BOOL |

**Important Notes on Mapping File:**
- Required columns: `UDT Type`, `Signal Type`, `Data Type`
- Support array types: Use format `ARRAY[start..end] OF BASETYPE`
- `Data Type` column determines the exported data type for each signal
- All required columns must be present

### Configuration

In the GUI, configure the column names to match your input file:

- **Tag Name Column**: Column name in input file (default: "Tag Name")
- **Data Block Column**: Column name in input file (default: "Data Block")
- **Description Column**: Column name in input file (default: "Description")
- **UDT Type Column**: Column name in input file (default: "UDT Type")
- **Area Column**: Column name in input file (default: "Area")
- **Comments Column**: Column name in input file (default: "Comments")
- **Origin Column**: Column name in input file (default: "Origin")

### Processing Steps

1. **Select Input Excel**: Click "Select Input Excel" and choose your data file
2. **Select Mapping File** (Optional): Click "Select Mapping Excel" to load UDT type mappings
3. **Configure Columns**: Adjust column names if they differ from defaults
4. **Process**: Click "🚀 Process and Convert" button
5. **Choose Output Location**: Select where to save the processed file
6. **Review**: Check the Processing Log for status updates
7. **Output**: Output folder automatically opens when complete

## Output Structure

The processed Excel file contains:

### Area Sheets
- One sheet per area from your input data
- Columns: Data Block, Tag Name, UDT Type, Signal Type, Comments, Is Alarm, Alarm Priority, Tag History, Origin, Description
- Color-coded by area
- Professional formatting with borders and styling

### SCADA_SIGNAL Sheet (If Mapping File Provided)
- Comprehensive signal sheet for SCADA export
- Columns: DB, Scada Tag Path, Type, Signal Type, Data Type, Comments, Origin, Description
- Automatically generated signal tags with proper naming
- Sorted by area and array indices
- Color-coded by area

## Signal Type Categories

The application automatically categorizes signals:

- **ANALOG**: ANL, ANL_TANK types
- **DIGITAL**: DIG_ALR, DIG_ALR_WO_INH, PUMP, BILGE, VALVE types
- **COMM**: Device communication types (DEIF, Automaskin, MTU, Consilium, NMEA, Modbus, GPS)
- **CALCULATED**: Position failure indicators or diagnostics area tags

## Array Type Handling

The application intelligently handles array types:

```
ARRAY[0..16] OF ANL        → Creates 17 array elements (0-16)
ARRAY[1..10] OF DIG_ALR    → Creates 10 array elements (1-10)
```

Array bounds can be defined in either the UDT Type column or the mapping file.

## Troubleshooting

### "Required columns not found"
- Verify your column names in the Configuration section match your input file exactly
- Check for leading/trailing spaces in column names

### Mapping file not loading
- Ensure mapping file has columns: `UDT Type`, `Signal Type`, `Data Type`
- Check that file is in .xlsx or .xls format

### SCADA_SIGNAL sheet not generated
- Ensure mapping file is selected before processing
- Check that UDT types in input match those in mapping file

### Slow processing
- Large files (10,000+ rows) may take several minutes
- Monitor the Processing Log for progress
