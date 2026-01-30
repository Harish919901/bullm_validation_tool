# Excel Validation Tool - Quote Win Files

A validation tool for Quote Win Excel files in CAM with both CLI and GUI interfaces.

## Features

- Validates headers against embedded template configuration
- Checks project name consistency
- Validates filters are not applied
- Validates award assignments for all part numbers
- Highlights issues with yellow cells and adds comments
- Generates validated output file

## Requirements

```bash
pip install openpyxl
```

## Usage

### GUI Interface (Recommended)

Run the graphical user interface:

```bash
python excel_validation_gui.py
```

The GUI provides:
- File selection with browse buttons
- Progress tracking during validation
- Detailed validation results display
- Button to open the validated output file

### Command Line Interface

Run validation from the command line:

```bash
python excel_validation_tool.py input.xlsx
python excel_validation_tool.py input.xlsx output_validated.xlsx
```

## Validation Checks

1. **Header Validation**: Verifies all required static and dynamic headers are present in both Row 12 (summary) and Row 16 (main headers)
2. **Project Name Validation**: Ensures project name matches between Row 3 and the data section
3. **Filter Validation**: Checks that no filters are applied to the spreadsheet
4. **Award Validation**: Verifies that each unique part number has at least one award with value "100"

## Output

The tool generates an output file with:
- Yellow highlighted cells indicating issues
- Comments explaining each validation failure
- Original data preserved

## Template Configuration

The template is embedded in the tool with the following key configurations:
- Summary header row: 12
- Main header row: 16
- Data start row: 17
- Project row: 3

Required headers include both static headers (e.g., "Part Number", "Project") and dynamic pattern-based headers (e.g., "Cost #X (Conv.)", "Award #X").
