# Excel Sheet Cleaner Utility

A Python utility for cleaning Excel spreadsheets by removing empty rows and handling duplicate values in specified columns with user interaction.

## Features

* **Empty Row Removal**: Automatically detects and removes completely empty rows
* **Column A Duplicate Processing**: Handles duplicate values in the first column with intelligent conflict resolution
* **Column B Duplicate Processing**: Interactive duplicate resolution for the second column
* **User-Friendly Interface**: Command-line prompts guide users through the cleaning process
* **Automatic Output**: Saves cleaned data to a new Excel file with "cleaned\_" prefix

## Requirements

* Python 3.6+
* pandas library
* openpyxl (for Excel file handling)

## Installation

1. Clone or download this repository
2. Install required dependencies:

```bash
pip install pandas openpyxl
```

## Usage

1. Run the script:

```bash
python excel_cleaner.py
```

2. Follow the interactive prompts:
    * Enter the path to your Excel file
    * Select a worksheet (if multiple sheets exist)
    * Specify the names of the two columns to process
    * Make decisions about duplicate handling when prompted
3. The cleaned file will be saved as `cleaned_[original_filename].xlsx`

## How It Works

### 1\. Empty Row Removal \(`remove_empty_rows`)

* Scans the entire DataFrame for completely empty rows
* Removes rows where all columns contain NaN values
* Reports the number of rows removed
* Resets index after removal

### 2\. Column A Duplicate Processing \(`process_column_a_duplicates`)

* Identifies duplicate values in the first specified column
* For each group of duplicates:
    * If corresponding Column B values are identical: automatically keeps one row
    * If Column B values differ: prompts user to choose which row to keep
* Removes redundant rows based on user decisions

### 3\. Column B Duplicate Processing \(`process_column_b_duplicates`)

* Identifies duplicate values in the second specified column
* For each duplicate group:
    * Displays all rows with the duplicate value
    * Asks user if they want to modify one of the values
    * Allows user to enter a new value to resolve the conflict
* Continues until no duplicates remain or user chooses to skip

## Example Workflow

```
Excel Sheet Cleaner Utility
==================================================
Enter the path to your Excel file: data.xlsx
Automatically selected sheet: 'Sheet1'
Enter the name of first column of your Excel file: ID
Enter the name of second column of your Excel file: Name

--- Cleaning sheet ---
  - Found and removed 2 completely empty row(s).

--- Processing Duplicates in ID ---
Found 1 group(s) of duplicates in ID.

Processing duplicates for value: '123' in ID
  - CONFLICT: Values in Name are different. Please choose which row to KEEP.
  Enter the index of the row to KEEP [5, 8]: 5
  - Keeping index 5 and marking others for deletion.

--- Processing Duplicates in Name ---
Found duplicate value 'John Doe' in Name at the following indices:
Do you want to provide a new value for one of these? (yes/no): yes
  Enter the index of the row to CHANGE [3, 7]: 7
  Enter the new value for Column B at index 7: John Smith
  - Updated index 7 with new value 'John Smith'.

✅ Success! The cleaned data has been saved to 'cleaned_data.xlsx'
```

## Error Handling

* **File Not Found**: Validates file path before processing
* **Invalid Sheet Names**: Checks if specified worksheet exists
* **Missing Columns**: Verifies that specified columns exist in the data
* **Invalid User Input**: Validates user choices and prompts for correction
* **File Save Errors**: Reports issues when saving the cleaned file

## Output

The utility generates a new Excel file with:

* Filename format: `cleaned_[original_filename].xlsx`
* Same structure as original file but with duplicates resolved
* Reset row indices for clean presentation
* All original data preserved except for removed duplicates and empty rows

## Limitations

* Processes only two columns at a time
* Requires user interaction for conflict resolution
* Column B duplicate processing stops after user chooses to skip a group
* Works with Excel files only (.xlsx, .xls formats)