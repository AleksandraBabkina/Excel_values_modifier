# Excel_values_modifier
## Description
This program modifies numeric values in an Excel file by applying random changes within a specified range. The script targets columns with names starting with "Rp" and rows containing certain values ("X", "Y", "H 0") in column C. It then adjusts the numeric values in the selected cells, ensuring they are slightly changed by either increasing or decreasing within a defined range.

## Functional Description
The program performs the following steps:
1. Opens the Excel file specified by the user.
2. Identifies columns whose names start with "Rp".
3. Searches for rows that contain values like "X", "Y", or "H 0" in column C.
4. For the selected rows and columns, it generates random values within a defined range and modifies the original numeric values.
5. Saves the modified data into a new Excel file.

## How It Works
1. The script loads the specified Excel file using the `openpyxl` library.
2. It scans the 6th row for column names starting with "Rp" and stores the corresponding columns.
3. The program then scans the 3rd column (C) for the presence of specific values ("X", "Y", "H 0").
4. Once the relevant columns and rows are identified, the script adjusts the numeric values in the matching cells by adding or subtracting a random value.
5. After making the modifications, the script saves the changes in a new file, which is named by appending "Измененный_" to the original filename.

## Input Structure
To run the program, the following parameters need to be provided:
1. Excel file name: A valid Excel file (.xlsx) that contains the data to be modified.
2. The script will target columns whose names start with "Rp" and rows with values "X", "Y", or "H 0" in column C.

## Technical Requirements
To run the program, the following are required:
1. Python 3.x
2. Installed libraries:
   - openpyxl
   - random
   - os
   - time

## Usage
1. Place the script in the directory with the Excel file to be modified.
2. Ensure that the file name is correctly specified in the `file_name` variable.
3. Run the script. It will:
   - Find columns starting with "Rp".
   - Find rows with specific values in column C ("X", "Y", or "H 0").
   - Modify numeric values in the corresponding cells.
   - Save the modified data in a new file with the prefix "Измененный_".

## Example Output
Output:
- A new Excel file will be saved with the name `Журналы.xlsx`.

## Conclusion
This script provides an automated way to modify specific numeric values in an Excel file based on certain criteria, making it useful for data manipulation tasks in various domains.
