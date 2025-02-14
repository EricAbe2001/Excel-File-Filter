# Excel-File-Filter
This Python script compares two Excel files and generates a new Excel file that highlights the differences. The output file contains only the rows where changes were detected, making it easy to see what has been modified.

## Description

This script uses the `pandas` and `openpyxl` libraries to efficiently compare two Excel files. It identifies cell-level differences and highlights them in yellow in the output file.  The script is designed to handle Excel files with potentially different dimensions (number of rows and columns). It also tracks metadata about the comparisons (last modified time, count) in a separate file.

## Features

* **Highlights Differences:**  Cells with different values are highlighted in yellow.
* **Outputs Changed Rows Only:** The output file only contains rows with at least one change.
* **Handles Different Dimensions:** Works with Excel files having varying row and column counts.
* **Empty File Handling:** Gracefully handles cases where one or both input files are empty.
* **Metadata Tracking:**  Keeps track of the last modification time and comparison count.

## Requirements

* Python 3.x
* pandas (`pip install pandas`)
* openpyxl (`pip install openpyxl`)

## How to Use

1. **Prepare your Excel files:** Place your two Excel files (e.g., `file_a.xlsx` and `file_b.xlsx`) in the same directory as the Python script or provide their full paths in the script.

2. **Run the script:** Execute the Python script.  The output file (named `comparison_output.xlsx` by default) will be created in the same directory (or the path you specify).

   ```bash
   python Excel_File_Comparsion.py
