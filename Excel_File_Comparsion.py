import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import os
from datetime import datetime

def compare_excel_files(file_a_path, file_b_path, output_file_path):
    try:
        df_a = pd.read_excel(file_a_path, header=None, engine='openpyxl')
        df_b = pd.read_excel(file_b_path, header=None, engine='openpyxl')

        if df_a.empty or df_b.empty:
            print("Warning: One or both Excel files are empty.")
            wb = openpyxl.Workbook()
            ws = wb.active
            wb.save(output_file_path)
            return

        max_rows = max(df_a.shape[0], df_b.shape[0])
        max_cols = max(df_a.shape[1], df_b.shape[1])

        df_a = df_a.reindex(index=range(max_rows), columns=range(max_cols)).fillna("")
        df_b = df_b.reindex(index=range(max_rows), columns=range(max_cols)).fillna("")

        wb = openpyxl.Workbook()
        ws = wb.active

        output_row = 1

        for row in range(max_rows):
            row_has_difference = False  # Flag to track if the current row has any differences
            for col in range(max_cols):
                if row < df_a.shape[0] and col < df_a.shape[1] and row < df_b.shape[0] and col < df_b.shape[1]:
                    cell_a_value = str(df_a.iloc[row, col])
                    cell_b_value = str(df_b.iloc[row, col])

                    if cell_a_value != cell_b_value:
                        row_has_difference = True  # Set the flag if a difference is found

            if row_has_difference: # Write the row only if a difference was found
                for col in range(max_cols): # Iterate through all columns again to write the whole row
                    if row < df_a.shape[0] and col < df_a.shape[1] and row < df_b.shape[0] and col < df_b.shape[1]:
                        cell_a_value = str(df_a.iloc[row, col])
                        cell_b_value = str(df_b.iloc[row, col])
                        ws.cell(row=output_row, column=col + 1, value=cell_b_value)
                        if cell_a_value != cell_b_value:
                            red_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                            ws.cell(row=output_row, column=col + 1).fill = red_fill
                    else:
                        ws.cell(row=output_row, column=col + 1, value="") # Handle out of bounds by writing blank
                output_row += 1

        metadata_file = "comparison_metadata.txt"

        if os.path.exists(metadata_file):
            with open(metadata_file, "r", encoding="utf-8") as f:
                last_modified_str = f.readline().strip()
                try:
                    last_modified = datetime.strptime(last_modified_str, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    last_modified = datetime.now()
                modification_count = int(f.readline().strip())
        else:
            last_modified = datetime.now()
            modification_count = 0

        last_modified = datetime.now()
        modification_count += 1

        with open(metadata_file, "w", encoding="utf-8") as f:
            f.write(last_modified.strftime("%Y-%m-%d %H:%M:%S") + "\n")
            f.write(str(modification_count) + "\n")

        wb.save(output_file_path)
        print(f"Comparison complete. Differences highlighted in {output_file_path}")
        print(f"Last modified: {last_modified.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Modification count: {modification_count}")

    except FileNotFoundError:
        print(f"Error: One or both files not found. Paths: {file_a_path}, {file_b_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

file_a = "Superstore Original Copy.xlsx"
file_b = "Superstore Updated.xlsx"
output_file = "comparison_output.xlsx" 

compare_excel_files(file_a, file_b, output_file)