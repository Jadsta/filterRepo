import os
import pandas as pd
import re
from openpyxl import load_workbook
from datetime import datetime

# Path to the folder containing Excel files
input_path = "./toreplace/"

# Output folder for modified files
output_folder = "./cleansed/"

# Read the replacement data from "replace.xlsx"
replace_df = pd.read_excel("replace.xlsx")

# Create a log file to record replacements
log_file = open("replacements_log.txt", "w")

# Iterate through all Excel files in the input folder
for filename in os.listdir(input_path):
    if filename.endswith(".xlsx"):
        file_path = os.path.join(input_path, filename)
        
        # Print filename and timestamp when starting processing
        print(f"Processing file: {filename} ({datetime.now()})")
        
        # Load workbook and all sheet names
        workbook = load_workbook(file_path)
        sheet_names = workbook.sheetnames
        
        # Perform find-and-replace within sentences for each sheet
        for sheet_name in sheet_names:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows(min_row=1, max_col=sheet.max_column, max_row=sheet.max_row, values_only=False):
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        for index, row in replace_df.iterrows():
                            find_word = row.iloc[0]  # First column
                            replace_word = row.iloc[1]  # Second column
                            cell.value = re.sub(r'\b' + find_word + r'\b', replace_word, cell.value, flags=re.IGNORECASE)

        # Save the modified workbook to a new file in the output folder
        output_file_path = os.path.join(output_folder, f"modified_{filename}")
        workbook.save(output_file_path)

        # Log the replacements (this part will need to be adjusted as we're not using pandas here)
        # log_file.write(f"File: {filename}, Replacements: {replacement_count}\n")

        # Print completion message
        print(f"Processing completed for file: {filename} ({datetime.now()})")

log_file.close()
print("Find-and-replace completed for all Excel files. Modified files saved with 'modified_' prefix in the specified output folder.")