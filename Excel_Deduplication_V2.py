# -*- coding: gbk -*-
import pandas as pd
import os

def get_excel_files_from_directory(directory):
    excel_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xls'):
                excel_files.append(os.path.join(root, file))
    return excel_files

# Read the root directory containing Excel files from user input
directory = input("Please enter the root directory containing the Excel files: ")
excel_files = get_excel_files_from_directory(directory)

# Display the number of Excel files found
print(f"Found {len(excel_files)} Excel files in the directory")

# Read each Excel file into a DataFrame and store in a list
dataframes = [pd.read_excel(file) for file in excel_files]

# Combine all DataFrames
combined_df = pd.concat(dataframes, ignore_index=True)

# Remove duplicate rows based on all columns
unique_df = combined_df.drop_duplicates()

# Save the deduplicated data to a new Excel file in the root directory
new_file_path = os.path.join(directory, "combined_table.xlsx")
unique_df.to_excel(new_file_path, index=False)

print(f"The combined file has been saved to: {new_file_path}")


