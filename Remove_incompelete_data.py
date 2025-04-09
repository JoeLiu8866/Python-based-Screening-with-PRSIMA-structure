import pandas as pd

# Manually enter the path of the Excel file
file_path = input("Please enter the path of the Excel file: ")

# Load the Excel file
df = pd.read_excel(file_path)

# Extract header and content rows
header = df.iloc[2]
content = df.iloc[3:]

# Rename columns
correct_columns = [
    'Title', 'href', 'year', 'Venue', 'Author', 'Abstract'
]

# correct_columns = [
#     'Title', 'href', 'year', 'Venue', 'Author', 'Abstract', 'Article type', 'Description', 'SAE Level',
#     'Takeover System', 'Takeover Situation', 'Takeover Request', 'Takeover Behavior', 'Trust', 'Task',
#     'Driving Simulator', 'Control Method', 'Driving Context', 'Experiment Time', 'Observation',
#     'Questionnaire', 'Interview', 'Finding', 'Result', 'Good Point', 'Limitation', 'Introduction',
#     'Analysis', 'Valuable Citation', 'Software', 'Data Collection'
# ]

# Adjust column names to match the number of columns in the original DataFrame
adjusted_columns = correct_columns + ['Unnamed'] * (content.shape[1] - len(correct_columns))

# Update DataFrame column names
content.columns = adjusted_columns

# Required columns
required_columns = ['Title', 'year', 'Author', 'Abstract']

# Initial row count
initial_row_count = content.shape[0]

# Drop rows missing required columns and save excluded rows
content_cleaned = content.dropna(subset=required_columns)
excluded_content = content[~content.index.isin(content_cleaned.index)]

# Final row count
final_row_count = content_cleaned.shape[0]

# Calculate the number of rows removed
rows_removed = initial_row_count - final_row_count + 2

# Save the cleaned DataFrame and excluded DataFrame to new Excel files
output_cleaned_path = file_path.replace('.xlsx', ' _S1C.xlsx')
output_excluded_path = file_path.replace('.xlsx', ' _S1E.xlsx')

content_cleaned.to_excel(output_cleaned_path, index=False)
excluded_content.to_excel(output_excluded_path, index=False)

# Output the number of rows removed and the paths of the saved files
print(f"Total rows removed: {rows_removed}")
print(f"Cleaned file saved to: {output_cleaned_path}")
print(f"Excluded file saved to: {output_excluded_path}")

