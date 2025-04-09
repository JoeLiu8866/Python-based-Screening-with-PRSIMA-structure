import pandas as pd

# Manually enter the path of the Excel file
file_path = input("Please enter the path of the Excel file: ")

# Load the Excel file
df = pd.read_excel(file_path)

# Define keywords related to automated driving
keywords = ['automated driving', 'autonomous', 'self driving', 'driverless', 'robotic', 'automated']

# Convert keywords to lowercase
keywords = [keyword.lower() for keyword in keywords]

# Initial row count
initial_row_count = df.shape[0]

# Filter rows where 'Title' or 'Abstract' contains any keywords
relevant_articles = df[
    df['Title'].str.contains('|'.join(keywords), case=False, na=False) |
    df['Abstract'].str.contains('|'.join(keywords), case=False, na=False)
]

# Exclude irrelevant articles
excluded_articles = df[~df.index.isin(relevant_articles.index)]

# Final row count
final_row_count = relevant_articles.shape[0]

# Calculate the number of rows removed
rows_removed = initial_row_count - final_row_count

# Save the cleaned DataFrame to a new Excel file
output_relevant_path = file_path.replace('.xlsx', '_S2_relevant.xlsx')
output_excluded_path = file_path.replace('.xlsx', '_S2_excluded.xlsx')

relevant_articles.to_excel(output_relevant_path, index=False)
excluded_articles.to_excel(output_excluded_path, index=False)

# Output the number of rows removed and the paths of the saved files
print(f"Total rows removed: {rows_removed}")
print(f"Relevant articles saved to: {output_relevant_path}")
print(f"Excluded articles saved to: {output_excluded_path}")

