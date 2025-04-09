import pandas as pd
import os

def extract_rows(input_file, output_file, row_indices, header_rows):
    # Read the Excel file
    df = pd.read_excel(input_file, header=None)
    
    # Extract header rows
    headers = df.iloc[:header_rows, :].values.tolist()
    
    # Extract specified rows
    sampled_rows = df.iloc[row_indices, :].values.tolist()
    
    # Combine header rows and content rows
    combined_rows = headers + sampled_rows
    
    # Create a new DataFrame
    new_df = pd.DataFrame(combined_rows)
    
    # Write to a new Excel file
    new_df.to_excel(output_file, index=False, header=False)

def main():
    try:
        # Get user input
        input_file = input("Please enter the source Excel file path: ")
        
        # Get the directory of the input file
        input_directory = os.path.dirname(input_file)
        output_file = os.path.join(input_directory, "Sampling Excel.xlsx")
        
        row_indices = list(map(int, input("Please enter the indices to sample (comma-separated): ").split(',')))
        header_rows = int(input("Please enter the number of header rows: "))
        
        # Extract rows and save to a new Excel file
        extract_rows(input_file, output_file, row_indices, header_rows)
        
        print(f"Sampling result has been saved to {output_file}")
        
    except Exception as e:
        print("An error occurred: ", e)

if __name__ == "__main__":
    main()



