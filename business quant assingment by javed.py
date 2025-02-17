import pandas as pd
import os

# Function to extract data from Excel and save it in CSV
def extract_tables_from_xlsx(xlsx_path, label):
    filename = os.path.basename(xlsx_path)  # Get the file name
    data = []  # Initialize list to hold extracted data
    
    # Read the entire Excel file
    df = pd.read_excel(xlsx_path, sheet_name=None)  # Read all sheets
    
    # Iterate over sheets and their data
    for sheet_name, sheet_data in df.items():
        # Try to find a table title (for example, the first non-empty cell in the first row)
        # If the first cell is empty, we can use the sheet name as the table title.
        tabletitle = sheet_data.iloc[0, 0] if pd.notna(sheet_data.iloc[0, 0]) else sheet_name
        
        # Iterate through each row in the sheet (excluding header row if needed)
        for row in sheet_data.itertuples(index=False):
            # For each row, iterate through columns (headers)
            for header, value in zip(sheet_data.columns, row):
                data.append([filename, label, tabletitle, header, value])
    
    return data

# Main part of the script
if __name__ == "__main__":
    xlsx_path = r"C:\Users\admin\Downloads\book.xlsx"  # Specify the path to your Excel file
    label = "Example_Label"  # Change the label as needed
    
    # Extract the data from the Excel file
    extracted_data = extract_tables_from_xlsx(xlsx_path, label)
    
    # Convert the extracted data into a DataFrame
    df = pd.DataFrame(extracted_data, columns=["filename", "label", "tabletitle", "header", "value"])
    
    # Save the DataFrame to a CSV file
    df.to_csv("output.csv", index=False)
    
    print("Extraction complete! Check output.csv")
