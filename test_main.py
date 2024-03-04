import pandas as pd
from openpyxl import load_workbook

# Data file paths (replace with actual names)
sales_data_path = "sales.xlsx"
#team_data_path = "team.xlsx"

# Read data from both Excel files, specifying the header rows
sales_data = pd.read_excel(sales_data_path, header=0)  # Read from the first row (header)
sales_data = sales_data.rename(columns={"Unnamed: 0": "New Column Name", "Unnamed: 1": "Testing"})  # Replace "New Column Name" with your desired name
#team_data = pd.read_excel(team_data_path, header=0)  # Read from the first row (header)
writer = pd.ExcelWriter("sales_data_modified.xlsx")  # Adjust the file path and name if needed
sales_data.to_excel(writer, sheet_name="Sheet1", index=False)  # Adjust sheet name if necessary
writer._save()

print("File saved successfully!")
#print(sales_data.columns.tolist())



sales_data_path = "sales_data_modified.xlsx"

# Read data from the Excel file, specifying the header rows
sales_data = pd.read_excel(sales_data_path, header=0)

# Rename columns (replace with your desired names)
sales_data = sales_data.rename(columns={"Unnamed: 2": "Column Name", "Unnamed: 3": "Testing2"})

# Save changes to the original file (overwrite existing data)
writer = pd.ExcelWriter(sales_data_path, engine='openpyxl')  # Use 'openpyxl' engine for compatibility
sales_data.to_excel(writer, sheet_name="Sheet1", index=False)  # Adjust sheet name if necessary
writer._save()

print("File modified successfully!")
