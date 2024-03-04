import pandas as pd
from openpyxl import load_workbook

# Data file paths (replace with actual names)
name_of_file = input("Provide name of file to get header: ")



file_path = f"file/{name_of_file}.xlsx"


# Read data from both Excel files, specifying the header rows
file_data = pd.read_excel(file_path, header=1)  # Read from the first row (header)






print(file_data.columns.tolist())


