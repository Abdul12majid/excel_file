import pandas as pd
from openpyxl import load_workbook
import time

# Data file paths (replace with actual names)
try:

	name_of_file = input("Provide name of file to get header: ")


	file_path = f"file/{name_of_file}.xlsx"


	# Read data from both Excel files, specifying the header rows
	file_data = pd.read_excel(file_path, header=0, dtype={'Opp ID': str})  # Read from the first row (header)


	x = file_data.columns.tolist()

	if 'Unnamed: 0' in x and 'Unnamed: 1' in x:
		print('Invalid column name, checking next header')

		time.sleep(1)

		file_data = pd.read_excel(file_path, header=1, dtype={'Opp ID': str})  # Read from the first row (header)

		print(file_data.columns.tolist())

except FileNotFoundError:
	time.sleep(2)

	print("Name not found, Kindly check your file folder")
