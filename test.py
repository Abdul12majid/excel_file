import pandas as pd
from openpyxl import load_workbook
from time import sleep


name_of_file = input("Provide name of file to get header: ")


file_path = f"file/{name_of_file}.xlsx"

read_file = pd.read_excel(file_path, header=0)

x = read_file.columns.tolist()
y = [1, 2, 3]
print(len(x))