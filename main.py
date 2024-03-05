import pandas as pd
from openpyxl import load_workbook
from time import sleep
from datetime import datetime

# Data file paths (replace with actual names of files)
try:
    deal_file = input("Provide name of deal file: ")
    team_file = input("Provide name of your team file: ")

    sales_header = int(input("ADRM header number: "))
    team_header = int(input("CST header number: "))

    sales_data_path = f"file/{deal_file}.xlsx"
    team_data_path = f"file/{team_file}.xlsx"

    # Read data from both Excel files, specifying the header rows
    try:
        sales_data = pd.read_excel(sales_data_path, header=sales_header, dtype={'Opp ID': str})
        team_data = pd.read_excel(team_data_path, header=team_header, dtype={'Opportunity ID': str})


        #print(sales_data.columns.tolist())


    

        # Ensure unique index for merging
        sales_data = sales_data.reset_index(drop=True)  # Create a new, unique integer index

        # Merge data based on the correct joining column
        merged_data = sales_data.merge(team_data, left_on="Opp ID", right_on="Opportunity ID", how="left")


        today = datetime.today().strftime("%d%m%y")
        filename = f"deal execution report {today} macro test.xlsx"

        # Add "Supported" column, adjusting for potential column name differences
        supported_column_name = "Opp ID" if "Opp ID" in merged_data.columns else "Opportunity ID"
        merged_data["CST Support"] = merged_data["Opp ID"].isin(team_data["Opportunity ID"]).map({True: "Supported", False: "Not supported"})

        # Create the final report without setting an index
        writer = pd.ExcelWriter(filename)
        merged_data.to_excel(writer, sheet_name="Cloud Details", index=False)
        writer._save()  # Prefer using 'save' for clarity i.e writer.save()

        print("Report generation complete!")

        sleep(2)
        
    except:
        print('Invalid column header number specified, kindly check both files to be sure.')

    # Renaming columns


except FileNotFoundError:

    print("File not Found, check file folder to see name of file")

