import pandas as pd
from openpyxl import load_workbook
from time import sleep

# Data file paths (replace with actual names of files)
try:
    deal_file = input("Provide name of deal file: ")
    team_file = input("Provide name of your team file: ")

    sales_data_path = f"file/{deal_file}.xlsx"
    team_data_path = f"file/{team_file}.xlsx"

    # Read data from both Excel files, specifying the header rows
    sales_data = pd.read_excel(sales_data_path, header=1, dtype={'Opp ID': str})
    team_data = pd.read_excel(team_data_path, header=1, dtype={'Opportunity ID': str})


    # Ensure unique index for merging
    sales_data = sales_data.reset_index(drop=True)  # Create a new, unique integer index

    # Merge data based on the correct joining column
    merged_data = sales_data.merge(team_data, left_on="Opp ID", right_on="Opportunity ID", how="left")

    # Add "Supported" column, adjusting for potential column name differences
    supported_column_name = "Opp ID" if "Opp ID" in merged_data.columns else "Opportunity ID"
    merged_data["CST Support"] = merged_data["Opp ID"].isin(team_data["Opportunity ID"]).map({True: "Supported", False: "Not supported"})

    # Create the final report without setting an index
    writer = pd.ExcelWriter("weekly_report.xlsx")
    merged_data.to_excel(writer, sheet_name="Report", index=False)
    writer._save()  # Prefer using 'save' for clarity i.e writer.save()

    print("Report generation complete!")

    sleep(2)

    # Renaming columns

    print("Renaming columns...")

    sleep(2)


    new_file = "weekly_report.xlsx"

    read_file = pd.read_excel(new_file, header=0)

    read_file = read_file.rename(columns={
        "Unnamed: 21": "ADRM + Upside",
        "Unnamed: 22": "ISBN (L2)",
        "Unnamed: 23": "BTP (L2)",
        "Unnamed: 24": "CX (L2)",
        "Unnamed: 25": "ERP (L2)",
        "Unnamed: 26": "HXM (L2)",
        "Unnamed: 28": "SAP Signavio (L2)",
        "Unnamed: 29": "S4 Public cloud (L3)",
        "Unnamed: 30": "S4 Private cloud (L3)",
        })

    writer = pd.ExcelWriter(new_file, engine="openpyxl")
    read_file.to_excel(writer, sheet_name="sheet2", index=False)
    writer._save()

    print("Modifications complete")

    sleep(2)

    print("Operation completed")

except FileNotFoundError:

    print("File not Found, check file folder to see name of file")