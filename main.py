import pandas as pd

# Data file paths (replace with actual names of files)
sales_data_path = "file/sales.xlsx"
team_data_path = "file/team.xlsx"

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

# Renaming columns



