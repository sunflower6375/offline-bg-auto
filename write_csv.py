import pandas as pd
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta

# Calculate month and formatted dates
month = datetime.datetime.now() + relativedelta(months=1)
formatted_date = month.strftime("%b'%y")
formatted_month = month.strftime("%B")
final_month = month.strftime("%b")

# Define sheet name and source path
sheet_nm = "Shopee Marketing Projects"
source = rf"C:\Users\linh.mynguyen\OneDrive - Seagroup\Linh's folder\0. Offline MKT\{formatted_date} Offline Budget\Shopee {formatted_date} Budget Template (Marketing Projects).xlsx"

# Read Excel file
df = pd.read_excel(source, header=25, sheet_name="Shopee Marketing Projects")
df = df.dropna(subset=["Owner"])

# Define new columns and their default values
new_col = ["Company", "Product Code", "Notify Instantly", "Add Reviewers", "Members View", "Attachments", "Period", "Category", "Budget", "Remarks"]
sort_col = ["Owner", "Project Name", "Company", "Region", "Product Code", "Project Type", "Start Date", "End Date", "Description",
            "Notify Instantly", "Add Reviewers", "Members View", "Members", "Attachments", "Period", "Category", "Budget", "Remarks"]

# Add new columns with default values
df["Company"] = "SPV"
df["Product Code"] = "EC_SPE_COM"
df["Notify Instantly"] = "N"
df["Members View"] = "N"
df["Period"] = f"{final_month} {str(month.year)}"
df["Category"] = "Others"
df["Remarks"] = "Basing on quotation"
df["Notify Instantly"] = np.nan
df["Add Reviewers"] = np.nan
df["Attachments"] = np.nan

# Handle "Budget" column based on condition
budget_col_name = f"{formatted_month} Budget (LC)"
if budget_col_name in df.columns:
    df["Budget"] = np.where(df[budget_col_name] == 0, np.nan, df[budget_col_name])
    df["Budget"] = df["Budget"].astype("Int64")
else:
    df["Budget"] = np.nan

# Ensure columns are sorted correctly
df = df[sort_col]

# Define path for CSV output
old_path = rf"C:\Users\linh.mynguyen\OneDrive - Seagroup\Linh's folder\0. Offline MKT\{formatted_date} Offline Budget\{formatted_date} Offline Budget.csv"

# Write DataFrame to CSV
df.to_csv(path_or_buf=old_path, index=False)