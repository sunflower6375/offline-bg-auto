import pandas as pd
import numpy as np
import os

df = pd.read_excel(r"C:\Users\linh.mynguyen\OneDrive - Seagroup\Linh's folder\0. Offline MKT\May 2024 Offline Budget\Shopee May'24 Budget Template (Marketing Projects).xlsx", header=25, sheet_name = "Shopee Marketing Projects")

new_col = ['Company','Product Code','Notelify Instantly','Add Reviewers','Members View','Attachments','Period','Category','Budget', "Remarks"]
sort_col = ['Owner','Project Name','Company','Region','Product Code','Project Type','Start Date','End Date','Description',
            'Notelify Instantly','Add Reviewers','Members View','Members','Attachments','Period','Category','Budget','Remarks']

for col in new_col:
    if col == "Company":
        df[col] = "SPV"
    elif col == "Product Code":
        df[col] = "EC_SPE_COM"
    elif col == "Notify Instantly":
        df[col] = "N"
    elif col == "Members View":
        df[col] = "N"
    elif col == "Period":
        df[col] = "May 2024"
    elif col == "Category":
        df[col] = "Others"
    elif col == "Remarks":
        df[col] = "Basing on quotation"    
    else:
        df[col] = np.nan

# df['Period'] = df['Period'].astype(str)

df["Start Date"] = df["Start Date"].dt.strftime('%d-%b-%Y')
df["End Date"] = df["End Date"].dt.strftime('%d-%b-%Y')

df = df[sort_col]

old_path = r"C:\Users\linh.mynguyen\OneDrive - Seagroup\Linh's folder\0. Offline MKT\May 2024 Offline Budget\May 2024 Offline Budget.xlsx"
new_path = r"C:\Users\linh.mynguyen\OneDrive - Seagroup\Linh's folder\0. Offline MKT\May 2024 Offline Budget\May 2024 Offline Budget.csv"

df.to_excel(old_path, index=False, encoding='utf-8')

os.rename(old_path, new_path)