import gspread
import pandas as pd
import numpy as np

gc = gspread.service_account(filename='./utilities/api-practice-at-mamon-11-711ff59d549b.json')
sh = gc.open('api-practice-at-mamon11')

dest_sheet = sh.sheet1
dest_sheet.clear()
orig_sheet = pd.read_excel('./sheets/1209_test_report.xlsx')

# remove index
orig_sheet.drop(orig_sheet.columns[0], axis=1, inplace=True)
orig_sheet['Date'] = orig_sheet['Date'].dt.strftime('%Y-%m-%d')
dest_sheet.update([orig_sheet.columns.values.tolist()] + orig_sheet.values.tolist())

print('phase 1 done')

dest_sheet = sh.sheet1
orig_sheet = pd.read_excel('./sheets/1210_test_report.xlsx')

# merge two sheets not replace all of them
orig_sheet.drop(orig_sheet.columns[0], axis=1, inplace=True)
dest_df = pd.DataFrame(dest_sheet.get_all_values())

# remove header index
dest_df.columns = orig_sheet.columns
dest_df.drop(dest_df.index[0], inplace=True)

# merge two sheets
merged_df = pd.concat([dest_df, orig_sheet], ignore_index=True, sort=False)

# remove duplicates
merged_df.drop_duplicates(subset=['Date'], keep='last', inplace=True)

# sort by date
merged_df.sort_values(by=['Date'], inplace=True)

# update data types
merged_df['Viewed displays'] = merged_df['Viewed displays'].astype(int)
merged_df['Clicks'] = merged_df['Clicks'].astype(int)
merged_df['Cost'] = merged_df['Cost'].astype(float)
merged_df['Sales PC30D'] = merged_df['Sales PC30D'].astype(int)
merged_df['Revenue PC30D'] = merged_df['Revenue PC30D'].astype(float)
merged_df['Visits'] = merged_df['Visits'].astype(int)

# update sheet
dest_sheet.update([merged_df.columns.values.tolist()] + merged_df.values.tolist())

print('phase 2 done')


