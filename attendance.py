import pandas as pd
import os
from pathlib import Path

folder_path = '2023-12-10'
base_dir = Path('attendance') / folder_path
all_files = [f.name for f in base_dir.iterdir() if f.name.endswith('.xlsx') and f.name.startswith('Attendance')]

all_dataframes = []

for file in all_files:
    file_path = os.path.join(base_dir, file)
    df = pd.read_excel(file_path)

    # Add a new column 'filename' with the modified filename
    modified_filename = file.replace('.xlsx', '').replace('Attendance - ', '')
    df['shop'] = modified_filename

    all_dataframes.append(df)

combined_dataframe = pd.concat(all_dataframes, ignore_index=True)

# Group by 'EmployeeID'
grouped = combined_dataframe.groupby('EmployeeID').agg(
    total_hours_sum=('Total Hours', 'sum'),
    row_count=('EmployeeID', 'size')
).reset_index()

# Rename the columns for clarity
grouped.columns = ['EmployeeID', 'Total Hours Sum', 'Row Count']

# Define the path for the new Excel file
new_file_path = os.path.join(base_dir, folder_path+'.xlsx')

# Using ExcelWriter to write to multiple sheets of the same Excel file
with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
    combined_dataframe.to_excel(writer, sheet_name='All', index=False)
    grouped.to_excel(writer, sheet_name='Grouped', index=False)