import pandas as pd
from openpyxl import load_workbook

# Function to check if data already exists in the previous columns
def check_previous_data(sheet, row, col, data):
    if not sheet.empty and col < sheet.max_column:
        for i in range(len(data)):
            if sheet.cell(row=row + i, column=col).value is not None:
                return True
    return False

# Load existing Excel file (if it exists)
try:
    existing_data = pd.read_excel(r'C:\Users\EN6509\OneDrive - CIGNA\Desktop\Organon Monthly Utilization Reporting_11.21.23.xlsx', sheet_name='Sheet1', header=None)
except FileNotFoundError:
    existing_data = pd.DataFrame()

# Creating the first dataframe
data1 = {'Column1': [1, 2, 3, 4, 5, 6],
         'Column2': ['A', 'B', 'C', 'D', 'E', 'F']}
df1 = pd.DataFrame(data1)

# Creating the second dataframe
data2 = {'Column1': [10, 20, 30, 40, 50, 60],
         'Column2': ['X', 'Y', 'Z', 'W', 'U', 'V']}
df2 = pd.DataFrame(data2)

# Creating the third dataframe
data3 = {'Column1': [100, 200, 300, 400, 500, 600],
         'Column2': ['AA', 'BB', 'CC', 'DD', 'EE', 'FF']}
df3 = pd.DataFrame(data3)

# Creating the fourth dataframe
data4 = {'Column1': [1000, 2000, 3000, 4000, 5000, 6000],
         'Column2': ['XX', 'YY', 'ZZ', 'WW', 'UU', 'VV']}
df4 = pd.DataFrame(data4)

# Transpose the dataframes
df1_transposed = df1.transpose()
df2_transposed = df2.transpose()
df3_transposed = df3.transpose()
df4_transposed = df4.transpose()

# Specify start positions for sheet 1
start_row_df1 = 127
start_col_df1 = 1

start_row_df2 = 18
start_col_df2 = 9  # Column I (9th column in 1-indexed notation)

# Specify start positions for sheet 2
start_row_df3 = 127
start_col_df3 = 1

start_row_df4 = 18
start_col_df4 = 9  # Column I (9th column in 1-indexed notation)

# Use openpyxl to load existing workbook
book = load_workbook(r'C:\Users\EN6509\OneDrive - CIGNA\Desktop\Organon Monthly Utilization Reporting_11.21.23.xlsx')

# Get or create Sheet1
try:
    sheet1 = book['Sheet1']
except KeyError:
    sheet1 = book.create_sheet('Sheet1')

# Get or create Sheet2
try:
    sheet2 = book['Sheet2']
except KeyError:
    sheet2 = book.create_sheet('Sheet2')

# Function to find the next available column in the sheet
def find_next_column(sheet):
    return sheet.max_column + 1 if sheet.max_column > 0 else 1

# Function to write data to the sheet
def write_to_sheet(sheet, start_row, start_col, data):
    for r_idx, row in enumerate(data.itertuples(index=False), start=start_row):
        for c_idx, value in enumerate(row, start=start_col):
            sheet.cell(row=r_idx, column=c_idx, value=value)

# Write data to Sheet1
current_col_df1 = find_next_column(sheet1)
if not check_previous_data(existing_data, start_row_df1, current_col_df1, df1_transposed.values):
    write_to_sheet(sheet1, start_row_df1, current_col_df1, df1_transposed)

current_col_df2 = find_next_column(sheet1)
if not check_previous_data(existing_data, start_row_df2, current_col_df2, df2_transposed.values):
    write_to_sheet(sheet1, start_row_df2, current_col_df2, df2_transposed)

# Write data to Sheet2
current_col_df3 = find_next_column(sheet2)
if not check_previous_data(existing_data, start_row_df3, current_col_df3, df3_transposed.values):
    write_to_sheet(sheet2, start_row_df3, current_col_df3, df3_transposed)

current_col_df4 = find_next_column(sheet2)
if not check_previous_data(existing_data, start_row_df4, current_col_df4, df4_transposed.values):
    write_to_sheet(sheet2, start_row_df4, current_col_df4, df4_transposed)

# Save the modified workbook
book.save(r'C:\Users\EN6509\OneDrive - CIGNA\Desktop\Organon Monthly Utilization Reporting_11.21.23.xlsx')
