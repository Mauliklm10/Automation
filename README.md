import openpyxl
from openpyxl import load_workbook

# Load the existing Excel file
file_path = 'your_excel_file.xlsx'
workbook = load_workbook(file_path)

# Assuming you have four dataframes: df1, df2, df3, and df4

# Get the sheet objects
sheet1 = workbook['Sheet1']
sheet2 = workbook['Sheet2']

# Convert dataframes to openpyxl data
data1 = df1.values.tolist()
data2 = df2.values.tolist()
data3 = df3.values.tolist()
data4 = df4.values.tolist()

# Append data to the sheets
for row_index, row_data in enumerate(data1, start=1):
    for col_index, value in enumerate(row_data, start=1):
        sheet1.cell(row=row_index + 127, column=col_index, value=value)

for row_index, row_data in enumerate(data2, start=1):
    for col_index, value in enumerate(row_data, start=1):
        sheet1.cell(row=row_index, column=col_index + 8, value=value)

for row_index, row_data in enumerate(data3, start=1):
    for col_index, value in enumerate(row_data, start=1):
        sheet2.cell(row=row_index + 127, column=col_index, value=value)

for row_index, row_data in enumerate(data4, start=1):
    for col_index, value in enumerate(row_data, start=1):
        sheet2.cell(row=row_index, column=col_index + 8, value=value)

# Save the changes
workbook.save(file_path)
