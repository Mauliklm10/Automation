import pandas as pd
from openpyxl import load_workbook

# Function to find the next available row and column in the sheet
def find_next_position(sheet, start_row, start_col):
    current_row = start_row
    current_col = start_col

    while sheet.cell(row=current_row, column=current_col).value is not None:
        current_row += 1

    return current_row, current_col

# Load existing Excel file (if it exists)
try:
    existing_book = load_workbook('your_existing_file.xlsx')
except FileNotFoundError:
    existing_book = None

# DataFrame 1
data1 = {'Name': ['Alice', 'Bob', 'Charlie', 'David'],
         'Age': [25, 30, 22, 35]}
df1 = pd.DataFrame(data1)

# DataFrame 2
data2 = {'City': ['New York', 'San Francisco', 'Los Angeles', 'Chicago'],
         'Population': [8500000, 884000, 3990000, 2716000]}
df2 = pd.DataFrame(data2)

# DataFrame 3
data3 = {'Subject': ['Math', 'English', 'Science', 'History'],
         'Grade': ['A', 'B', 'A', 'C']}
df3 = pd.DataFrame(data3)

# DataFrame 4
data4 = {'Product': ['Laptop', 'Smartphone', 'Tablet', 'Headphones'],
         'Price': [1200, 800, 400, 150]}
df4 = pd.DataFrame(data4)

# Specify the starting positions for each DataFrame
start_row_df1, start_col_df1 = 127, 1  # Sheet 1, starting from row 127, column A
start_row_df2, start_col_df2 = 16, 7   # Sheet 1, starting from row 16, column G
start_row_df3, start_col_df3 = 127, 1  # Sheet 2, starting from row 127, column A
start_row_df4, start_col_df4 = 16, 7   # Sheet 2, starting from row 16, column G

# Open existing Excel file or create a new one
with pd.ExcelWriter('your_existing_file.xlsx', engine='openpyxl', mode='a') as writer:
    if existing_book:
        writer.book = existing_book
    else:
        writer.book.create_sheet('Sheet1')
        writer.book.create_sheet('Sheet2')

    # Get or create the specified sheets
    sheet1 = writer.book['Sheet1']
    sheet2 = writer.book['Sheet2']

    # Find next available position in Sheet 1 for DataFrame 1
    next_row_df1, next_col_df1 = find_next_position(sheet1, start_row_df1, start_col_df1)
    df1.to_excel(writer, sheet_name='Sheet1', startrow=next_row_df1 - 1, startcol=next_col_df1 - 1, index=False, header=False)

    # Find next available position in Sheet 1 for DataFrame 2
    next_row_df2, next_col_df2 = find_next_position(sheet1, start_row_df2, start_col_df2)
    df2.to_excel(writer, sheet_name='Sheet1', startrow=next_row_df2 - 1, startcol=next_col_df2 - 1, index=False, header=False)

    # Find next available position in Sheet 2 for DataFrame 3
    next_row_df3, next_col_df3 = find_next_position(sheet2, start_row_df3, start_col_df3)
    df3.to_excel(writer, sheet_name='Sheet2', startrow=next_row_df3 - 1, startcol=next_col_df3 - 1, index=False, header=False)

    # Find next available position in Sheet 2 for DataFrame 4
    next_row_df4, next_col_df4 = find_next_position(sheet2, start_row_df4, start_col_df4)
    df4.to_excel(writer, sheet_name='Sheet2', startrow=next_row_df4 - 1, startcol=next_col_df4 - 1, index=False, header=False)
