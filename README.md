import pandas as pd

# Creating the first dataframe
data1 = {'Column1': [1, 2, 3, 4, 5, 6],
         'Column2': ['A', 'B', 'C', 'D', 'E', 'F']}
df1 = pd.DataFrame(data1)

# Creating the second dataframe
data2 = {'Column1': [10, 20, 30, 40, 50, 60],
         'Column2': ['X', 'Y', 'Z', 'W', 'U', 'V']}
df2 = pd.DataFrame(data2)

# Save the initial data to the Excel file
with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
    df1.to_excel(writer, sheet_name='Sheet1', startrow=2, index=False)
    df2.to_excel(writer, sheet_name='Sheet1', startrow=11, index=False)

# Assume this is done on a monthly basis
# New data for the first dataframe (for the next month)
new_data1 = {'Column1': [7, 8, 9],
             'Column2': ['G', 'H', 'I']}
new_df1 = pd.DataFrame(new_data1)

# New data for the second dataframe (for the next month)
new_data2 = {'Column1': [70, 80, 90],
             'Column2': ['Y', 'Z', 'W']}
new_df2 = pd.DataFrame(new_data2)

# Read the existing data from the Excel file
existing_df1 = pd.read_excel('output.xlsx', sheet_name='Sheet1', skiprows=2)
existing_df2 = pd.read_excel('output.xlsx', sheet_name='Sheet1', skiprows=11)

# Append the new data to the existing data
updated_df1 = pd.concat([existing_df1, new_df1], ignore_index=True)
updated_df2 = pd.concat([existing_df2, new_df2], ignore_index=True)

# Write the updated data back to the Excel file
with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
    updated_df1.to_excel(writer, sheet_name='Sheet1', startrow=2, index=False)
    updated_df2.to_excel(writer, sheet_name='Sheet1', startrow=11, index=False)# Automation



import pandas as pd

# Creating the first dataframe
data1 = {'Column1': [1, 2, 3, 4, 5, 6],
         'Column2': ['A', 'B', 'C', 'D', 'E', 'F']}
df1 = pd.DataFrame(data1)

# Creating the second dataframe
data2 = {'Column1': [10, 20, 30, 40, 50, 60],
         'Column2': ['X', 'Y', 'Z', 'W', 'U', 'V']}
df2 = pd.DataFrame(data2)

# Check if the Excel file already exists
try:
    existing_data = pd.read_excel('output.xlsx', sheet_name='Sheet1')
    existing_columns = existing_data.columns.tolist()
except FileNotFoundError:
    existing_columns = []

# Exporting to Excel with specified starting cells
with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
    # Write existing data
    if 'Sheet1' in writer.sheets:
        existing_data.to_excel(writer, sheet_name='Sheet1', index=False)

    # Append or write the first dataframe starting from cell A3
    if 'Column1' not in existing_columns:
        df1.to_excel(writer, sheet_name='Sheet1', startrow=2, index=False)
    else:
        print("Data already present in Column1")

    # Append or write the second dataframe starting from cell A12
    if 'Column1' not in existing_columns:
        df2.to_excel(writer, sheet_name='Sheet1', startrow=11, index=False)
    else:
        print("Data already present in Column1")



import pandas as pd

# Function to check if data already exists in the previous columns
def check_previous_data(sheet, row, col, data):
    for i in range(len(data)):
        if sheet.iloc[row + i, col] is not None:
            return True
    return False

# Load existing Excel file (if it exists)
try:
    existing_data = pd.read_excel('output.xlsx', sheet_name='Sheet1', header=None)
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

# Transpose the dataframes
df1_transposed = df1.transpose()
df2_transposed = df2.transpose()

# Append data to the existing data
start_row = 2
start_col = 0

with pd.ExcelWriter('output.xlsx', engine='openpyxl', mode='a') as writer:
    if not existing_data.empty:
        start_col = len(existing_data.columns)

    if check_previous_data(existing_data, start_row, start_col, df1_transposed.values):
        print("Data already present in previous columns for DataFrame 1")
    else:
        df1_transposed.to_excel(writer, sheet_name='Sheet1', startrow=start_row, startcol=start_col, header=None, index=False)

    if check_previous_data(existing_data, start_row + 9, start_col, df2_transposed.values):
        print("Data already present in previous columns for DataFrame 2")
    else:
        df2_transposed.to_excel(writer, sheet_name='Sheet1', startrow=start_row + 9, startcol=start_col, header=None, index=False)
