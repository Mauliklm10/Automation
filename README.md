with pd.ExcelWriter(r'C:\Users\EN6509\OneDrive - CIGNA\Desktop\Organon Monthly Utilization Reporting_11.21.23.xlsx', engine='openpyxl', mode='a') as writer:
    if not check_previous_data(existing_data, start_row_df1, start_col_df1, df1_transposed.values):
        df1_transposed.to_excel(writer, sheet_name='Sheet1', startrow=start_row_df1, startcol=start_col_df1, header=None, index=False)

    if not check_previous_data(existing_data, start_row_df2, start_col_df2, df2_transposed.values):
        df2_transposed.to_excel(writer, sheet_name='Sheet1', startrow=start_row_df2, startcol=start_col_df2, header=None, index=False)

# Write data to Sheet2
with pd.ExcelWriter(r'C:\Users\EN6509\OneDrive - CIGNA\Desktop\Organon Monthly Utilization Reporting_11.21.23.xlsx', engine='openpyxl', mode='a') as writer:
    if not check_previous_data(existing_data, start_row_df3, start_col_df3, df3_transposed.values):
        df3_transposed.to_excel(writer, sheet_name='Sheet2', startrow=start_row_df3, startcol=start_col_df3, header=None, index=False)

    if not check_previous_data(existing_data, start_row_df4, start_col_df4, df4_transposed.values):
        df4_transposed.to_excel(writer, sheet_name='Sheet2', startrow=start_row_df4, startcol=start_col_df4, header=None, index=False)
