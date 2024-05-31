import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

# Set the file path
file_path = "C:\\Users\\tiago\\botao\\EmpregadosemExcel.xls"

# Load the workbook
df = pd.read_excel(file_path)

# Remove empty rows
df = df.dropna(how='all')

# Remove empty columns
df = df.loc[:, (df != '').any(axis=0)]

# Create a new workbook
new_wb = openpyxl.Workbook()

# Create a new worksheet in the new workbook
new_ws = new_wb.active

# Copy the modified data to the new worksheet
new_ws.title = df.columns[0]

# Convert DataFrame to rows
rows = dataframe_to_rows(df, index=False, header=True)

# Write rows to the worksheet
for r_idx, row in enumerate(rows, 1):
    for c_idx, value in enumerate(row, 1):
        new_ws.cell(row=r_idx, column=c_idx, value=value)

# Save the new workbook
new_wb.save("C:\\Users\\tiago\\botao\\EmpregadosemExcel_updated.xls")
