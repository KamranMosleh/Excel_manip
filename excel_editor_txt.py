import pandas as pd
from openpyxl import load_workbook

# Load the Excel file and the Input sheet
input_file = "value_messages_excel.xlsx"
df = pd.read_excel(input_file, sheet_name="Input", index_col=0)

print("Loaded Table:")
print(df)

# Load the workbook and access existing Output sheet
wb = load_workbook(input_file)

# Ensure 'Output' sheet exists
if "Output" not in wb.sheetnames:
    raise ValueError("Sheet 'Output' does not exist in the file.")

ws_output = wb["Output"]

# Write messages to the existing Output sheet
print("\nWriting messages to existing 'Output' sheet...")

row_index = 1
for row_name in df.index:
    for col_name in df.columns:
        value = df.loc[row_name, col_name]
        message = f"The value from row: {row_name}, and column: {col_name}, is: {value}"
        ws_output.cell(row=row_index, column=1, value=message)
        print(message)
        row_index += 1

# Save the changes back to the same file
wb.save(input_file)
print(f"\n✅ Messages written to existing sheet and saved in: '{input_file}'")
