import pandas as pd
from openpyxl import load_workbook

# === STEP 1: Load Excel data using pandas ===

# Define the path to your existing Excel file
input_file = "value_messages_excel.xlsx"

# Read the 'Input' sheet into a DataFrame
# We use `index_col=0` because the first column contains row labels like "Row1", "Row2", etc.
df = pd.read_excel(input_file, sheet_name="Input", index_col=0)

# Print the data to confirm it's loaded correctly
print("Loaded Table:")
print(df)

# === STEP 2: Load the workbook and target sheet using openpyxl ===

# Load the full Excel workbook (preserves formatting and multiple sheets)
wb = load_workbook(input_file)

# Check if the 'Output' sheet already exists
if "Output" not in wb.sheetnames:
    # If not, raise an error so we don't accidentally write somewhere else
    raise ValueError("Sheet 'Output' does not exist in the file.")

# Select the existing 'Output' sheet (don't create or delete anything)
ws_output = wb["Output"]

# === STEP 3: Write messages into the Output sheet ===

# Start writing messages from the first row
row_index = 1

print("\nWriting messages to existing 'Output' sheet...")

# Loop through every row and column of the DataFrame
for row_name in df.index:
    for col_name in df.columns:
        # Get the value at the intersection of current row and column
        value = df.loc[row_name, col_name]

        # Create the message using f-string formatting
        message = f"The value from row: {row_name}, and column: {col_name}, is: {value}"

        # Write the message into column A of the 'Output' sheet
        ws_output.cell(row=row_index, column=1, value=message)

        # Print to terminal (optional debugging)
        print(message)

        # Move to the next row in the sheet
        row_index += 1

# === STEP 4: Save changes back to the same file ===

# Save the modified workbook, overwriting the original file
wb.save(input_file)

print(f"\n✅ Messages written to existing sheet and saved in: '{input_file}'")
