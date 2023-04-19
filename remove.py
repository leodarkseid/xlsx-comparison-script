import pandas as pd

# Read the Excel sheet
df = pd.read_excel('sheet_b.xlsx')

# Convert all values to strings
df = df.astype(str)

# Iterate over all cells in the sheet
for i in range(df.shape[0]):
    for j in range(df.shape[1]):

        # Get the cell value as a string
        cell_value = df.iloc[i, j]

        # Remove all non-numeric characters
        numeric_value = ''.join([char for char in cell_value if char.isnumeric()])

        # If the cell value is not numeric, replace it with NaN
        if not numeric_value:
            df.iloc[i, j] = pd.NA
        else:
            df.iloc[i, j] = numeric_value

# Convert all values back to numeric types
df = df.apply(pd.to_numeric, errors='coerce')

# Write the cleaned data to a new Excel file
df.to_excel('cleaned_sheet.xlsx', index=False)

# Print a message to confirm that the script has finished running
print('Script completed successfully.')