import pandas as pd
import time

x = 0
y = 10
i = 0

# while x != y:
#     time.sleep(5)  # pause for 5 seconds
#     i += 1
#     print(i)

# Read Excel sheets A and B
print('Reading Excel sheets...')
df_a = pd.read_excel('sheet_a.xlsx')
df_b = pd.read_excel('cleaned_sheet.xlsx')

# Create an empty list to store the matching values
matching_values = []

# Iterate over all cells in sheet A

print('Searching for matching values...')
for cell in df_a.values.flatten():


# Extract the last 7 digits from the cell value
    search_value = str(cell)[-7:]
    # Check if the cell value is in sheet B
    if search_value in df_b.values:
        matching_values.append(cell)
        

# Print out the list of matching values
print('Values found in both sheets:')
for val in sorted(matching_values):
    print(val)

if not matching_values:
    print('No matching values found')
    
