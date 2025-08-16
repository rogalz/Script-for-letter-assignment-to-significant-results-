import pandas as pd
from itertools import combinations
import string

input_file = "input.xlsx"
significant_file = "significance.xlsx"

# Load the Excel file and set 'x' column as index
df = pd.read_excel(input_file).set_index("x")

# Calculate row averages and sort descending
row_averages = df.mean(axis=1)
sorted_averages = row_averages.sort_values(ascending=False)
sorted_averages.name = "x"

#*********************************************************************************************************

# Extract the indices in sorted order
indices = list(sorted_averages.index)

# Load the comparison Excel file
comp_df = pd.read_excel(significant_file, header=None)
comp_df.columns = ["pair", "value", "range", "fourth_col", "ns","dupa"]

# Function to get the 4th column given a pair
def get_fourth_col(a, b):
    pair1 = f"{a} vs. {b}"
    pair2 = f"{b} vs. {a}"
    row = comp_df[(comp_df["pair"] == pair1) | (comp_df["pair"] == pair2)]
    if not row.empty:
        return row.iloc[0]["fourth_col"]
    return None

# Generate all combinations of two elements
pairs = list(combinations(indices, 2))

# Extract the 4th column for each pair
results = [(a, b, get_fourth_col(a, b)) for a, b in pairs]

# Convert results to dictionary for fast lookup
pair_dict = {(a, b): val for a, b, val in results}

# Function to check if a pair is not significant
def check_pair(a, b):
    if a == b:
        return '-'  # diagonal
    val = pair_dict.get((a, b)) or pair_dict.get((b, a))
    if val is None:
        return ' '
    return '-' if val == "No" else ' '  # '-' for No, blank for Yes

# Build table with early stop and skip redundant rows
table = []
last_dash_index = -1  # position of last '-' in previous row

for i, row_val in enumerate(indices[:-1]):  # skip last row
    row = ['|']
    row_dash_index = -1
    for j, col_val in enumerate(indices):
        if j < i:
            row.append('   |')
        elif i == j:
            row.append(' - |')
            row_dash_index = max(row_dash_index, j)
        else:
            val = check_pair(row_val, col_val)
            row.append(f' {val} |')
            if val == '-':
                row_dash_index = max(row_dash_index, j)

    # Skip row if last '-' is at the same column as previous
    if row_dash_index != last_dash_index:
        table.append(row)
        last_dash_index = row_dash_index

    # Stop if last column is '-'
    if row_dash_index == len(indices) - 1:
        break


#*********************************************************************************************************
# Generate letter indicators below the table
letter_row = ['|']
for col_idx in range(len(indices)):
    letters = ''
    for row_idx, row in enumerate(table):
        cell_value = row[col_idx + 1]  # komórka w tabeli
        if '-' in cell_value:
            letters += string.ascii_lowercase[row_idx]  # litera wg numeru wiersza
    letter_row.append(f' {letters} |')

# Append the letter row to the table
table.append(letter_row)

# Print header
header = '| ' + ' | '.join(map(str, indices)) + ' |'
print(header)

# Print each row including the letter row
for row in table:
    print(''.join(row))