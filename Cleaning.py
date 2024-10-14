import pandas as pd
import re

# Load the dataset
file_path = r'data\Rolex Chrono24.xlsx'
df_og = pd.read_excel(file_path)
df = df_og.copy()

# Fill missing values in the 'Model' column
df['Model'].fillna('Unknown', inplace=True)

# Replace specific values in the 'Model' column
df['Model'] = df['Model'].replace({
    'Submariner (No Date)': 'Submariner',
    '6066': 'Oysterdate Precision',
    '116655': 'Yacht-Master 40',
    '17013': 'Datejust Oysterquartz',
    'Cartier Tank': 'Tank Must XL Silver Roman Dial',
    '15010': 'Oyster Perpetual Date',
    '16030': 'Datejust 36',
    '126503': 'Datejust 41',
    '40 224270': 'Explorer II',
    '124060': 'Submariner',
    '116680': 'Yacht-Master II',
    '69174': 'Lady-Datejust',
    '169623': 'Yacht-Master 37',
    'Rolex 1550': 'Oyster Perpetual Date',
    '18038A': 'Day-Date 36',
    'Manual winding': 'Precision',
    'ROLEX': 'Explorer I',
    '336934': 'Rolex Sky-Dweller',
    'Art Déco': 'Ladies Cocktail Art Déco silver dial Yellow Gold 14KT',
    '1500': 'Oyster Perpetual Date',
    'Automatic': 'Explorer I',
    'Explorer': 'Explorer I'
})

# Extract reference numbers using regex
def extract_reference_number(x):
    ref_number = re.search(r'\b\d{4,}[A-Za-z0-9\-]*\b', str(x))
    return ref_number.group(0) if ref_number else "No Reference"

df['Reference number'] = df['Reference number'].apply(extract_reference_number)

# Define patterns for watch movement types
movement_patterns = {
    'Automatic': r'(?i)\bauto(?:matic)?|self-winding\b',
    'Manual winding': r'(?i)\bmanual(?: winding| wind)?\b',
    'Quartz': r'(?i)\bquartz\b',
}

# Check if a value matches a known movement type
def is_valid_movement(val):
    val_str = str(val)
    for movement, pattern in movement_patterns.items():
        if re.search(pattern, val_str):
            return movement
    return None

# Apply function to correct the 'Movement' column
def correct_movement(row):
    movement = is_valid_movement(row['Movement'])
    if movement:
        return movement
    for col in row.index:
        valid_movement = is_valid_movement(row[col])
        if valid_movement:
            return valid_movement
    return row['Movement']

df['Movement'] = df.apply(correct_movement, axis=1)

# Save the cleaned DataFrame to Excel and CSV files
excel_file_path = r'data/Cleaned_Rolex_Chrono24.xlsx'
csv_file_path = r'data/Cleaned_Rolex_Chrono24.csv'
df.to_excel(excel_file_path, index=False)
df.to_csv(csv_file_path, index=False)
