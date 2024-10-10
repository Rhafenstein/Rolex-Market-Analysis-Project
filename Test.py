import pandas as pd
import re

file_path = r'data\Rolex Chrono24.xlsx'
df_og = pd.read_excel(file_path)
df = df_og.copy()

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
# Function to extract valid reference numbers that start with a number
def extract_reference_number(x):
    # Extract sequence that starts with a number, followed by at least 3 alphanumeric characters
    ref_number = re.search(r'\b[0-9][A-Za-z0-9]{3,}\b', str(x))  # {3,} ensures at least 4 characters in total

    if ref_number:
        return ref_number.group(0)  # Return the matched reference number
    else:
        return "No Reference"

# Save the df as an excel and csv file
excel_file_path = r'data/Cleaned_Rolex_Chrono24.xlsx'
df.to_excel(excel_file_path, index=False)
csv_file_path = r'data/Cleaned_Rolex_Chrono24.csv'
df.to_csv(csv_file_path, index=False)

print(f"Data saved to {excel_file_path} and {csv_file_path}")