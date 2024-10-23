import pandas as pd
import re
import pycountry
import numpy as np

# Load the dataset
file_path = r'data\Chrono24_151024_excel.xlsx'
df_og = pd.read_excel(file_path)
df = df_og.copy()

#--------------------------------------------------------------------------------------------------------------------------------

# Fill missing values and replace specific Model names
df['Model'] = df['Model'].fillna('Unknown').replace({
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

#--------------------------------------------------------------------------------------------------------------------------------

# Extract reference numbers
def extract_reference_number(x):
    ref_number = re.search(r'\b\d{4,}[A-Za-z0-9\-]*\b', str(x))
    return ref_number.group(0) if ref_number else "No Reference"

df['Reference number'] = df['Reference number'].apply(extract_reference_number)
df['Confirm Reference Number '] = df['Confirm Reference Number '].apply(extract_reference_number)

#--------------------------------------------------------------------------------------------------------------------------------

# Correct movement values
movement_patterns = {
    'Automatic': r'(?i)\bauto(?:matic)?|self-winding\b',
    'Manual winding': r'(?i)\bmanual(?: winding| wind)?\b',
    'Quartz': r'(?i)\bquartz\b',
}

def is_valid_movement(val):
    val_str = str(val)
    for movement, pattern in movement_patterns.items():
        if re.search(pattern, val_str):
            return movement
    return None

def correct_movement(row):
    movement = is_valid_movement(row['Movement'])
    if movement:
        return movement
    for col in row.index:
        valid_movement = is_valid_movement(row[col])
        if valid_movement:
            return valid_movement
    return 'Unknown'

df['Movement'] = df.apply(correct_movement, axis=1)

#--------------------------------------------------------------------------------------------------------------------------------

# Correct case material values
valid_case_materials = [
    'Steel', 'Rose gold', 'Gold/Steel', 'White gold', 'Yellow gold',
    'Platinum', 'Red gold', 'Titanium', 'Ceramic', 'Rubber',
    'Silver', 'Gold-plated', 'Leather', 'Unknown'
]

def is_valid_case_material(val):
    val_str = str(val)
    if val_str in valid_case_materials:
        return val_str
    return None

def correct_case_material(row):
    case_material = is_valid_case_material(row['Case material'])
    if case_material:
        return case_material
    for col in row.index:
        valid_material = is_valid_case_material(row[col])
        if valid_material:
            return valid_material
    return 'Unknown'

df['Case material'] = df.apply(correct_case_material, axis=1)

#--------------------------------------------------------------------------------------------------------------------------------

# Correct bracelet material values
valid_bracelet_materials = [
    'Steel', 'Rose gold', 'Gold/Steel', 'White gold', 'Yellow gold',
    'Platinum', 'Red gold', 'Titanium', 'Ceramic', 'Rubber',
    'Silver', 'Gold-plated', 'Leather', 'Crocodile skin', 'Calf skin',
    'Snake skin', 'Lizard skin', 'Alligator skin', 'Unknown'
]

def is_valid_bracelet_material(val):
    val_str = str(val)
    if val_str in valid_bracelet_materials:
        return val_str
    return None

def correct_bracelet_material(row):
    bracelet_material = is_valid_bracelet_material(row['Bracelet material'])
    if bracelet_material:
        return bracelet_material
    for col in row.index:
        valid_material = is_valid_bracelet_material(row[col])
        if valid_material:
            return valid_material
    return 'Unknown'

df['Bracelet material'] = df.apply(correct_bracelet_material, axis=1)

#--------------------------------------------------------------------------------------------------------------------------------

# Correct year values
year_pattern = r'\b(19[0-9]{2}|20[0-2][0-9])\b'

def is_valid_year(val):
    val_str = str(val)
    match = re.search(year_pattern, val_str)
    if match:
        return match.group(0)
    return None

def correct_year(row):
    year = is_valid_year(row['Year of production'])
    if year:
        return year
    for col in row.index:
        valid_year = is_valid_year(row[col])
        if valid_year:
            return valid_year
    return 'Unknown'

df['Year of production'] = df.apply(correct_year, axis=1)

#--------------------------------------------------------------------------------------------------------------------------------

# Clean condition details
def extract_main_condition(val):
    val = str(val).split('\n')[0].strip()
    return re.sub(r'\s*\(.*?\)', '', val).strip()

def extract_condition_details(val):
    match = re.search(r'\((.*?)\)', str(val))
    if match:
        return match.group(1)
    return 'Unknown'

df['Condition Details'] = df['Condition'].apply(extract_condition_details)
df['Condition'] = df['Condition'].apply(extract_main_condition)

df['Condition'] = df['Condition'].replace(['nan', 'NaN'], 'Unknown').fillna('Unknown')
#--------------------------------------------------------------------------------------------------------------------------------

# Extract box and papers info
def extract_box(val):
    return 'Yes' if 'Original box' in str(val) else 'No'

def extract_papers(val):
    return 'Yes' if 'original papers' in str(val) else 'No'

df['Box'] = df['Scope of delivery'].apply(extract_box)
df['Papers'] = df['Scope of delivery'].apply(extract_papers)

#--------------------------------------------------------------------------------------------------------------------------------

# Clean gender values
def clean_gender(row):
    for col in row.index:
        match = re.search(r"Men's watch/Unisex|Women's watch", str(row[col]))
        if match:
            return 'Male' if "Men's watch" in match.group(0) else 'Female'
    return 'Unknown'

df['Gender'] = df.apply(clean_gender, axis=1)

#--------------------------------------------------------------------------------------------------------------------------------

# Split location into Country and City
def split_location(val):
    parts = str(val).split(',', 1)
    country = parts[0].strip()
    city = parts[1].strip() if len(parts) == 2 else 'Unknown'
    return pd.Series([country, city])

df[['Country', 'City']] = df['Location'].apply(split_location)

# Validate country and city values
valid_countries = [country.name for country in pycountry.countries]

def is_valid_country(country):
    country_str = str(country).strip().lower()
    return country_str.title() if country_str in [c.lower() for c in valid_countries] else None

def is_valid_city(city):
    city_str = str(city).strip()
    return city_str if len(city_str) > 2 else None

def clean_country(row):
    if is_valid_country(row['Country']):
        return row['Country']
    for col in row.index:
        valid_country = is_valid_country(row[col])
        if valid_country:
            return valid_country
    return 'Unknown'

def clean_city(row):
    if is_valid_city(row['City']):
        return row['City']
    for col in row.index:
        valid_city = is_valid_city(row[col])
        if valid_city:
            return valid_city
    return 'Unknown'

df['Country'] = df.apply(clean_country, axis=1)
df['City'] = df.apply(clean_city, axis=1)

#--------------------------------------------------------------------------------------------------------------------------------

# Clean price
def extract_currency(val):
    match = re.match(r'([A-Za-z]+)\$', str(val))
    return match.group(1) if match else None

def clean_price(val):
    cleaned_price = re.sub(r'[^\d]', '', str(val))
    return cleaned_price if cleaned_price else 'Unknown'

def correct_price(row):
    if extract_currency(row['Price']):
        return clean_price(row['Price'])
    for col in row.index:
        if extract_currency(row[col]):
            return clean_price(row[col])
    return 'Unknown'

df['Price'] = df.apply(correct_price, axis=1)

#--------------------------------------------------------------------------------------------------------------------------------

# Clean shipping price
def clean_shipping_price(val):
    return re.sub(r'[^\d]', '', str(val)) if pd.notna(val) else '0'

df['shipping'] = df['shipping'].apply(clean_shipping_price)

#--------------------------------------------------------------------------------------------------------------------------------

# Fill missing seller information
df['Seller Information'] = df['Seller Information'].fillna('Unknown')

#--------------------------------------------------------------------------------------------------------------------------------

# Clean base caliber
def clean_base_caliber(val):
    match = re.search(r'\b\d{4}\b', str(val))
    return match.group(0) if match else 'Unknown'

df['Base caliber'] = df['Base caliber'].apply(clean_base_caliber)

#--------------------------------------------------------------------------------------------------------------------------------

# Fix dial color
valid_colors = ['Green', 'Black', 'White', 'Silver', 'Brown', np.nan, 'Gold', 'Blue',
                'Champagne', 'Grey', 'Pink', 'Turquoise', 'Purple', 'Mother of pearl',
                'Bronze', 'Meteorite', 'Red', 'Yellow', 'Lines', 'Orange', 'Bordeaux', 'Skeletonized']

def fix_dial_color(row):
    dial_color = row['Dial']
    if dial_color in valid_colors:
        return dial_color
    for col in ['Case material', 'Bracelet material']:
        if row[col] in valid_colors:
            return row[col]
    return 'Unknown'

df['Dial'] = df.apply(fix_dial_color, axis=1)

#--------------------------------------------------------------------------------------------------------------------------------

# Fix dial numerals
valid_numerals = ['No numerals', 'Arabic numerals', 'Roman numerals', np.nan]

def fix_dial_numerals(row):
    dial_numerals = row['Dial numerals']
    if dial_numerals in valid_numerals:
        return dial_numerals
    for col in ['Case material', 'Bracelet material']:
        if row[col] in valid_numerals:
            return row[col]
    return np.nan

df['Dial numerals'] = df.apply(fix_dial_numerals, axis=1).replace({
    'No numerals': 'None',
    'Arabic numerals': 'Arabic',
    'Roman numerals': 'Roman'
}).fillna('Unknown')

#--------------------------------------------------------------------------------------------------------------------------------

# Clean up columns and save files
df = df.drop(columns=['Scope of delivery', 'Location'])
df.rename(columns={'Seller Information': 'Seller Name'}, inplace=True)

excel_file_path = r'data/Cleaned_Rolex_Chrono24.xlsx'
csv_file_path = r'data/Cleaned_Rolex_Chrono24.csv'
df.to_excel(excel_file_path, index=False)
df.to_csv(csv_file_path, index=False)
print(f"DataFrame successfully saved to:\n- Excel file: {excel_file_path}\n- CSV file: {csv_file_path}")
