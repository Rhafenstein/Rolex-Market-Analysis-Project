import pandas as pd
import re
import pycountry

# Load the dataset
file_path = r'data/Rolex Chrono24.xlsx'
df_og = pd.read_excel(file_path)
df = df_og.copy()

#--------------------------------------------------------------------------------------------------------------------------------

df['Model'] = df['Model'].fillna('Unknown')

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

#--------------------------------------------------------------------------------------------------------------------------------

def extract_reference_number(x):
    ref_number = re.search(r'\b\d{4,}[A-Za-z0-9\-]*\b', str(x))
    return ref_number.group(0) if ref_number else "No Reference"

df['Reference number'] = df['Reference number'].apply(extract_reference_number)

#--------------------------------------------------------------------------------------------------------------------------------

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

df['Case material'] = df['Case material'].fillna('Unknown')

#--------------------------------------------------------------------------------------------------------------------------------

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

df['Bracelet material'] = df['Bracelet material'].fillna('Unknown')

#--------------------------------------------------------------------------------------------------------------------------------

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

#--------------------------------------------------------------------------------------------------------------------------------

def extract_box(val):
    if 'Original box' in str(val):
        return 'Yes'
    return 'No'

def extract_papers(val):
    if 'original papers' in str(val):
        return 'Yes'
    return 'No'

df['Box'] = df['Scope of delivery'].apply(extract_box)
df['Papers'] = df['Scope of delivery'].apply(extract_papers)

#--------------------------------------------------------------------------------------------------------------------------------

def clean_gender(row):
    for col in row.index:
        match = re.search(r"Men's watch/Unisex|Women's watch", str(row[col]))
        if match:
            return 'Male' if "Men's watch" in match.group(0) else 'Female'
    return 'Unknown'

df['Gender'] = df.apply(clean_gender, axis=1)

#--------------------------------------------------------------------------------------------------------------------------------

def split_location(val):
    # Try to split based on comma
    parts = str(val).split(',', 1)  # Split at the first comma only
    if len(parts) == 2:
        # Strip leading/trailing spaces from both parts
        country = parts[0].strip()
        city = parts[1].strip()
    else:
        # If no comma is found, we mark the country as the location and city as 'Unknown'
        country = parts[0].strip()
        city = 'Unknown'

    return pd.Series([country, city])

# Apply the function to create 'Country' and 'City' columns
df[['Country', 'City']] = df['Location'].apply(split_location)

valid_countries = [country.name for country in pycountry.countries]

# Ensure both empty strings and NaNs are handled in validation functions
def is_valid_country(country):
    country_str = str(country).strip().lower()  # Convert to lowercase for case-insensitive comparison
    if country_str in [c.lower() for c in valid_countries]:
        return country_str.title()  # Return the title case (e.g., 'United States')
    return None

# Function to check if a city name is valid (considered valid if it's a string and not too short)
def is_valid_city(city):
    city_str = str(city).strip()
    if len(city_str) > 2:  # Cities should generally be longer than 2 characters
        return city_str
    return None

# Function to clean the 'Country' column and check other columns if invalid
def clean_country(row):
    if is_valid_country(row['Country']):
        return row['Country']

    # Check other columns if the 'Country' column is invalid
    for col in row.index:
        valid_country = is_valid_country(row[col])
        if valid_country:
            return valid_country

    return 'Unknown'

# Function to clean the 'City' column and check other columns if invalid
def clean_city(row):
    if is_valid_city(row['City']):
        return row['City']

    # Check other columns if the 'City' column is invalid
    for col in row.index:
        valid_city = is_valid_city(row[col])
        if valid_city:
            return valid_city

    return 'Unknown'

# Apply the functions to clean the 'Country' and 'City' columns
df['Country'] = df.apply(clean_country, axis=1)
df['City'] = df.apply(clean_city, axis=1)

# Known abbreviations and their expansions
abbreviation_expansions = {
    "NSW": "New South Wales",
    "NYC": "New York City",
    "LA": "Los Angeles",
    "CA": "Canada",
    "UK": "United Kingdom",
    "US": "United States",
    "AU": "Australia",
    "FR": "France",
}

def is_real_city(city):
    city_str = str(city).strip()

    # Remove any parentheses and content inside them
    city_str = re.sub(r'\s*\(.*?\)\s*', '', city_str)

    # Check for abbreviations and expand them
    if city_str.upper() in abbreviation_expansions:
        city_str = abbreviation_expansions[city_str.upper()]

    # Pattern to detect if the city contains numbers or invalid characters
    if re.search(r'\d', city_str):  # Check if city contains digits
        return False
    if len(city_str) <= 2:  # Cities usually have more than 2 characters
        return False
    # Exclude city names with certain keywords
    invalid_keywords = ['unknown', 'papers', 'negotiable', 'au$', '=', 'original']
    if any(keyword in city_str.lower() for keyword in invalid_keywords):
        return False
    # Check if city is entirely uppercase (flag as abbreviation if too short)
    if city_str.isupper() and len(city_str) <= 4:
        return False
    return city_str.title()

# Apply the filter, remove parentheses, and convert to title case, expanding abbreviations
df['City'] = df['City'].apply(lambda x: is_real_city(x) if is_real_city(x) else 'Unknown')
#--------------------------------------------------------------------------------------------------------------------------------

def extract_currency(val):
    match = re.match(r'([A-Za-z]+)\$', str(val))
    if match:
        return match.group(1)
    return None

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

def clean_shipping_price(val):
    if pd.isna(val):
        return '0'
    cleaned_price = re.sub(r'[^\d]', '', str(val))
    return cleaned_price if cleaned_price else '0'

df['shipping'] = df['shipping'].apply(clean_shipping_price)

#--------------------------------------------------------------------------------------------------------------------------------

df['Seller Information'] = df['Seller Information'].fillna('Unknown')

#--------------------------------------------------------------------------------------------------------------------------------

def clean_base_caliber(val):
    match = re.search(r'\b\d{4}\b', str(val))
    if match:
        return match.group(0)
    return 'Unknown'

df['Base caliber'] = df['Base caliber'].apply(clean_base_caliber)

#--------------------------------------------------------------------------------------------------------------------------------

df = df.drop(columns=['Scope of delivery'])
df = df.drop(columns=['Location'])

#--------------------------------------------------------------------------------------------------------------------------------

excel_file_path = r'data/Cleaned_Rolex_Chrono24.xlsx'
csv_file_path = r'data/Cleaned_Rolex_Chrono24.csv'
df.to_excel(excel_file_path, index=False)
df.to_csv(csv_file_path, index=False)
print(f"DataFrame successfully saved to:\n- Excel file: {excel_file_path}\n- CSV file: {csv_file_path}")

#--------------------------------------------------------------------------------------------------------------------------------
