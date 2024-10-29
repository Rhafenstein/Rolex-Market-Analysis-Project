import re
import pandas as pd
import numpy as np
import pycountry

def clean_model(df):
    model_replacements = {
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
    }
    df['Model'] = df['Model'].fillna('Unknown').replace(model_replacements)
    return df

def extract_reference_number(df, columns):
    def extract_ref(x):
        ref_number = re.search(r'\b\d{4,}[A-Za-z0-9\-]*\b', str(x))
        return ref_number.group(0) if ref_number else "No Reference"
    for col in columns:
        df[col] = df[col].apply(extract_ref)
    return df

def clean_movement(df):
    movement_patterns = {
        'Automatic': r'(?i)\bauto(?:matic)?|self-winding\b',
        'Manual winding': r'(?i)\bmanual(?: winding| wind)?\b',
        'Quartz': r'(?i)\bquartz\b'
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
    return df

def clean_material(df, column_name, valid_materials):
    def is_valid_material(val):
        return val if val in valid_materials else None

    def correct_material(row):
        material = is_valid_material(row[column_name])
        if material:
            return material
        for col in row.index:
            valid_material = is_valid_material(row[col])
            if valid_material:
                return valid_material
        return 'Unknown'

    df[column_name] = df.apply(correct_material, axis=1)
    return df

def clean_year(df):
    year_pattern = r'\b(19[0-9]{2}|20[0-2][0-9])\b'

    def is_valid_year(val):
        match = re.search(year_pattern, str(val))
        return match.group(0) if match else None

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
    return df

def clean_condition(df):
    def extract_main_condition(val):
        return re.sub(r'\s*\(.*?\)', '', str(val).split('\n')[0].strip())

    def extract_condition_details(val):
        match = re.search(r'\((.*?)\)', str(val))
        return match.group(1) if match else 'Unknown'

    df['Condition Details'] = df['Condition'].apply(extract_condition_details)
    df['Condition'] = df['Condition'].apply(extract_main_condition).replace(['nan', 'NaN'], 'Unknown').fillna('Unknown')
    return df

def extract_box_and_papers(df):
    df['Box'] = df['Scope of delivery'].apply(lambda val: 'Yes' if 'Original box' in str(val) else 'No')
    df['Papers'] = df['Scope of delivery'].apply(lambda val: 'Yes' if 'original papers' in str(val) else 'No')
    return df

def clean_gender(df):
    def clean_gender_value(row):
        for col in row.index:
            match = re.search(r"Men's watch/Unisex|Women's watch", str(row[col]))
            if match:
                return 'Male' if "Men's watch" in match.group(0) else 'Female'
        return 'Unknown'

    df['Gender'] = df.apply(clean_gender_value, axis=1)
    return df

def split_location(df):
    def split_location_value(val):
        parts = str(val).split(',', 1)
        country = parts[0].strip()
        city = parts[1].strip() if len(parts) == 2 else 'Unknown'
        return pd.Series([country, city])

    df[['Country', 'City']] = df['Location'].apply(split_location_value)
    return df

def clean_country_and_city(df):
    valid_countries = [country.name for country in pycountry.countries]

    def is_valid_country(country):
        country_str = str(country).strip().lower()
        return country_str.title() if country_str in [c.lower() for c in valid_countries] else None

    def is_valid_city(city):
        city_str = str(city).strip()
        return city_str if len(city_str) > 2 else None

    df['Country'] = df.apply(lambda row: is_valid_country(row['Country']) or 'Unknown', axis=1)
    df['City'] = df.apply(lambda row: is_valid_city(row['City']) or 'Unknown', axis=1)
    return df

def clean_price(df):
    def extract_currency(val):
        match = re.match(r'([A-Za-z]+)\$', str(val))
        return match.group(1) if match else None

    def clean_price_value(val):
        return re.sub(r'[^\d]', '', str(val)) if pd.notna(val) else 'Unknown'

    df['Price'] = df.apply(lambda row: clean_price_value(row['Price']) if extract_currency(row['Price']) else 'Unknown', axis=1)
    return df

def clean_shipping_price(df):
    df['shipping'] = df['shipping'].apply(lambda val: re.sub(r'[^\d]', '', str(val)) if pd.notna(val) else '0')
    return df

def fill_missing_seller_info(df):
    df['Seller Information'] = df['Seller Information'].fillna('Unknown')
    return df

def clean_base_caliber(df):
    df['Base caliber'] = df['Base caliber'].apply(lambda val: re.search(r'\b\d{4}\b', str(val)).group(0) if re.search(r'\b\d{4}\b', str(val)) else 'Unknown')
    return df

def fix_dial(df):
    valid_colors = ['Green', 'Black', 'White', 'Silver', 'Brown', np.nan, 'Gold', 'Blue', 'Champagne', 'Grey', 'Pink',
                    'Turquoise', 'Purple', 'Mother of pearl', 'Bronze', 'Meteorite', 'Red', 'Yellow', 'Lines',
                    'Orange', 'Bordeaux', 'Skeletonized']

    def correct_dial_color(row):
        dial_color = row['Dial']
        if dial_color in valid_colors:
            return dial_color
        for col in ['Case material', 'Bracelet material']:
            if row[col] in valid_colors:
                return row[col]
        return 'Unknown'

    df['Dial'] = df.apply(correct_dial_color, axis=1)
    return df

def fix_dial_numerals(df):
    valid_numerals = ['No numerals', 'Arabic numerals', 'Roman numerals', np.nan]

    def correct_dial_numerals(row):
        dial_numerals = row['Dial numerals']
        if dial_numerals in valid_numerals:
            return dial_numerals
        for col in ['Case material', 'Bracelet material']:
            if row[col] in valid_numerals:
                return row[col]
        return np.nan

    df['Dial numerals'] = df.apply(correct_dial_numerals, axis=1).replace({
        'No numerals': 'None',
        'Arabic numerals': 'Arabic',
        'Roman numerals': 'Roman'
    }).fillna('Unknown')
    return df
# Main function to execute all cleaning steps
def clean_data(df):
    # Include the data cleaning steps from the refactored example
    df = clean_model(df)
    df = extract_reference_number(df, ['Reference number', 'Confirm Reference Number '])
    df = clean_movement(df)
    df = clean_material(df, 'Case material', [
        'Steel', 'Rose gold', 'Gold/Steel', 'White gold', 'Yellow gold',
        'Platinum', 'Red gold', 'Titanium', 'Ceramic', 'Rubber', 'Silver',
        'Gold-plated', 'Leather', 'Unknown'
    ])
    df = clean_material(df, 'Bracelet material', [
        'Steel', 'Rose gold', 'Gold/Steel', 'White gold', 'Yellow gold',
        'Platinum', 'Red gold', 'Titanium', 'Ceramic', 'Rubber', 'Silver',
        'Gold-plated', 'Leather', 'Crocodile skin', 'Calf skin', 'Snake skin',
        'Lizard skin', 'Alligator skin', 'Unknown'
    ])
    df = clean_year(df)
    df = clean_condition(df)
    df = extract_box_and_papers(df)
    df = clean_gender(df)
    df = split_location(df)
    df = clean_country_and_city(df)
    df = clean_price(df)
    df = clean_shipping_price(df)
    df = fill_missing_seller_info(df)
    df = clean_base_caliber(df)
    df = fix_dial(df)
    df = fix_dial_numerals(df)
    return df