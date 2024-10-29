import pandas as pd
import requests
from bs4 import BeautifulSoup

# Web scraping function for Chrono24
def scrape_chrono24(url):
    try:
        response = requests.get(url)
        response.raise_for_status()  # Check if the request was successful
    except requests.RequestException as e:
        print(f"Failed to retrieve data. Error: {e}")
        return pd.DataFrame()

    # Parse the HTML content
    soup = BeautifulSoup(response.content, 'html.parser')

    # Lists to store scraped data
    models = []
    reference_numbers = []
    movements = []
    case_materials = []
    bracelet_materials = []
    genders = []
    locations = []
    bezel_materials = []
    dials = []
    dial_numerals = []
    scope_of_deliveries = []
    bracelet_colors = []
    clasps = []
    clasp_materials = []


    # Find all the watch listings on the page
    listings = soup.find_all('div', class_='article-item-container')

    for listing in listings:
        # Extract the model name
        model = listing.find('strong')
        model_text = model.get_text(strip=True) if model else 'Unknown'
        models.append(model_text)

        # Extract the reference number, which is usually the text below the model
        reference_number = listing.find(text=True, recursive=False)
        reference_number_text = reference_number.strip() if reference_number else 'No Reference'
        reference_numbers.append(reference_number_text)

# Create a DataFrame with the scraped data
    scraped_data = pd.DataFrame({
        'Model': models,
        'Reference Number': reference_numbers,
        'Movement': movements,
        'Case Material': case_materials,
        'Bracelet Material': bracelet_materials,
        'Gender': genders,
        'Location': locations,
        'Bezel Material': bezel_materials,
        'Dial': dials,
        'Dial Numerals': dial_numerals,
        'Scope of Delivery': scope_of_deliveries,
        'Bracelet Color': bracelet_colors,
        'Clasp': clasps,
        'Clasp Material': clasp_materials
    })
    return scraped_data

# URL for Chrono24 Rolex page
url = 'https://www.chrono24.com.au/rolex/index.htm'

# Scrape data from the website
scraped_df = scrape_chrono24(url)

    # # Save the scraped DataFrame to Excel and CSV files
    # excel_file_path = 'data/Scraped_Chrono24_Rolex.xlsx'
    # csv_file_path = 'data/Scraped_Chrono24_Rolex.csv'
    # scraped_df.to_excel(excel_file_path, index=False)
    # scraped_df.to_csv(csv_file_path, index=False)