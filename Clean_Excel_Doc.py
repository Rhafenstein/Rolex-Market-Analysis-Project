import importlib
import data_cleaning
importlib.reload(data_cleaning)

# Load the dataset
file_path = r'data\Chrono24_151024_excel.xlsx'
df_og = pd.read_excel(file_path)
df = df_og.copy()

# Apply data cleaning steps
df = data_cleaning.clean_data(df)

# Drop unnecessary columns
columns_to_drop = ['Scope of delivery', 'Location', 'Confirm Reference Number ']
df = df.drop(columns=columns_to_drop, errors='ignore')

# Rename columns if needed
if 'Seller Information' in df.columns:
    df.rename(columns={'Seller Information': 'Seller Name'}, inplace=True)

# Save the cleaned DataFrame to Excel and CSV files
excel_file_path = r'data/Cleaned_Rolex_Chrono24.xlsx'
csv_file_path = r'data/Cleaned_Rolex_Chrono24.csv'

df.to_excel(excel_file_path, index=False)
df.to_csv(csv_file_path, index=False)

print(f"DataFrame successfully saved to:\n- Excel file: {excel_file_path}\n- CSV file: {csv_file_path}")
