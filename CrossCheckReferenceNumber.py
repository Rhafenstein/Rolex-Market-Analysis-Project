import pandas as pd

# Load Rolex data
file_path = 'data/Cleaned_Rolex_Chrono24.xlsx'
data = pd.read_excel(file_path)
# Remove rows
reference_numbers = data['Confirm Reference Number '].dropna()
reference_numbers = reference_numbers[~reference_numbers.str.contains("No Reference", case=False)].astype(str).tolist()  # Removes "No Reference"


# Reference guides
# Type reference guide
type_reference_guide = {
    "Submariner": [55, 140, 1140],
    "Submariner Date": [16, 166, 168, 1166, 1266],
    "GMT Master II": [167, 1167],
    "Day-Date (President)": [65, 66, 18, 180, 182, 183],
    "Datejust": [16, 162, 126],
    "Daytona Manual Wind": [62],
    "Daytona Cosmograph": [165, 1165],
    "Explorer II": [165, 2165],
    "Oyster Perpetual": [10, 140, 142, 176, 177, 116, 114, 124, 126, 276, 277, 67, 76],
    "Oysterquartz Datejust": [170],
    "Oysterquartz Day-Date": [190],
    "Yacht-Master": [166, 686, 696,],
    "Air-King": [55, 140, 114, 116],
    "Air-King Date": [57],
    "Oyster Perpetual Date": [15, 150, 115,],
    "Datejust 31": [68, 782, 1782, 278],
    "Datejust 36": [66, 16, 162, 1162, 1262, 1603, ],
    "Datejust II": [1163],
    "Datejust 41": [1263],
    "Day-Date 36": [65, 18, 180, 182, 183, 118, 128],
    "Day-Date II": [2182],
    "Day-Date 40": [228,],
    "Daytona": [62, 165, 1165],
    "Explorer": [60, 61, 63, 66, 10, 142, 1142, 2142, 224],
    "Lady-Datejust": [692, 679, 690, 691, 791, 1791, 2791, 2793, 2794],
    "Milgauss": [65, 10, 1164],
    "Sea-Dweller": [16, 166, 1166, 1266, 136],
    "Sea-Dweller-Deepsea": [136, 1260],
    "GMT-Master": [65, 167, 16, 167, 12671],
    "GMT-Master II": [167, 1167, 1267],
    "Oysterquartz": [170, 190, 191],
    "Oyster Precision": [64],
    "Yacht-Master 29": [696, 1696],
    "Yacht-Master 31": [686, 1686],
    "Yacht-Master 37": [268],
    "Yacht-Master 40": [166, 1166, 1266],
    "Yacht-Master 42": [2266],
    "Yacht-Master II": [1166],
    "Pearlmaster": list(range(80, 90)),
    "Sky-Dweller": [32, 33]
}

# Additional guides for bezel and case material
bezel_type_guide = {
    0: "Polished",
    1: "Finely Engine Turned",
    2: "Engine Turned",
    3: "Fluted",
    4: "Hand-Crafted",
    5: "Pyramid",
    6: "Rotating Bezel"
}

case_material_type_guide = {
    0: "Stainless",
    1: "Yellow Gold Filled",
    2: "White Gold Filled",
    3: "Stainless & Yellow Gold",
    4: "Stainless w/ 18k White Gold",
    5: "Gold Shell",
    6: "Platinum",
    7: "14k Yellow Gold",
    8: "18k Yellow Gold",
    9: "18k White Gold"
}

# Rolex reference letters codes
reference_letters_guide = {
    "gv": "Glace Verte (Green Crystal)",
    "lb": "Lunette Bleu (Blue Bezel)",
    "tbr": "Tsetse Baguette Diamonds",
    "ln": "Lunette Noir (Black Bezel)",
    "lv": "Lunette Verte (Green Bezel)",
    "blnr": "Bleu Noir (Blue Black)",
    "blro": "Bleu Rouge (Blue Red)",
    "chnr": "Chocolate Noir (Brown Black)",
    "rbow": "Rainbow (Multi-colored Sapphires)",
    "sabr": "Sapphirs Brilliants (Sapphires Diamonds)",
    "sanr": "Sapphirs Noir (Black Sapphires)",
    "saru": "Saphirs Rubis (Sapphires Rubies)",
    "sats": "Sapphirs Tsavorite (Sapphires Tsavorites)",
    "rbr": "Rhinestone Baguette Diamonds",
    "br": "Brilliant-cut Diamonds",
    "sa": "Set with Sapphires",
    "wa": "White Accent",
    "rt": "Rolesor (Two-tone - Steel and Gold)",
    "sr": "Steel and Rose Gold",
    "tt": "Two-tone (Mix of metals)",
    "yg": "Yellow Gold",
    "wg": "White Gold",
    "pg": "Pink Gold",
    "pvd": "Physical Vapor Deposition (Black Coating)",
    "pt": "Platinum",
    "vtnr": "Verte Noir (Green Black Bezel, Left-handed GMT)",
    "m": "Modified (Movement Update)",
    "grnr": "Gris Noir (Gray Black Bezel)"
}
extra_reference_map = {
    "Explorer II": [
        "226570"
    ],
    "Submariner": [
        "168000",
        "5508",
        "116610LV",
        "116619LB",
        "116610LN",
        "126610LN",
        "126619LB"
    ],
    "GMT Master": [
        "6542/8",
        "116710BLNR",
        "126710BLRO",
        "126710BLNR"
    ],
    "Yacht-Master": [
        "16622",
        "16628",
        "226659",
        "116655"
    ],
    "Air-King": [
        "116900"
    ],
    "Daytona": [
        "116500LN",
        "16520",
        "116520",
        "116509",
        "116506",
        "116519LN"
    ],
    "Sea-Dweller": [
        "126600",
        "116600",
        "126603"
    ],
    "Sea-Dweller-Deepsea": [
        "126660",
        "136660",
        "116660"
    ],
    "Datejust": [
        "126334",
        "126300",
        "116200",
        "116233",
        "116234",
        "116231"
    ],
    "Sky-Dweller": [
        "4000"
    ],
    "Oyster Perpetual Date": [
        "6924"
    ],
    "Cellini Moonphase": [
        "50535"
    ],
    "Cellini": [
        "4127"
    ],
    "Oysterquartz": [
        "19018",
        "19019",
        "17013",
        "17014"
    ]
}


# Function to analyze reference numbers
def analyze_reference(ref_num):

    numeric_part = ''.join(filter(str.isdigit, ref_num))
    reference_letters = ''.join(filter(str.isalpha, ref_num))

    model = None

    # Check for exact match in extra_reference_map
    for model_name, references in extra_reference_map.items():
        if ref_num in references:
            model = model_name
            break

    # If no exact match, perform progressive lookup
    if model is None and len(numeric_part) >= 2:
        for length in range(2, min(len(numeric_part), 5) + 1):
            model_code = int(numeric_part[:length])
            possible_models = [key for key, value in type_reference_guide.items() if model_code in value]
            if possible_models:
                model = possible_models[0]

    # Bezel and case material lookup for 5 or 6 digit references
    if len(numeric_part) >= 5:
        bezel_code = int(numeric_part[-2])
        case_material_code = int(numeric_part[-1])
        bezel = bezel_type_guide.get(bezel_code, "Unknown Bezel")
        case_material = case_material_type_guide.get(case_material_code, "Unknown Case Material")
    else:
        bezel = "N/A"
        case_material = "N/A"

    reference_letter_meaning = reference_letters_guide.get(reference_letters.lower(), "Unknown Reference Letters")

    return {
        "Reference Number": ref_num,
        "Model": model or "Unknown Model",
        "Bezel": bezel,
        "Case Material": case_material,
        "Reference Letters": reference_letters,
        "Reference Letters Meaning": reference_letter_meaning
    }


# Analyze the reference number "126234"
#result = analyze_reference("126234")

# Print the result for this one reference
#print(result)

# Analyze all reference numbers
results = [analyze_reference(ref_num) for ref_num in reference_numbers]

# Convert results to DataFrame
results_df = pd.DataFrame(results)

# Save the results
excel_file_path = r'data/Rolex_Reference.xlsx'
csv_file_path = r'data/Rolex_Reference.csv'
results_df.to_excel(excel_file_path, index=False)
results_df.to_csv(csv_file_path, index=False)

# print(f"DataFrame saved to:\n- Excel file: {excel_file_path}\n- CSV file: {csv_file_path}")
