import pandas as pd
import json

def excel_to_json_hierarchy(file_path, sheet_name):
    # Load the Excel sheet into a DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="A:D", skiprows=3)

    # Creating a nested dictionary to hold the hierarchical relationships
    taxonomy_hierarchy = {}

    # Build the hierarchy based on columns A to D
    for _, row in df.iterrows():
        level2, level3, level4, level5 = row

        if pd.notna(level2):
            if level2 not in taxonomy_hierarchy:
                taxonomy_hierarchy[level2] = {}

            if pd.notna(level3):
                if level3 not in taxonomy_hierarchy[level2]:
                    taxonomy_hierarchy[level2][level3] = {}

                if pd.notna(level4):
                    if level4 not in taxonomy_hierarchy[level2][level3]:
                        taxonomy_hierarchy[level2][level3][level4] = []

                    if pd.notna(level5) and level5 not in taxonomy_hierarchy[level2][level3][level4]:
                        taxonomy_hierarchy[level2][level3][level4].append(level5)

    # Convert the dictionary to JSON
    with open('taxonomy_hierarchy.json', 'w') as json_file:
        json.dump(taxonomy_hierarchy, json_file, indent=4)

    print("JSON file has been created and saved.")

# Specify the path to your Excel file and the sheet name
excel_file_path = 'data.xlsx'
sheet_name = 'Tab 2'  # Make sure to use the correct sheet name
excel_to_json_hierarchy(excel_file_path, sheet_name)
