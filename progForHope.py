import itertools
import pandas as pd
import os

def abbreviate_color(color):
    color = color.strip()  
    vowels = 'aeiouAEIOU' 
    for vowel in vowels:
        color = color.replace(vowel, '')
    return color

def truncate_id_sections(id, delimiter='-'):
    id_parts = id.split('_')

    # Keep parent SKU and source as they are
    parent_sku = id_parts[0]
    source = id_parts[1]

    # Truncate the rest of the sections if they exist
    for i in range(2, len(id_parts)):
        while len(id_parts[i]) > 1 and len("_".join(id_parts)) > 42:
            section_parts = id_parts[i].split(delimiter)
            for j in range(len(section_parts)):
                if len(section_parts[j]) > 1:  # If the section has more than 1 character, truncate it.
                    section_parts[j] = section_parts[j][:-1]
            id_parts[i] = delimiter.join(section_parts)

    id = "_".join(id_parts)
    return id


def generate_combinations(products_df):
    combinations = []
    for _, product_row in products_df.iterrows():
        # Skip row if necessary data is missing
        if pd.isnull(product_row['post_title']):
            continue
        
        product_name = product_row['post_title']
        parent_sku = product_row['sku']
        price = product_row['regular_price']
        colors = product_row['attribute:pa_color'].split('|') if isinstance(product_row['attribute:pa_color'], str) else [None]
        abbr_colors = product_row['AbbrLogoColor'].split('|') if isinstance(product_row['AbbrLogoColor'], str) else [None]
        logoColorLong = product_row['attribute:pa_logo'].split('|') if isinstance(product_row['attribute:pa_logo'], str) else [None]
        locations = product_row['AbbrLocation'].split('|') if isinstance(product_row['AbbrLocation'], str) else [None]
        locationsLong = product_row['attribute:pa_location'].split('|') if isinstance(product_row['attribute:pa_location'], str) else [None]

        if not colors or not abbr_colors or not locations or not logoColorLong or not locationsLong: 
            print(f"Skipping row {_} due to missing colors or locations.")
            continue

        # Create pairs of shortened and longer versions
        color_pairs = list(zip(colors, colors))  # Use colors as colorLong
        logo_color_pairs = list(zip(abbr_colors, logoColorLong))
        location_pairs = list(zip(locations, locationsLong)) 

        for (color, color_long), (abbr_color, logo_color_long), (location, location_long) in itertools.product(color_pairs, logo_color_pairs, location_pairs):
            color = color and abbreviate_color(color)  # Abbreviate the color only if it exists
            id_parts = [str(parent_sku), str(product_row['Source'])]
            if color:
                id_parts.append(color)
            if abbr_color:
                id_parts.append(abbr_color)
            if location:
                id_parts.append(location)
            id = "_".join(id_parts)
            id = truncate_id_sections(id)

            bc_count = len(id)
            combination = {
                'post_title': product_name,
                'parent_sku': parent_sku,
                'ID': '',
                'post_status': 'publish',
                'sku': id,
                'BC Count': bc_count,
                'downloadable': 'no',
                'virtual': 'no',
                'stock': '', 
                'stock_status': 'instock',
                'regular_price': price,
                'tax:product_visibility': 'visible',
                'AbbrColor': color if color else 'N/A',
                'meta:attribute_pa_color': color_long,
                'AbbrLogo': abbr_color if abbr_color else 'N/A',
                'meta:attribute_pa_logo': logo_color_long,
                'AbbrLocation': location if location else 'N/A',
                'meta:attribute_pa_location': location_long
            }
            combinations.append(combination)

    return combinations



try:
    products_df = pd.read_excel('C:/testing/bookForTesting.xlsx', engine='openpyxl')
    print("Excel file loaded successfully.")

    combinations = generate_combinations(products_df)
    df = pd.DataFrame(combinations)
    try:
        df.to_excel('C:/testing/variationsForHope.xlsx', index=False)

    except PermissionError:
        print("File is not writable.")
    except FileNotFoundError:
        print("Directory does not exist.")
    except Exception as e:
        print(f"An error occurred: {e}")
        
except FileNotFoundError:
    print("File could not be opened")
