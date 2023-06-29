import itertools
import pandas as pd    #used to read an excel file and load the fata into a Dataframe
import os

## use itertools for generating combinations of products
## use pandas for handling data and exporting to a CSV file


def abbreviate_color(color):
    """Remove vowels from colors"""
    color = color.strip()                                       # Ensure no leading/trailing white spaces
    vowels = 'aeiouAEIOU'                                       # List of vowels
    for vowel in vowels:
        color = color.replace(vowel, '')                        # Remove the vowel
    return color


def generate_combinations(products_df):
    combinations = []
    
    #iterrows iterates over each row extracting the information and assigning the value to a variable
    for _, product_row in products_df.iterrows():
        product_name = product_row['PRODUCT NAME']
        parent_sku = product_row['SKU #']
        product_number = product_row['PRODUCT #']
        colors = product_row['COLORS'].split('|') if isinstance(product_row['COLORS'], str) else []
        sizes = product_row['SIZES'].split('|')
        price = product_row['PRICE']

        # Create combinations of colors and sizes
        for color, size in itertools.product(colors, sizes):    #colors and sizes are lists  the function itertools.product produces tuples of all possible combintations of the elements
            abbreviated_color = abbreviate_color(color)         # Abbreviate the color
            id = f"{parent_sku}_{product_number}_{abbreviated_color}_{size}"     #f" is the start of a formatted string and it combines the values of the expressions inside
            id = id[:45]                                        # Ensure 'sku' is not more than 45 characters long
            bc_count = len(id)
            combination = {
                'Parent': product_name,
                'parent_sku': parent_sku,
                'PRODUCT #': product_number,
                'ID': '',
                'post_status': 'publish',
                'sku': id,
                'BC Count': bc_count,
                'downloadable': 'no',
                'virtual': 'no',
                'stock': '',                                    # update this accordingly.
                'stock_status': 'instock',
                'regular_price': price,
                'tax:product_visibility': 'visible',
                'meta:attribute_pa_color': color,
                'meta:attribute_pa_size': size
            }
            combinations.append(combination)

    return combinations

try:
    # Attempt to read the input Excel file
    products_df = pd.read_excel('C:/testing/Book1_.xlsx', engine='openpyxl', header=6)

    print("Excel file loaded successfully.")
    
    # Print column names
    print(products_df.columns)

    combinations = generate_combinations(products_df)
    
    # Convert the list of dictionaries to a DataFrame
    df = pd.DataFrame(combinations)
    
    # Try to write the DataFrame to an Excel file
    try:
        df.to_excel('C:/testing/variations.xlsx', index=False)
    except PermissionError:
        # If the file is not writable, this block will be executed
        print("File is not writable.")
    except FileNotFoundError:
        # If the directory does not exist, this block will be executed
        print("Directory does not exist.")
    except Exception as e:
        # If there is any other exception, this block will be executed
        print(f"An error occurred: {e}")
        
except FileNotFoundError:
    # If the file does not exist, this block will be executed
    print("File could not be opened")




