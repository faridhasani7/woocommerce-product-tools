import sys
sys.stdout.reconfigure(encoding='utf-8')

import pandas as pd
from woocommerce import API
import json

# Function to convert attribute value to string if needed
def clean_attr(attr):
    if not isinstance(attr, str):
        return str(attr)
    return attr

# Utility function to check if main_list contains all dictionaries in sublist
def contains_dicts(main_list, sublist):
    return all(any(d == sub_dict for d in main_list) for sub_dict in sublist)

# Read Excel file containing product and attribute data
# Expected Excel file name: 'product_attributes.xlsx'
df = pd.read_excel('product_attributes.xlsx')

# Get product name from the first cell (assumed location of product name)
product_name = df.iloc[0, 0]
print(f"Product Name: {product_name}")

# Create a list of attribute dictionaries with id, name, and value
product_attributes = []
for index, row in df.iterrows():
    # Assuming the following order in Excel:
    # Column 1: Product Name, Column 2: Attribute ID, Column 3: Attribute Name, Column 4: Attribute Value
    attr_id = row.iloc[1]
    attr_name = clean_attr(row.iloc[2])
    attr_value = clean_attr(row.iloc[3])
    product_attributes.append({"name": attr_name, "id": attr_id, "value": attr_value})

# WooCommerce API credentials (fill these with your actual credentials)
WC_URL = "https://your_store_url.com/"
WC_CONSUMER_KEY = ""
WC_CONSUMER_SECRET = ""

# Setup WooCommerce API connection
wcapi = API(
    url=WC_URL,
    consumer_key=WC_CONSUMER_KEY,
    consumer_secret=WC_CONSUMER_SECRET,
    wp_api=True,
    version="wc/v3",
    timeout=20
)

# Function to fetch all products from WooCommerce (paginated)
def get_all_products():
    products = []
    page = 1
    while True:
        print("Fetching products, page", page)
        response = wcapi.get("products", params={"per_page": 100, "page": page}).json()
        if not response:
            break
        products.extend(response)
        page += 1
    return products

# Get existing products from WooCommerce by their names and IDs
products_response = get_all_products()
existing_products = {product['name']: product['id'] for product in products_response}

# Function to fetch all attributes from WooCommerce (paginated)
def get_all_product_attributes():
    page = 1
    all_attributes = []
    while True:
        print("Fetching attributes, page", page, "Total attributes:", len(all_attributes))
        response = wcapi.get("products/attributes", params={"page": page, "per_page": 100})
        if response.status_code != 200:
            print(f"Error: {response.status_code}")
            break
        attributes = response.json()
        if not attributes or contains_dicts(all_attributes, attributes):
            break
        all_attributes.extend(attributes)
        page += 1
    return all_attributes

# Get existing product attributes from WooCommerce
attributes_response = get_all_product_attributes()
existing_attributes = {attribute['name']: attribute['id'] for attribute in attributes_response}

# Check if the product exists
if product_name in existing_products:
    product_id = existing_products[product_name]
    print(f"Product '{product_name}' exists. ID: {product_id}")
    
    # Fetch existing attributes for the product
    existing_product_response = wcapi.get(f"products/{product_id}")
    existing_product_attributes = existing_product_response.json().get('attributes', [])
    
    # Add new attributes that are not already present in the product
    for attr in product_attributes:
        attr_name = attr['name']
        attr_id = existing_attributes[attr_name]
        # Check if attribute is already assigned to the product
        if attr_id not in [existing_attr['id'] for existing_attr in existing_product_attributes]:
            existing_product_attributes.append({
                'id': attr_id,
                'name': attr_name,
                'options': [attr['value']],
                "visible": True
            })
    # Update product with new attributes
    updated_data = {
        "attributes": existing_product_attributes
    }
    wcapi.put(f"products/{product_id}", updated_data)
    print(f"New attributes added to product '{product_name}'.")
else:
    # If the product does not exist, create a new product with the attributes
    new_product_data = {
        "name": product_name,
        "attributes": []
    }
    for attr in product_attributes:
        attr_name = attr['name']
        attr_id = existing_attributes[attr_name]
        new_product_data["attributes"].append({
            "id": attr_id,
            "options": [attr['value']],
            "visible": True
        })
    # Create new product via WooCommerce API
    response = wcapi.post("products", new_product_data)
    if response.status_code == 201:
        print(f"Product '{product_name}' added successfully.")
    else:
        print(f"Error adding product '{product_name}': {json.dumps(response.json(), indent=0, ensure_ascii=False)}")

input("Press Enter to exit...")

