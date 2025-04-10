import pandas as pd
from woocommerce import API

# Function to convert attribute value to string if needed
def clean_attr(attr):
    if not isinstance(attr, str):
        return str(attr)
    return attr

# Read Excel file containing attributes data
# Expected Excel file name: 'attrs_bundle.xlsx'
df = pd.read_excel('attrs_bundle.xlsx')

# Create a list of attributes with name and slug
attributes_list = []
for index, row in df.iterrows():
    # Using column names: 'attr name' and 'attr slug'
    attr_name = clean_attr(row['attr name'])
    attr_slug = clean_attr(row['attr slug'])
    attributes_list.append({"name": attr_name, "slug": attr_slug})

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
    version="wc/v3"
)

# Fetch existing attributes from WooCommerce
attributes_response = wcapi.get("products/attributes")
existing_attributes = {attribute['name']: attribute['id'] for attribute in attributes_response.json()}

# Loop through the attributes list and add new attributes to WooCommerce if they do not exist
for attr in attributes_list:
    attr_name = attr['name']
    if attr_name not in existing_attributes:
        # Data payload for creating a new attribute
        new_attribute_data = {
            "name": attr_name,
            "slug": attr['slug'],
            "type": "select",
            "order_by": "menu_order",
            "has_archives": True
        }
        response = wcapi.post("products/attributes", new_attribute_data)
        if response.status_code == 201:
            attr_id = response.json()['id']
            existing_attributes[attr_name] = attr_id
            print(f"Attribute '{attr_name}' created with ID {attr_id}.")
        else:
            print(f"Error creating attribute '{attr_name}': {response.text}")

input("Press Enter to exit...")

