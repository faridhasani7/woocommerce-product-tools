# WooCommerce Product Tools

This repository contains a set of Python scripts to manage WooCommerce products using the WooCommerce API. The tools include:

- Exporting products from WooCommerce to an Excel file.
- Updating WooCommerce products based on Excel data.
- Adding new product attributes from an Excel file.
- Adding or updating product attributes, including creating new products if they do not exist.

## Project Files

1. **export_woocommerce_products.py**  
   Exports product details (ID, Name, Price, Discount Price, Stock Quantity) from WooCommerce to `woocommerce_products.xlsx`.  
   **Excel Table Structure:**
   
   | ID   | Name         | Price | Discount Price | Stock Quantity |
   |------|--------------|-------|----------------|----------------|
   
2. **update_woocommerce_products.py**  
   Reads `woocommerce_products.xlsx` and updates existing products in WooCommerce.  
   **Excel Table Structure:**  
   Same as above.
   
3. **add_attribute_woocommerce.py**  
   Reads attributes from `attrs_bundle.xlsx` and adds new attributes to WooCommerce.  
   **Excel Table Structure:**
   
   | attr name | attr slug |
   |-----------|-----------|
   
4. **new_add_attribute_product3.py**  
   Reads product attributes from `product_attributes.xlsx` and updates or creates a product with the given attributes in WooCommerce.  
   **Excel Table Structure:**
   
   | Product Name | Attribute ID | Attribute Name | Attribute Value |
   |--------------|--------------|----------------|-----------------|

## Setup

1. **Install Dependencies:**  
   Run the following command to install the required packages:
   
   ```bash
   pip install -r requirements.txt
