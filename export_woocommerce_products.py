import openpyxl
from woocommerce import API

# Replace these values with your actual WooCommerce credentials
WC_URL = ""
WC_CONSUMER_KEY = ""
WC_CONSUMER_SECRET = ""

# Setup WooCommerce API connection
wcapi = API(
    url=WC_URL,
    consumer_key=WC_CONSUMER_KEY,
    consumer_secret=WC_CONSUMER_SECRET,
    version="wc/v3"
)

# Function to fetch all products from WooCommerce
def get_all_products():
    products = []
    page = 1
    while True:
        response = wcapi.get("products", params={"per_page": 100, "page": page}).json()
        if not response:
            break
        products.extend(response)
        page += 1
    return products

# Function to save product data to an Excel file
def save_products_to_excel(products):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Products"
    sheet.append(["ID", "Name", "Price", "Discount Price", "Stock Quantity"])

    for product in products:
        product_id = product.get('id', "N/A")
        name = product.get('name', "N/A")
        price = product.get('regular_price', "N/A")
        sale_price = product.get('sale_price', "N/A")
        stock_quantity = product.get('stock_quantity', "N/A")
        sheet.append([product_id, name, price, sale_price, stock_quantity])

    workbook.save("woocommerce_products.xlsx")
    print("Products saved to woocommerce_products.xlsx")

# Main execution
if __name__ == "__main__":
    products = get_all_products()
    save_products_to_excel(products)

