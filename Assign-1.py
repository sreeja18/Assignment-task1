import requests
from bs4 import BeautifulSoup
import openpyxl

# Read the HTML file
with open('Amazon.html', 'r', encoding='utf-8') as file:
    content = file.read()

# Parse the HTML content using BeautifulSoup
soup = BeautifulSoup(content, 'html.parser')

# Find all divs with the specified class
product_divs = soup.find_all('div', class_='s-card-container s-overflow-hidden aok-relative puis-wide-grid-style puis-wide-grid-style-t3 puis-include-content-margin puis puis-v3b48cl1js792724v4d69zlbwph s-latency-cf-section s-card-border')

# Initialize lists to store data
product_data = []

# Iterate through each product div
for div in product_divs:
    product_url_tag = div.find('a', class_='a-link-normal s-underline-text s-underline-link-text s-link-style a-text-normal')
    product_url = product_url_tag['href'] if product_url_tag else ""

    product_name_tag = div.find('span', class_='a-size-medium a-color-base a-text-normal')
    product_name = product_name_tag.get_text() if product_name_tag else ""

    ratings_tag = div.find('span', class_='a-size-base puis-normal-weight-text')
    ratings = ratings_tag.get_text() if ratings_tag else ""

    product_price_tag = div.find('span', class_='a-offscreen')
    product_price = product_price_tag.get_text() if product_price_tag else ""

    no_of_reviews_tag = div.find('span', class_='a-size-base s-underline-text')
    no_of_reviews = no_of_reviews_tag.get_text() if no_of_reviews_tag else ""

    product_data.append((product_url, product_name, ratings, product_price, no_of_reviews))

# Create or load an Excel file
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Amazon Products"

# Write headers
headers = ["Product URL", "Product Name", "Ratings", "Product Price", "No. of Reviews"]
ws.append(headers)

# Write product data to the Excel sheet
for product in product_data:
    ws.append(product)

# Save the Excel file
wb.save('Amazon_Products.xlsx')
