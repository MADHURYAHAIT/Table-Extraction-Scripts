# import fitz
# import os


# image_folder = 'images'
# if not os.path.exists(image_folder):
#     os.makedirs(image_folder)

# # Open the PDF file
# doc = fitz.open('/Users/mxdy/Developer/Table-Extraction-Scripts/PdfTableExtractionScript/Input/Gallery Specialties 2023 Pricelist.pdf')

# # Iterate over each page
# for page_num in range(len(doc)):
#     page = doc[page_num]
#     # Set the matrix to scale up the image to 300 dpi
#     zoom_matrix = fitz.Matrix(4, 4)  # 3x3 matrix, scales up by 3 in both x and y directions
#     # Render the page to an image
#     image = page.get_pixmap(matrix=zoom_matrix)
#     # Save the image to a file in the image folder
#     image.save(os.path.join(image_folder, f'page_{page_num+1}.png'))

# # Close the PDF file
# doc.close()

import pandas as pd
from PIL import Image
import pytesseract

# Load the image
image = Image.open('./images/page_7.png')

# Extract the table data using OCR
table_data = pytesseract.image_to_data(image, config='--psm 6')

# Split the data into rows and columns
rows = table_data.splitlines()
table = []
for row in rows:
    cells = row.split('\t')
    table.append(cells)

# Create a Pandas DataFrame from the table data
df = pd.DataFrame(table[1:], columns=table[0])

# Extract the column names
column_names = [i for i in df.text if i.isupper()]

# Create a new DataFrame with the extracted column names
new_df = pd.DataFrame(columns=column_names)

# Iterate over the rows and extract the data
for index, row in df.iterrows():
    product_number = ''
    description = ''
    finish = ''
    list_price = ''
    for i, cell in enumerate(row):
        if i == 0:
            product_number = cell
        elif i == 1:
            description = cell
        elif i == 2:
            finish = cell
        elif i == 3:
            list_price = cell
    new_row = pd.DataFrame({'PRODUCT NUMBER': [product_number], 'DESCRIPTION': [description], 'FINISH': [finish], 'LIST PRICE': [list_price]})
    new_df = pd.concat([new_df.reset_index(drop=True), new_row], ignore_index=True)

# Print the new DataFrame
print(new_df)
    
    
