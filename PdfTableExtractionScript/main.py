import fitz
import os
from img2table.ocr import TesseractOCR
from img2table.document import Image
import pandas as pd
import shutil


import warnings
warnings.simplefilter(action='ignore', category=Warning)

def pdf_to_Img(foldername, location):
    print("\033[92m[INFO] Starting image conversion...\033[0m")
    image_folder = foldername
    if not os.path.exists(image_folder):
        os.makedirs(image_folder)

    # Open the PDF file
    doc = fitz.open(location)

    # Iterate over each page
    for page_num in range(len(doc)):
        page = doc[page_num]
        # Set the matrix to scale up the image to 400 dpi
        zoom_matrix = fitz.Matrix(4, 4)  # 4x4 matrix, scales up by 3 in both x and y directions
        # Render the page to an image
        image = page.get_pixmap(matrix=zoom_matrix)
        # Save the image to a file in the image folder
        image.save(os.path.join(image_folder, f'page_{page_num+1}.png'))
    doc.close()
    print("\033[92m[INFO] Image conversion completed!\033[0m")


def df_to_xlsx(df, columns, filename):
    print("\033[92m[INFO] Starting XLSX conversion...\033[0m")
    dr=0
    if columns and len(columns) == df.shape[1]:
        df.columns = columns
        dr+=1
        
    # Check for rows where all columns have the same value 
    for index, row in df.iterrows():
        if len(set(row)) == 1:  # Check if all values in the row are the same
            df.iloc[index, 1:] = [None] * (len(df.columns) - 1)
    
    # Check for consecutive columns with the same value
    for col in range(len(df.columns) - 1):
        for index, row in df.iterrows():
            if row[col] != 'P.O.A' and row[col]!='P.0.A' and row[col]!='P.0.A.':
                if row[col] is not None and (not row[col].replace('$', '').replace('.', '', 1).isdigit()):
                    if col + 1 < len(row):
                        if row[col] == row[col + 1]:
                            df.iloc[index, col + 1] = None
                    if col + 2 < len(row):
                        if row[col] == row[col + 2]:
                            df.iloc[index, col + 2] = None
                    if col + 3 < len(row):
                        if row[col] == row[col + 3]:
                            df.iloc[index, col + 3] = None
                    if col + 4 < len(row):
                        if row[col] == row[col + 4]:
                            df.iloc[index, col + 4] = None
             
    
    if dr!=0:
        df.drop(index=0, inplace=True)  # Drop the first row
    
    with pd.ExcelWriter(filename) as writer:
        df.to_excel(writer, index=False)
    print(f"\033[92m[INFO] XLSX file created: {filename}\033[0m")
        
def Image_to_xlsx(src,output_folder):
    
    print("\033[94m[INFO] Processing image folder...\033[0m")
    file_name = src
    src = f'./images/{src}'
    # Check if src is None
    if src is None:
        print("Error: src is None")
        exit()
    if not os.path.exists(src) or not os.path.isfile(src):
        print("\033[91m[ERROR] Invalid file: {src}\033[0m")
        exit()

    # Instantiation of OCR
    oc = TesseractOCR(n_threads=4, lang="eng", quiet=True)  # Set quiet=True to suppress unwanted output

    # Instantiation of document, either an image or a PDF
    doc = Image(src)

    # Table extraction
    extracted_tables = doc.extract_tables(ocr=oc,
                                          implicit_rows=False,
                                          borderless_tables=False,
                                          min_confidence=50)
    table_count = 1
    for table in extracted_tables:
        df = table.df
        # print(df.to_string())

        # Check if the first column is consistent, if so, use it as column names
        if (df[0][0] == df[1][0] != None):
            if '\n' in df[0][0]:
                columns = (df[0][0]).split('\n')
            else:
                columns = [df[0][0]] + [None] * (len(df.columns) - 1)
        else:
            columns = False

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        filename_without_extension = os.path.splitext(file_name)[0]
        output_filename = f"./{output_folder}/{filename_without_extension}_table_{table_count}.xlsx"
        df_to_xlsx(table.df, columns, output_filename)
        print(f"\033[92m[INFO] Table {table_count} saved to: {output_filename}\033[0m")
        table_count += 1

    print("\033[92m[INFO] processing completed!\033[0m")
    
if __name__ == "__main__":
    # variables
    image_folder = 'images'
    output_folder='output'
    input_file='./Input/Gallery Specialties 2023 Pricelist.pdf'
    
    
    pdf_to_Img(image_folder,input_file)
    
    for filename in os.listdir(image_folder):
        if filename.endswith(".png") or filename.endswith(".jpg") or filename.endswith(".jpeg"):
            file_location = os.path.join(image_folder, filename)
            print(f"\033[94m[INFO] Processing image: {filename}\033[0m")
            Image_to_xlsx(filename,output_folder)
    if os.path.exists(image_folder):
        shutil.rmtree(image_folder)
    print("\033[92m[INFO] Script execution completed successfully!\033[0m")
    



