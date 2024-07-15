import fitz
import os
from img2table.ocr import TesseractOCR
from img2table.document import Image
import pandas as pd
import shutil
import sys
import time

import warnings
warnings.simplefilter(action='ignore', category=Warning)


# Define a function to convert PDF to images
def pdf_to_Img(foldername, location):
    print("\033[92m[INFO] Starting image conversion...\033[0m")
    image_folder = foldername
    if not os.path.exists(image_folder):
        os.makedirs(image_folder)

    # Open the PDF file
    try:
        doc = fitz.open(location)
        print(f"\033[92m[INFO] Opened PDF file: {location}\033[0m")
    except Exception as e:
        print(f"\033[91m[ERROR] Unable to open PDF file: {e}\033[0m")
        return

    # Iterate over each page
    for page_num in range(len(doc)):
        page = doc[page_num]
        print(f"\033[92m[INFO] Processing page {page_num+1} of {len(doc)}\033[0m")
        # Set the matrix to scale up the image to 400 dpi
        zoom_matrix = fitz.Matrix(4, 4)  # 4x4 matrix, scales up by 3 in both x and y directions
        # Render the page to an image
        try:
            image = page.get_pixmap(matrix=zoom_matrix)
            print(f"\033[92m[INFO] Rendered page {page_num+1} to image\033[0m")
        except Exception as e:
            print(f"\033[91m[ERROR] Unable to render page to image: {e}\033[0m")
            continue

        # Save the image to a file in the image folder
        try:
            image.save(os.path.join(image_folder, f'page_{page_num+1}.png'))
            print(f"\033[92m[INFO] Saved image to: {os.path.join(image_folder, f'page_{page_num+1}.png')}\033[0m")
        except Exception as e:
            print(f"\033[91m[ERROR] Unable to save image: {e}\033[0m")
            continue

    doc.close()
    print("\033[92m[INFO] Image conversion completed!\033[0m")

# Splits the column with group values.
def split_finish_column(df):
    print("Splitting FINISH column...")
    new_df = pd.DataFrame(columns=df.columns)

    for index, row in df.iterrows():
        # Check if the "FINISH" column value is not None
        if pd.notna(row['FINISH']):
            finish_values = [x.strip() for x in row['FINISH'].split(',') if x.strip()]
            for finish_value in finish_values:
                new_row = row.copy()
                new_row['FINISH'] = finish_value
                new_row_df = pd.DataFrame([new_row])
                new_df = pd.concat([new_df, new_row_df], ignore_index=True)
        else:
            new_row_df = pd.DataFrame([row])
            new_df = pd.concat([new_df, new_row_df], ignore_index=True)

    return new_df


# Define a function to convert a DataFrame to an XLSX file
def df_to_xlsx(df, columns, filename):
    print("\033[92m[INFO] Starting XLSX conversion...\033[0m")
    dr=0
    if columns and len(columns) == df.shape[1]:
        df.columns = columns
        dr+=1
        
    if "FINISH" in df.columns:
        df = split_finish_column(df)
        
    # Check for rows where all columns have the same value 
    for index, row in df.iterrows():
        if len(set(row)) == 1:  # Check if all values in the row are the same
            df.iloc[index, 1:] = [None] * (len(df.columns) - 1)
    
    # Check for consecutive columns with the same value
    for col in range(len(df.columns) - 1):
        for index, row in df.iterrows():
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
    try:
        with pd.ExcelWriter(filename) as writer:
            df.to_excel(writer, index=False)
        print(f"\033[92m[INFO] XLSX file created: {filename}\033[0m")
    except Exception as e:
        print(f"\033[91m[ERROR] Unable to create XLSX file: {e}\033[0m")
        return

        
# Define a function to convert an image to an XLSX file
def Image_to_xlsx(src, output_folder):
    print("\033[94m[INFO] Accessing image folder...\033[0m")
    file_name = src
    src = f'./images/{src}'
    # Check if src is None
    if src is None:
        print("Error: src is None")
        exit()
    if not os.path.exists(src) or not os.path.isfile(src):
        print("\033[91m[ERROR] Invalid file: {src}\033[0m")
        exit()

    null_device = open(os.devnull, 'w')
    original_stdout = sys.stdout
    sys.stdout = null_device

    # Instantiation of OCR
    try:
        oc =TesseractOCR(n_threads=4, lang="eng")
        print(f"\033[92m[INFO] Instantiated OCR with {oc.n_threads} threads and language {oc.lang}\033[0m")
    except Exception as e:
        print(f"\033[91m[ERROR] Unable to instantiate OCR: {e}\033[0m")
        sys.stdout = original_stdout
        null_device.close()
        return

    # Instantiation of document, either an image or a PDF
    try:
        doc = Image(src)
        print(f"\033[92m[INFO] Instantiated document from image: {src}\033[0m")
    except Exception as e:
        print(f"\033[91m[ERROR] Unable to instantiate document: {e}\033[0m")
        sys.stdout = original_stdout
        null_device.close()
        return

    # Table extraction
    try:
        extracted_tables = doc.extract_tables(ocr=oc,
                                              implicit_rows=False,
                                              borderless_tables=False,
                                              min_confidence=50)
        print(f"\033[92m[INFO] Extracted {len(extracted_tables)} tables from image\033[0m")
    except Exception as e:
        print(f"\033[91m[ERROR] Unable to extract tables: {e}\033[0m")
        sys.stdout = original_stdout
        null_device.close()
        return

    table_count = 1
    for table in extracted_tables:
        df = table.df
        print(f"\033[92m[INFO] Processing table {table_count} of {len(extracted_tables)}\033[0m")
        # Call split_finish_column for each table
        # Check if the first column is consistent, if so, use it as column names
        if len(df) > 1 and df.iloc[0, 0] == df.iloc[1, 0] and df.iloc[0, 0] is not None:
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
        print(f"\033[92m[INFO] Saving table {table_count} to: {output_filename}\033[0m")
        try:
            df_to_xlsx(table.df, columns, output_filename)
        except Exception as e:
            print(f"\033[91m[ERROR] Unable to create XLSX file: {e}\033[0m")
            continue

        table_count += 1

    sys.stdout = original_stdout
    null_device.close()

    print("\033[92m[INFO] Image processing completed!\033[0m")

    
    
if __name__ == "__main__":
    # variables
    image_folder = 'images'
    output_folder='output'
    input_file='./Input/Von Price Book Feb 13 to Jun 23.pdf'

    print("\033[94m[INFO] Starting script execution...\033[0m")
    start_time = time.time()

    pdf_to_Img(image_folder,input_file)

    for filename in os.listdir(image_folder):
        if filename.endswith(".png") or filename.endswith(".jpg") or filename.endswith(".jpeg"):
            file_location = os.path.join(image_folder, filename)
            print(f"\033[94m[INFO] Processing image: {filename}\033[0m")
            Image_to_xlsx(filename,output_folder)

    if os.path.exists(image_folder):
        shutil.rmtree(image_folder)

    end_time = time.time()
    print(f"\033[92m[INFO] Script execution completed successfully in {end_time - start_time:.2f} seconds!\033[0m")