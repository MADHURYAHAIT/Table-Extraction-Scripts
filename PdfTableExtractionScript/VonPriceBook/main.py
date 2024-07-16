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

def pdf_to_img_folder(pdf_location, img_folder):
    print("\033[92m[INFO] Converting PDF to images...\033[0m")
    if not os.path.exists(img_folder):
        os.makedirs(img_folder)

    try:
        pdf_doc = fitz.open(pdf_location)
        print(f"\033[92m[INFO] Opened PDF file: {pdf_location}\033[0m")
    except Exception as e:
        print(f"\033[91m[ERROR] Unable to open PDF file: {e}\033[0m")
        return

    for page_num in range(len(pdf_doc)):
        page = pdf_doc[page_num]
        print(f"\033[92m[INFO] Processing page {page_num+1} of {len(pdf_doc)}\033[0m")
        zoom_matrix = fitz.Matrix(4, 4)
        try:
            page_img = page.get_pixmap(matrix=zoom_matrix)
            print(f"\033[92m[INFO] Rendered page {page_num+1} to image\033[0m")
        except Exception as e:
            print(f"\033[91m[ERROR] Unable to render page to image: {e}\033[0m")
            continue

        img_path = os.path.join(img_folder, f'page_{page_num+1}.png')
        try:
            page_img.save(img_path)
            print(f"\033[92m[INFO] Saved image to: {img_path}\033[0m")
        except Exception as e:
            print(f"\033[91m[ERROR] Unable to save image: {e}\033[0m")
            continue

    pdf_doc.close()
    print("\033[92m[INFO] Image conversion completed!\033[0m")

def split_finish_column(df):
    print("Splitting FINISH column...")
    new_df = pd.DataFrame(columns=df.columns)

    for index, row in df.iterrows():
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
    
    try:
        with pd.ExcelWriter(filename) as writer:
            df.to_excel(writer, index=False)
        print(f"\033[92m[INFO] XLSX file created: {filename}\033[0m")
    except Exception as e:
        print(f"\033[91m[ERROR] Unable to create XLSX file: {e}\033[0m")
        return

    


def img_to_xlsx(img_path, output_folder):
    print("\033[94m[INFO] Processing image...\033[0m")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    img_name = os.path.basename(img_path)
    output_filename = os.path.join(output_folder, f"{os.path.splitext(img_name)[0]}.xlsx")

    null_device = open(os.devnull, 'w')
    original_stdout = sys.stdout
    sys.stdout = null_device

    try:
        ocr_instance = TesseractOCR(n_threads=4, lang="eng")
        print(f"\033[92m[INFO] Instantiated OCR with {ocr_instance.n_threads} threads and language {ocr_instance.lang}\033[0m")
    except Exception as e:
        print(f"\033[91m[ERROR] Unable to instantiate OCR: {e}\033[0m")
        sys.stdout = original_stdout
        null_device.close()
        return

    try:
        img_doc = Image(img_path)
        print(f"\033[92m[INFO] Instantiated document from image: {img_path}\033[0m")
    except Exception as e:
        print(f"\033[91m[ERROR] Unable to instantiate document: {e}\033[0m")
        sys.stdout = original_stdout
        null_device.close()
        return

    try:
        extracted_tables = img_doc.extract_tables(ocr=ocr_instance,
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

        if len(df) > 1 and df.iloc[0, 0] == df.iloc[1, 0] and df.iloc[0, 0] is not None:
            if '\n' in df[0][0]:
                columns = (df[0][0]).split('\n')
            else:
                columns = [df[0][0]] + [None] * (len(df.columns) - 1)
        else:
            columns = False

        try:
            df_to_xlsx(df, columns, output_filename)
        except Exception as e:
            print(f"\033[91m[ERROR] Unable to create XLSX file: {e}\033[0m")
            continue

        table_count += 1

    sys.stdout = original_stdout
    null_device.close()

    print("\033[92m[INFO] Image processing completed!\033[0m")

if __name__ == "__main__":
    img_folder = 'images'
    output_folder='output'
    pdf_location='./Input/Von Price Book Feb 13 to Jun 23.pdf'

    print("\033[94m[INFO] Starting script execution...\033[0m")
    start_time = time.time()

    pdf_to_img_folder(pdf_location, img_folder)
    for img_name in os.listdir(img_folder):
        if img_name.endswith(".png") or img_name.endswith(".jpg") or img_name.endswith(".jpeg"):
            img_path = os.path.join(img_folder, img_name)
            print(f"\033[94m[INFO] Processing image: {img_name}\033[0m")
            img_to_xlsx(img_path, output_folder)

    if os.path.exists(img_folder):
        shutil.rmtree(img_folder)

    end_time = time.time()
    print(f"\033[92m[INFO] Script execution completed successfully in {end_time - start_time:.2f} seconds!\033[0m")