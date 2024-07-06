from img2table.ocr import TesseractOCR
from img2table.document import Image
import os
import pandas as pd

def data_to_xlsx(df, columns, filename):
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
            if row[col] != 'P.O.A' or row[col]!='P.0.A' or row[col]!='P.0.A.':
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
        
def ImagetoXls(src):
    file_name = src
    src = f'./images/{src}'
    # Check if src is None
    if src is None:
        print("Error: src is None")
        exit()

    if not os.path.exists(src) or not os.path.isfile(src):
        print("Error: src is not a valid file")
        exit()

    # Instantiation of OCR
    oc = TesseractOCR(n_threads=4, lang="eng")

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
        print(df.to_string())

        # Check if the first column is consistent, if so, use it as column names
        if (df[0][0] == df[1][0] != None):
            if '\n' in df[0][0]:
                columns = (df[0][0]).split('\n')
            else:
                columns = [df[0][0]] + [None] * (len(df.columns) - 1)
        else:
            columns = False

        filename_without_extension = os.path.splitext(file_name)[0]
        output_filename = f"./output/{filename_without_extension}_table_{table_count}.xlsx"
        data_to_xlsx(table.df, columns, output_filename)
        print(f"Table {table_count} saved to: {output_filename}")
        table_count += 1

    print("Script Success !")


if __name__ == "__main__":
    image_folder = 'images'

    for filename in os.listdir(image_folder):
        if filename.endswith(".png") or filename.endswith(".jpg") or filename.endswith(".jpeg"):
            file_location = os.path.join(image_folder, filename)
            print(f"Processing image: {filename}")
            ImagetoXls(filename)

    print("Script Success!")