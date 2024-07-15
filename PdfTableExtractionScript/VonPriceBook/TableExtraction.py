from img2table.ocr import TesseractOCR
from img2table.document import Image
import os
import pandas as pd

def split_finish_column(df):
    # Create a new dataframe to store the result
    new_df = pd.DataFrame(columns=df.columns)

    # Iterate over each row in the original dataframe
    for index, row in df.iterrows():
        # Split the "Finish" column into separate values
        finish_values = [x.strip() for x in row['Finish'].split(',')]
        
        # Create a new row for each finish value
        for finish_value in finish_values:
            new_row = row.copy()
            new_row['Finish'] = finish_value
            new_row_df = pd.DataFrame([new_row])
            new_df = pd.concat([new_df, new_row_df], ignore_index=True)

    return new_df

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
        if len(df) > 1 and (df.iloc[0, 0] == df.iloc[1, 0] and df.iloc[0, 0] is not None):
            if '\n' in df[0][0]:
                columns = (df[0][0]).split('\n')
            else:
                columns = [df[0][0]] + [None] * (len(df.columns) - 1)
        else:
            columns = False

        if columns:
            df.columns = columns

        # Check if there's a column named "Finish"
        if "Finish" in df.columns:
            df = split_finish_column(df)

        filename_without_extension = os.path.splitext(file_name)[0]
        output_filename = f"./output/{filename_without_extension}_table_{table_count}.xlsx"
        data_to_xlsx(df, columns, output_filename)
        print(f"Table {table_count} saved to: {output_filename}")
        table_count += 1

    print("Script Success !")

if __name__ == "__main__":
    image_folder = 'images'

    # for filename in os.listdir(image_folder):
    #     if filename.endswith(".png") or filename.endswith(".jpg") or filename.endswith(".jpeg"):
    #         file_location = os.path.join(image_folder, filename)
    #         print(f"Processing image: {filename}")
    ImagetoXls('page_26.png')

    print("Script Success!")