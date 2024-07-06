from img2table.ocr import TesseractOCR
from img2table.document import Image
import os

import pandas as pd
def data_to_xlsx(df,columns, filename):
        if columns and len(columns) == df.shape[1]:
            df.columns = columns
            df.drop(index=0, inplace=True)


        with pd.ExcelWriter(filename) as writer:
            df.to_excel(writer, index=False)


def ImagetoXls(src):
    file_name=src
    src = f'./images/{src}'
    # Check if src is None
    if src is None:
        print("Error: src is None")
        exit()


    if not os.path.exists(src) or not os.path.isfile(src):
        print("Error: src is not a valid file")
        exit()

    # Instantiation of OCR
    oc= TesseractOCR(n_threads=4, lang="eng")

    # Instantiation of document, either an image or a PDF
    doc = Image(src)

    # Table extraction
    extracted_tables = doc.extract_tables(ocr=oc,
                                        implicit_rows=False,
                                        borderless_tables=False,
                                        min_confidence=50)
    for table in extracted_tables:
        df=table.df
        print(df)

        if (df[0][0]==df[1][0]!=None):
            columns=(df[0][0]).split('\n')
        
        else:
            columns=False    
        filename_without_extension = os.path.splitext(file_name)[0]
        data_to_xlsx(table.df,columns, f"./output/{filename_without_extension}.xlsx")
        print("Script Success !")



if __name__ == "__main__":
    image_folder = 'images'

    for filename in os.listdir(image_folder):
        if filename.endswith(".png") or filename.endswith(".jpg") or filename.endswith(".jpeg"):
            file_location = os.path.join(image_folder, filename)
            print(f"Processing image: {filename}")
            
            
            ImagetoXls(filename)
           
            
            # print(f"Output saved to: {output_filename}")
            
    print("Script Success!")