'''
Author: Danny Huang
Date: 27 June 2023

Purpose: to extract texts from excel file, and use them to annotate pdf files located in given folder.
system reqirements: python 3.8.5, annaconda build
module requirements: os, pandas, and pymupdf(inside anaconda prompt, run 'pip install pymupdf')

Use Instructioin:
    1) index.xls to contain all the pdf files
    2) .\raw folder to contain all the pdf files extracted from CDS system (Empower, OpenLab)
    3) .\output folder to contain the final output.
'''
import os
import pandas as pd
import fitz

# initialisation of file locations

pdf_location = r".\\raw\\" # use '\' in windows system.
pdf_annotated = r".\\annotated\\"

index_location = r".\\Index.xlsx"
cover_page = r".\\cover_page.pdf"
output_file = r".\\output\\Annotated_and_Merged.pdf"

# end of initialisation

df_index = pd.read_excel(index_location, sheet_name = "Index", na_values = ['NA']) # read the excel file into pandas dataframe
series_index = df_index['Index'] # slice the dataframe into dataframe series
series_annotation = df_index['Annotations']
# now that i have obtained to series, one for pdf file names, and one for annotations

file_names = series_index.tolist()
file_annotations = series_annotation.tolist()
file_names_full = [] # create an empty list to host file location + file name on harddisk

# populate the empty list with file location + file name
for name in file_names: # to obtain a list of file name (e.g., 101.pdf)
    file_names_full.append(str(name) + '.pdf')

# Annotate files
for (file_name,annotation) in zip(file_names_full, file_annotations):
    for pdf_file in os.listdir(pdf_location):
        if (file_name == pdf_file):
            # found the file
            # open file for annotation
            pdf_file_full_path = pdf_location + pdf_file
            
            pdf_annotated_path = pdf_annotated + pdf_file

            # Open file for annotation
            doc = fitz.open(pdf_file_full_path)
            # perform conversion to ensure pdf versions are all aligned to pymupdf
            pdfbytes = doc.convert_to_pdf()
            doc =  fitz.open("pdf", pdfbytes)

            red = (1,0,0) # define color = red for use in text 
            x2 = fitz.get_text_length(annotation, fontsize = 7, fontname = 'china-s')
            r1 = fitz.Rect(110,13,x2+111,25)
            r2 = fitz.Rect(108,12,x2+110,26)
            # note: annotation can not contain special chars like +/- signs
            for page in doc:
                page.draw_rect(r2, color = (1,0,0))
                annot = page.insert_textbox(r1,annotation,fontsize=7,border_width=1, color = red, fontname = 'china-s')
            doc.save(pdf_annotated_path)
            doc.close() # close fitz object

# do merging
doc_index = fitz.open(cover_page)

# merging
file_list = os.listdir(pdf_annotated)
file_list.sort() # sort in ascending order

for pdf in file_list:
    
    pdf_with_path = pdf_annotated + pdf
    doc_pdf = fitz.open(pdf_with_path)
    doc_index.insert_pdf(doc_pdf)

doc_index.save(output_file)
doc_index.close()
print("file annotation successfully comleted.")

