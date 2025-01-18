import os
from pdf2docx import Converter

def convert_pdfs_to_docx(input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)

    for file_name in os.listdir(input_folder):
        if file_name.endswith('.pdf'):
            pdf_file = os.path.join(input_folder, file_name)
            docx_file = os.path.join(output_folder, file_name.replace('.pdf', '.docx'))

            print(f'Converting: {file_name} to {docx_file}')

            cv = Converter(pdf_file)
            cv.convert(docx_file)
            cv.close()

    print("Success!")


if __name__ == "__main__":
    input_folder = 'docs/'  
    output_folder = 'output/' 
    convert_pdfs_to_docx(input_folder, output_folder) 