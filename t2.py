import os
from PyPDF2 import PdfReader, PdfWriter
from pdf2docx import Converter
from docx import Document
from docx.shared import Inches
import re

def flatten_pdf(input_pdf, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    with open(output_pdf, "wb") as f:
        writer.write(f)
    print(f"Flattened PDF saved at: {output_pdf}")

def convert_pdf_to_docx(flattened_pdf, output_docx):
    cv = Converter(flattened_pdf)
    cv.convert(output_docx, start=0, end=None, layout_analysis=True)
    cv.close()
    print(f"Converted PDF to DOCX: {output_docx}")

def format_index_page(docx_file):
    doc = Document(docx_file)
    index_regex = r'(\d+\.\s[^\d]+)\s+(\.+)\s*(\d+)'

    max_description_width = 50
    for para in doc.paragraphs:
        match = re.match(index_regex, para.text)
        if match:
        
            description = match.group(1).strip()  
            leader_dots = match.group(2)  
            page_number = match.group(3).strip() 

            formatted_entry = f"{description.ljust(max_description_width)} {leader_dots} {page_number.rjust(3)}"

            para.text = formatted_entry

    doc.save(docx_file)
    print(f"Formatted index page for: {docx_file}")

def adjust_docx_formatting(docx_file):
    doc = Document(docx_file)

    for section in doc.sections:
        section.top_margin = Inches(0.3)
        section.bottom_margin = Inches(0.3)
        section.left_margin = Inches(0.3)
        section.right_margin = Inches(0.3)

    for table in doc.tables:
        for row in table.rows:
            row.allow_break_across_pages = True

    format_index_page(docx_file)

    doc.save(docx_file)
    print(f"Formatting adjusted for: {docx_file}")

def process_pdfs_in_folder(input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)

    for file_name in os.listdir(input_folder):
        if file_name.endswith(".pdf"):
            input_pdf = os.path.join(input_folder, file_name)
            output_subfolder = os.path.join(output_folder, os.path.splitext(file_name)[0])
            os.makedirs(output_subfolder, exist_ok=True)

            flattened_pdf = os.path.join(output_subfolder, "flattened.pdf")
            docx_output = os.path.join(output_folder, f"{os.path.splitext(file_name)[0]}.docx")

            print(f"\n=== Processing {file_name} ===")

            print("\n=== Flattening PDF ===")
            flatten_pdf(input_pdf, flattened_pdf)

            print("\n=== Converting Flattened PDF to DOCX ===")
            convert_pdf_to_docx(flattened_pdf, docx_output)

            print("\n=== Adjusting DOCX Formatting ===")
            adjust_docx_formatting(docx_output)

            print(f"\n=== Completed Processing for: {file_name} ===")

if __name__ == "__main__":
    input_folder = "doc1" 
    output_folder = "processed_output"  
    process_pdfs_in_folder(input_folder, output_folder)
