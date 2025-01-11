import cv2
import pytesseract
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Pt
from table_transformer.predictor import TablePredictor
import os
import io
from PIL import Image

# Initialize Table Transformer
model = TablePredictor()

# OCR Configuration
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Function to convert PDF to images
def pdf_to_images(pdf_path):
    images = convert_from_path(pdf_path, dpi=300)
    return images

# Function to extract tables and text using Table Transformer
def extract_tables_from_image(image):
    image_np = cv2.cvtColor(cv2.imread(image), cv2.COLOR_BGR2RGB)
    tables = model.predict(image_np)
    return tables

# Function to perform OCR on cropped cells
def extract_text_from_image(image):
    text = pytesseract.image_to_string(image, lang='rus+eng')
    return text.strip()

# Function to create Word document with identical layout
def create_word_document(tables, word_path):
    doc = Document()

    # Iterate through each detected table
    for table in tables:
        bbox = table['bbox']
        structure = table['structure']
        cells = structure['cells']
        num_rows = structure['nrows']
        num_cols = structure['ncols']

        # Create Word table
        word_table = doc.add_table(rows=num_rows, cols=num_cols)
        word_table.style = 'Table Grid'

        # Fill the table
        for cell in cells:
            row_idx = cell['row']
            col_idx = cell['col']
            cell_bbox = cell['bbox']
            cell_image = image.crop(cell_bbox)
            cell_text = extract_text_from_image(cell_image)

            # Add text to the Word cell
            word_table.cell(row_idx, col_idx).text = cell_text

    # Save the document
    doc.save(word_path)
    print(f"Word document saved at: {word_path}")

# Function to process PDF and generate Word document
def pdf_to_word(pdf_path, output_path):
    images = pdf_to_images(pdf_path)

    for i, image in enumerate(images):
        image_path = f"temp_page_{i + 1}.png"
        image.save(image_path)
        tables = extract_tables_from_image(image_path)

        # Save extracted content to Word
        word_path = os.path.join(output_path, f"output_page_{i + 1}.docx")
        create_word_document(tables, word_path)

        # Clean up temporary image files
        os.remove(image_path)

# Example usage
pdf_path = "your_pdf_file.pdf"
output_path = "output_directory"
os.makedirs(output_path, exist_ok=True)

pdf_to_word(pdf_path, output_path)
