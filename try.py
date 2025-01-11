import os
from PIL import Image
import fitz  # PyMuPDF
import pytesseract
from transformers import TableTransformerForObjectDetection, DetrImageProcessor
from docx import Document
from docx.shared import Pt


pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # Update path to Tesseract
 
model_name = "microsoft/table-transformer-detection"
processor = DetrImageProcessor.from_pretrained(model_name)
model = TableTransformerForObjectDetection.from_pretrained(model_name)

   
def pdf_to_images(pdf_path):
    pdf_document = fitz.open(pdf_path)
    images = []
    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        pix = page.get_pixmap(dpi=300)  
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    return images

def extract_table_cells(image):
    inputs = processor(images=image, return_tensors="pt")
    outputs = model(**inputs)
    results = processor.post_process_object_detection(outputs, threshold=0.5, target_sizes=[image.size[::-1]])

    tables = []
    for result in results:
        table_cells = []
        for box, label in zip(result["boxes"], result["labels"]):
            if label == 1:  # Table cell
                x0, y0, x1, y1 = map(int, box.tolist())
                table_cells.append((x0, y0, x1, y1))
        tables.append(table_cells)
    return tables


def get_table_dimensions(cells):
    rows, cols = set(), set()
    for cell in cells:
        x0, y0, x1, y1 = cell
        rows.add(y0)
        rows.add(y1)
        cols.add(x0)
        cols.add(x1)
    return len(sorted(rows)) // 2, len(sorted(cols)) // 2


def perform_ocr(image, box):
    cropped_image = image.crop(box)
    return pytesseract.image_to_string(cropped_image, config="--psm 6")


def create_word_table(doc, table_cells, image):
    rows, cols = get_table_dimensions(table_cells)
    table = doc.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"

    for cell in table_cells:
        x0, y0, x1, y1 = cell
        cropped_image = image.crop((x0, y0, x1, y1))
        text = perform_ocr(cropped_image, cell)
        row_idx = sorted({box[1] for box in table_cells}).index(y0)
        col_idx = sorted({box[0] for box in table_cells}).index(x0)
        table.cell(row_idx, col_idx).text = text.strip()

    return doc


def map_non_table_content(doc, image, table_bboxes):
    text_outside_tables = pytesseract.image_to_string(image, config="--psm 6")
    for bbox in table_bboxes:
        x0, y0, x1, y1 = bbox
        cropped_image = image.crop((x0, y0, x1, y1))
        table_text = pytesseract.image_to_string(cropped_image)
        text_outside_tables = text_outside_tables.replace(table_text, "")  
    doc.add_paragraph(text_outside_tables.strip())
    return doc


def process_pdf(pdf_path, output_word_path):
    images = pdf_to_images(pdf_path)
    doc = Document()

    for page_num, image in enumerate(images):
        doc.add_paragraph(f"Page {page_num+1}").bold = True
        tables = extract_table_cells(image)

        table_bboxes = [(min(cell[0] for cell in table), min(cell[1] for cell in table),
                         max(cell[2] for cell in table), max(cell[3] for cell in table))
                        for table in tables]
        doc = map_non_table_content(doc, image, table_bboxes)


        for table_cells in tables:
            doc = create_word_table(doc, table_cells, image)

    doc.save(output_word_path)



if __name__ == "__main__":
    pdf_file = "docs\Original.pdf"  
    output_word_file = "output.docx"
    process_pdf(pdf_file, output_word_file)
