import os
from transformers import DetrImageProcessor, TableTransformerForObjectDetection
from PIL import Image
from pdf2image import convert_from_path
import pdfplumber
from docx import Document
import cv2
import numpy as np
import torch


def preprocess_image(image_path):
    image = cv2.imread(image_path)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
    processed_image_path = "processed_image.jpg"
    cv2.imwrite(processed_image_path, binary)
    return processed_image_path

def extract_tables(image_path, processor, model):
    image = Image.open(image_path)
    inputs = processor(images=image, return_tensors="pt")
    outputs = model(**inputs)

    # Decode bounding boxes and labels
    target_sizes = torch.tensor([image.size[::-1]])
    results = processor.post_process_object_detection(outputs, target_sizes=target_sizes, threshold=0.9)

    tables = []
    for result in results:
        for box, label in zip(result["boxes"], result["labels"]):
            if model.config.id2label[label.item()] == "table":
                x_min, y_min, x_max, y_max = map(int, box.tolist())
                table_crop = image.crop((x_min, y_min, x_max, y_max))
                table_image_path = f"table_{len(tables)}.jpg"
                table_crop.save(table_image_path)
                tables.append(table_image_path)
    return tables

def extract_text(pdf_path):
    text_content = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text_content.append(page.extract_text())
    return "\n".join(text_content)

def create_word_document(text, tables, output_path):
    doc = Document()
    doc.add_paragraph(text)  # Add extracted text

    for table_image_path in tables:
        doc.add_paragraph("Extracted Table:")
        table = doc.add_table(rows=1, cols=1)
        cell = table.cell(0, 0)
        cell.text = "Table image attached below."
        run = cell.paragraphs[0].add_run()
        run.add_picture(table_image_path, width=docx.shared.Inches(6))
    
    doc.save(output_path)


def process_pdf(pdf_path, output_word_path):
    # Initialize Table Transformer
    processor = DetrImageProcessor.from_pretrained("microsoft/table-transformer-detection")
    model = TableTransformerForObjectDetection.from_pretrained("microsoft/table-transformer-detection")

    # Convert PDF to images
    images = convert_from_path(pdf_path, dpi=300, fmt='jpeg')
    os.makedirs("temp_images", exist_ok=True)
    image_paths = []
    for i, image in enumerate(images):
        image_path = f"temp_images/page_{i + 1}.jpg"
        image.save(image_path)
        image_paths.append(image_path)

    # Process each page for tables and text
    all_text = extract_text(pdf_path)
    all_tables = []
    for image_path in image_paths:
        processed_image = preprocess_image(image_path)
        tables = extract_tables(processed_image, processor, model)
        all_tables.extend(tables)

    # Create Word document
    create_word_document(all_text, all_tables, output_word_path)

    # Cleanup
    for path in image_paths + all_tables:
        os.remove(path)
    os.rmdir("temp_images")


if __name__ == "__main__":
    # Input PDF and output Word paths
    pdf_file = "docs\Original.pdf"
    output_word_file = "output.docx"

    process_pdf(pdf_file, output_word_file)
    print(f"Word document saved to {output_word_file}")