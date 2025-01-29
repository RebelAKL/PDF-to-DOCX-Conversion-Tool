# cython: language_level=3
import os
import time
from PyPDF2 import PdfReader, PdfWriter
from pdf2docx import Converter
from docx import Document
from docx.shared import Inches, Pt

cdef extern from "Python.h":
    void PyEval_InitThreads()

# Optimized PDF processing functions
cdef class PDFProcessor:
    cdef str input_folder
    cdef str output_folder
    
    def __cinit__(self, str input_folder, str output_folder):
        self.input_folder = input_folder
        self.output_folder = output_folder
        PyEval_InitThreads()  # Initialize GIL for threading

    cdef void _flatten_pdf(self, str input_pdf, str output_pdf):
        cdef object reader = PdfReader(input_pdf)
        cdef object writer = PdfWriter()
        cdef object page
        for page in reader.pages:
            writer.add_page(page)
        with open(output_pdf, "wb") as f:
            writer.write(f)

    cdef void _convert_to_docx(self, str input_pdf, str output_docx):
        cdef object cv = Converter(input_pdf)
        cv.convert(output_docx, 
                  start=0, 
                  end=None,
                  layout_analysis=False,  # Disable for speed
                  multi_processing=True,  # Enable parallel processing
                  cpu_count=4)  # Adjust based on your CPU cores
        cv.close()

    cdef void _process_single(self, str file_name):
        cdef str base_name = os.path.splitext(file_name)[0]
        cdef str input_pdf = os.path.join(self.input_folder, file_name)
        cdef str temp_folder = os.path.join(self.output_folder, base_name)
        cdef str flattened_pdf = os.path.join(temp_folder, "temp.pdf")
        cdef str output_docx = os.path.join(self.output_folder, f"{base_name}.docx")

        if os.path.exists(output_docx):
            return

        # Create temp directory
        os.makedirs(temp_folder, exist_ok=True)

        try:
            # Perform core operations
            self._flatten_pdf(input_pdf, flattened_pdf)
            self._convert_to_docx(flattened_pdf, output_docx)
        finally:
            # Cleanup temp files
            for f in os.listdir(temp_folder):
                os.remove(os.path.join(temp_folder, f))
            os.rmdir(temp_folder)

    cpdef void process_all(self):
        cdef double start = time.time()
        cdef str file_name
        
        os.makedirs(self.output_folder, exist_ok=True)
        
        for file_name in os.listdir(self.input_folder):
            if file_name.endswith(".pdf"):
                self._process_single(file_name)
        
        cdef double total_time = time.time() - start
        print(f"\nTotal processing time: {total_time:.2f} seconds")

def main():
    processor = PDFProcessor("docs", "processed_output")
    processor.process_all()

if __name__ == "__main__":
    main()