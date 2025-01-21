# PDF-to-DOCX Conversion Tool

## Overview
This project provides a robust and automated solution for converting PDF files into high-quality, well-formatted DOCX documents. It is designed to handle large batches of PDFs efficiently while ensuring accurate conversion and consistent formatting.

## Features

### 1. Batch Processing of PDFs
- Automatically processes all PDF files in a specified input folder.
- Supports large-scale conversions with ease.

### 2. PDF Flattening
- Prepares layered or interactive PDFs for accurate conversion.
- Ensures compatibility across varied document types.

### 3. High-Fidelity Conversion
- Converts PDFs to DOCX using advanced layout analysis.
- Retains document structure, including tables, text, and formatting.

### 4. Formatting Adjustments
- **Margin Adjustment**: Sets uniform 0.3-inch margins on all sides of the DOCX files.
- **Table Handling**: Prevents tables from splitting across pages for better readability.

### 5. Duplicate Conversion Prevention
- Checks if a PDF has already been converted to a DOCX file.
- Skips already processed files to save time and resources.

### 6. Intermediate File Cleanup
- Removes temporary files (e.g., flattened PDFs) after processing.
- Keeps the output directory clean and organized.

### 7. User-Friendly Workflow
- Provides detailed console logs for each step, ensuring transparency.
- Clearly indicates progress and completion for each file.

### 8. Customizable Input and Output Paths
- Allows users to define the input folder for PDFs and the output folder for DOCX files.
- Automatically creates directories if they do not exist.

### 9. Robust Error Handling
- Gracefully manages missing input files or empty directories.
- Ensures stable processing even for complex PDFs.

## Installation

1. Clone the repository:
   ```bash
   git clone hhttps://github.com/RebelAKL/PDF-to-DOCX-Conversion-Tool.git
   cd PDF-to-DOCX-Conversion-Tool
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Ensure you have the following installed on your system:
   - [Poppler](https://poppler.freedesktop.org/) (for `pdf2image` library)

## Usage

1. Place your PDF files in the `docs` folder (or any folder of your choice).
2. Run the script:
   ```bash
   python pdf2docx_converter.py
   ```
3. Specify the input and output folders in the script if different from the defaults:
   - Input Folder: `docs`
   - Output Folder: `processed_output`

## File Structure
```
project-root/
├── docs/                   # Input folder for PDF files
├── processed_output/       # Output folder for converted DOCX files
├── pdf2docx_convertert.py  # Main script
├── requirements.txt        # Python dependencies
├── README.md               # Project documentation
```

## Console Logs
The script provides detailed logs for each stage:
- **Flattening PDF**: Preparing the PDF for conversion.
- **Converting PDF to DOCX**: Performing the main conversion.
- **Adjusting DOCX Formatting**: Ensuring consistent margins and table handling.
- **Cleaning Up Intermediate Files**: Removing temporary files and directories.

Example log output:
```
=== Processing example.pdf ===
=== Flattening PDF ===
Flattened PDF saved at: processed_output/example/flattened.pdf
=== Converting PDF to DOCX ===
Converted PDF to DOCX: processed_output/example.docx
=== Adjusting DOCX Formatting ===
Margins adjusted for: processed_output/example.docx
=== Cleaning Up Intermediate Files ===
Removed folder: processed_output/example
=== Completed Processing for: example.pdf ===
```

## Requirements
- Python 3.8+
- Dependencies listed in `requirements.txt`:
  - `pdf2image`
  - `PyPDF2`
  - `pytesseract`
  - `pdf2docx`
  - `python-docx`
  - `Pillow`

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.