# PDF to Word Converter

This project provides a comprehensive solution for converting PDF files into well-formatted DOCX documents, focusing on automation, accuracy, and efficient file handling. Below are the high-level features:

-Batch Processing of PDFs:

Automatically processes all PDF files within a specified input folder.
Supports batch conversion, ensuring scalability for handling large datasets.
-PDF Flattening:

Prepares PDFs for conversion by flattening layered or interactive PDFs.
Ensures compatibility and accurate extraction of content.
PDF to DOCX Conversion:

Converts flattened PDFs into DOCX format with high fidelity.
Utilizes advanced layout analysis and multiprocessing to ensure precision and performance.
-Formatting Adjustments:

Margin Adjustment: Sets uniform margins for the converted DOCX files (0.3 inches on all sides).
Table Handling: Prevents tables from splitting across pages to maintain document readability.
-Duplicate Conversion Prevention:

Checks if a PDF has already been converted to DOCX.
Skips previously processed files, optimizing resource usage and avoiding redundancy.
Intermediate File Cleanup:

Removes temporary files and folders (e.g., flattened PDFs) created during processing.
Maintains a clean and organized output directory structure.
-User-Friendly Workflow:

Provides detailed console logs for each stage of the process, enabling transparency and easy debugging.
Clearly indicates the start, progress, and completion of tasks for each PDF file.
-Customizable Input and Output Paths:

Flexible configuration of input and output folders to adapt to different workflows.
Automatically creates output directories if they do not exist.
-Error Handling and Robustness:

Ensures stable processing even for complex PDFs with varied layouts.
Gracefully handles scenarios such as missing input files or empty folders.

## Requirements

- Python 3.6 or higher
- `pdf2docx` library

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/pdf-to-word-converter.git
   cd pdf-to-word-converter
   ```
2. Install the required dependencies:
   ```bash
   pip install pdf2docx
   ```

## Usage

1. Place all your PDF files in the `docs/` folder (or any folder you specify).
2. Run the script:
   ```bash
   python pdf_to_word_batch.py
   ```
3. The converted Word files will be saved in the `output/` folder (or any folder you specify).

## Customization

You can specify custom input and output folders by modifying the `input_folder` and `output_folder` variables in the script:

```python
if __name__ == "__main__":
    input_folder = 'your_input_folder/'
    output_folder = 'your_output_folder/'
    convert_pdfs_to_docx(input_folder, output_folder)
```

## Example Output

For a PDF file named `example.pdf` in the `docs/` folder, the script will generate a Word file named `example.docx` in the `output/` folder.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.


