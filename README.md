# PDF to Word Converter

This project is a Python module that converts all PDF files in a specified folder to Word documents (.docx). It uses the `pdf2docx` library to ensure high-quality conversion.

## Features

- Batch conversion of PDF files to Word documents.
- Automatically preserves the original file names with a `.docx` extension.
- Easy-to-use and lightweight.

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

## Contributing

Contributions are welcome! Feel free to submit issues or pull requests to improve the functionality or fix bugs.

## Acknowledgments

- Thanks to the creators of the `pdf2docx` library for making this project possible.

---

Feel free to customize the `README.md` file further based on your needs or preferences!

