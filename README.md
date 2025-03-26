# PDF Table Extractor

A tool to detect and extract tables from system-generated PDFs and store them in an Excel sheet without using Tabula, Camelot, or image conversion.

## 🚀 Features

- Extracts text-based tables from PDFs
- Handles tables with or without borders and irregular shapes
- Cleans extracted text and maintains structured formatting
- Saves extracted tables as an Excel file

## 📌 Technologies Used

- `PyMuPDF (fitz)` for PDF text extraction
- `pandas` for data handling
- `openpyxl` for Excel file creation
- `re` for text processing and cleanup

## 🛠 Installation

Make sure you have Python installed. Then install dependencies:

```bash
pip install pymupdf pandas openpyxl
```

## 📂 Usage

Run the script with a PDF file:

```python
python script.py
```

Modify the script to process your desired PDF:

```python
pdf_file = "test3.pdf"
excel_file = "output3.xlsx"

processingpdf(pdf_file, excel_file)
```

## 📝 How It Works

1. **Extract Text** – Reads and extracts text from the PDF.
2. **Format Data** – Cleans and structures text into a table format.
3. **Export to Excel** – Saves structured data to an Excel file.

## 🔍 Constraints

- Does **not** use Tabula or Camelot.
- Does **not** convert PDFs to images.
- Handles various table formats, including irregular shapes.

## ⚡ Example Output

**Input:** PDF with tables  
**Output:** Structured Excel file (`output.xlsx`)

## ⚠ Limitations & Future Improvements

- Might not handle complex merged cells perfectly.
- Can be improved for multi-line table cells.
- Will not work on the encoded pdf files

## 📷 Screenshot of the Output

Here is a preview of the extracted table:

<img src="output_example1.png" alt="Sample Output 1" width="600">

<img src="output_example2.png" alt="Sample Output 2" width="600">
