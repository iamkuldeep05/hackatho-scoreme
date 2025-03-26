import fitz  # PyMuPDF
import pandas as pd
from openpyxl import Workbook
import re

def processingpdf(pdf_file, excel_file):
    # Handles text extraction, cleaning, and saving to an Excel file
    content = extract_text(pdf_file)
    structured_data = format_as_table(content)
    saving_excel(structured_data, excel_file)
    print(f"Extracted data saved to: {excel_file}")

def extract_text(pdf_file):
    #Extracts raw text from a PDF
    doc = fitz.open(pdf_file)
    all_text = []

    for page in doc:
        text = page.get_text("text")  # Extract text from page
        lines = text.split("\n")  # Break text into lines
        all_text.extend(lines)

    return all_text

def format_as_table(lines):
    #Formats extracted lines into structured table format.
    structured_data = []

    for line in lines:
        clean_line = remove_special_chars(line)
        row = clean_line.split("\t")  # Columns are separated by tab space
        if any(row):  # Skip empty rows
            structured_data.append(row)

    return structured_data

def remove_special_chars(text):
    #Removes unwanted symbols & non-printable characters from text.
    text = text.replace(":", "")  # Remove colons
    text = "".join(char if 32 <= ord(char) <= 126 else " " for char in text)  # Remove non-printable chars
    return re.sub(r'\s{2,}', '\t', text.strip())  # Normalize spaces

def saving_excel(data, excel_file):
    #Saves structured data into an Excel file.
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Extracted Data"

    for row in data:
        sheet.append(row)

    wb.save(excel_file)

# Output usage
pdf_file = "test3.pdf"
excel_file = "output3.xlsx"

processingpdf(pdf_file, excel_file)
