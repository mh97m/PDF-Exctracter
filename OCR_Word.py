from pdf2image import convert_from_path
import pytesseract
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from PIL import Image
import re

# Set Tesseract path and language
pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'  # Adjust path if needed
lang = 'fas'  # Persian language

def remove_parentheses_content(text):
    # Remove everything inside parentheses (including the parentheses themselves)
    return re.sub(r'\(.*?\)', '', text)

def set_font_b_nazanin(paragraph):
    if paragraph.runs:
        run = paragraph.runs[0]
        run.font.name = 'B Nazanin'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'B Nazanin')
        run.font.size = Pt(14)  # Optional: Set font size

def extract_persian_ocr_to_word(pdf_path, output_docx_path):
    images = convert_from_path(pdf_path)
    doc = Document()

    for i, image in enumerate(images, start=1):
        # OCR image to string
        text = pytesseract.image_to_string(image, lang=lang)
        # Remove parentheses and strip whitespace
        text = remove_parentheses_content(text).strip()

        # Add page title
        title_paragraph = doc.add_paragraph(f"\n=== PAGE {i} ===")
        set_font_b_nazanin(title_paragraph)

        # Add OCR content if not empty
        if text:
            content_paragraph = doc.add_paragraph(text)
            set_font_b_nazanin(content_paragraph)

    # Save the Word document
    doc.save(output_docx_path)
    print(f"OCR Persian text saved to {output_docx_path}")

# Example usage
extract_persian_ocr_to_word("file2.pdf", "extracted_text_ocr2.docx")
