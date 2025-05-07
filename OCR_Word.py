from pdf2image import convert_from_path
import pytesseract
from docx import Document
from PIL import Image

# Set Tesseract path and language
pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'  # Adjust path if needed
lang = 'fas'  # Persian

def extract_persian_ocr_to_word(pdf_path, output_docx_path):
    images = convert_from_path(pdf_path)
    doc = Document()

    for i, image in enumerate(images, start=1):
        text = pytesseract.image_to_string(image, lang=lang)
        doc.add_paragraph(f"\n=== PAGE {i} ===")
        doc.add_paragraph(text.strip())

    doc.save(output_docx_path)
    print(f"OCR Persian text saved to {output_docx_path}")

# Example usage
extract_persian_ocr_to_word("file.pdf", "extracted_text_ocr.docx")
