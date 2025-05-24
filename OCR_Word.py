from pdf2image import convert_from_path
import pytesseract
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import nsdecls
from PIL import Image
import re

# Set Tesseract path and language
pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'  # Adjust if needed
lang = 'fas'  # Persian language

def remove_parentheses_content(text):
    return re.sub(r'\(.*?\)', '', text)

def set_paragraph_rtl(paragraph):
    """Set paragraph direction to RTL."""
    p = paragraph._element
    pPr = p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

def set_font_run(run):
    """Set fonts: Persian to B Nazanin, English to Times New Roman."""
    run.font.name = 'B Nazanin'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'B Nazanin')
    run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
    run._element.rPr.rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    run.font.size = Pt(14)

def extract_persian_ocr_to_word(pdf_path, output_docx_path):
    images = convert_from_path(pdf_path)
    doc = Document()

    for i, image in enumerate(images, start=1):
        text = pytesseract.image_to_string(image, lang=lang)
        text = remove_parentheses_content(text).strip()

        title_paragraph = doc.add_paragraph(f"\n=== PAGE {i} ===")
        set_paragraph_rtl(title_paragraph)
        if title_paragraph.runs:
            set_font_run(title_paragraph.runs[0])

        if text:
            content_paragraph = doc.add_paragraph()
            set_paragraph_rtl(content_paragraph)

            run = content_paragraph.add_run(text)
            set_font_run(run)

    doc.save(output_docx_path)
    print(f"OCR Persian text saved to {output_docx_path}")

# Example usage
extract_persian_ocr_to_word(
    input('PDF file name (without .pdf): ') + ".pdf",
    input('Output DOCX file name (without .docx): ') + ".docx"
)
