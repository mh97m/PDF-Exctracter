from pdf2image import convert_from_path
import pytesseract
from PIL import Image

# Use Persian (Farsi) OCR
pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'  # adjust path if needed
lang = 'fas'

def extract_persian_ocr(pdf_path, output_txt_path):
    images = convert_from_path(pdf_path)
    with open(output_txt_path, "w", encoding="utf-8") as out_file:
        for i, image in enumerate(images, start=1):
            text = pytesseract.image_to_string(image, lang=lang)
            out_file.write(f"\n=== PAGE {i} ===\n")
            out_file.write(text.strip())
            out_file.write("\n")
    print(f"OCR text saved to {output_txt_path}")

# Example usage
extract_persian_ocr("file.pdf", "extracted_text_ocr.txt")
