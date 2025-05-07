import fitz  # from PyMuPDF

def extract_persian_text(pdf_path, output_txt_path):
    doc = fitz.open(pdf_path)
    with open(output_txt_path, "w", encoding="utf-8") as out_file:
        for page_num, page in enumerate(doc, start=1):
            text = page.get_text()
            out_file.write(f"\n=== PAGE {page_num} ===\n")
            out_file.write(text.strip())
            out_file.write("\n")

    print(f"Text extracted to {output_txt_path}")

extract_persian_text("file.pdf", "extracted_text.txt")
