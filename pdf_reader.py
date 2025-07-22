import fitz  # PyMuPDF

def extract_text_by_page(pdf_path):
    doc = fitz.open(pdf_path)
    pages_text = []
    for page in doc:
        blocks = page.get_text("dict")["blocks"]
        text_items = []
        for block in blocks:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        text_items.append({
                            "text": span["text"],
                            "bold": span["flags"] & 2 != 0,
                            "underline": span["flags"] & 4 != 0,
                            "size": span["size"],
                            "font": span["font"],
                        })
        pages_text.append(text_items)
    return pages_text
