import fitz 
import re


def _pt_to_mm(pt: float) -> float:
    # 1 pt = 1/72 inch, 1 inch = 25.4 mm
    return (pt or 0.0) * 25.4 / 72.0

def extract_text_by_page(pdf_path):
    doc = fitz.open(pdf_path)
    all_pages = []

    for page_index in range(len(doc)):
        page = doc.load_page(page_index)
        blocks = page.get_text("dict")["blocks"]
        page_items = []

        for block in blocks:
            if "lines" not in block:
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"].strip()
                    if not text:
                        continue

                    size_pt  = float(span.get("size", 0) or 0)
                    size_mm  = _pt_to_mm(size_pt)
                    fontname = span.get("font", "") or ""
                    flags    = int(span.get("flags", 0) or 0)

                    item = {
                        "text": text,
                        "bold": (flags & 2) != 0 or ("bold" in fontname.lower()),
                        "italic": (flags & 1) != 0,
                        "underline": ((flags & 8) != 0) or ("underline" in fontname.lower()),
                        "size_pt": size_pt,     
                        "size_mm": size_mm,     
                        "size_unit": "pt",      
                        "size": size_mm,
                        "font": fontname,
                    }
                    page_items.append(item)

        # ข้ามหน้าเปล่าหรือไม่มีข้อความเลย
        if len(page_items) >= 3: 
            all_pages.append(page_items)

    return all_pages

def extract_product_info_by_page(pages, size_threshold=1.6):
    product_infos = []
    for page_num, page_items in enumerate(pages, start=1):
        products = []
        part_no = ""
        rev = ""
        for item in page_items:
            size_mm = float(item.get("size", 0))
            text = item.get("text", "").strip()

            if size_mm >= size_threshold:
                products.append(text)

            if not part_no:
                match = re.search(r'\b[A-Z0-9]{2,5}[-][A-Z0-9]{2,6}\b', text)
                if match:
                    part_no = match.group()

            if not rev:
                match = re.search(r'\bA\d\b', text)
                if match:
                    rev = match.group()

        product_name = " ".join(products) if products else "-"
        product_infos.append({
            "page": page_num,
            "product_name": product_name.strip(),
            "part_no": part_no or "-",
            "rev": rev or "-"
        })

    return product_infos