import fitz 
import re


# Helpers to detect graphic underlines
def _collect_underline_segments(page):
    segs = []
    try:
        for d in page.get_drawings():
            for it in d.get("items", []):
                op = it[0]
                if op == "l":  # เส้นตรง
                    p0, p1 = it[1], it[2]
                    if abs(p0.y - p1.y) <= 1.2 and abs(p1.x - p0.x) >= 4:
                        segs.append((min(p0.x, p1.x), (p0.y + p1.y)/2.0, max(p0.x, p1.x)))
                elif op == "re":  # สี่เหลี่ยมเตี้ย ๆ
                    rect = it[1]
                    if rect.height <= 2.0 and rect.width >= 4.0:
                        yline = rect.y1 - rect.height/2.0
                        segs.append((rect.x0, yline, rect.x1))
    except Exception:
        pass
    return segs

def _x_overlap(a0, a1, b0, b1):
    return max(0.0, min(a1, b1) - max(a0, b0))


def _pt_to_mm(pt: float) -> float:
    # 1 pt = 1/72 inch, 1 inch = 25.4 mm
    return (pt or 0.0) * 25.4 / 72.0

def extract_text_by_page(pdf_path):
    doc = fitz.open(pdf_path)
    all_pages = []

    for page_index in range(len(doc)):
        page = doc.load_page(page_index)
        blocks = page.get_text("dict")["blocks"]
        raw_spans = [] 
        for block in blocks:
            if "lines" not in block:
                continue
            for line in block["lines"]:
                if "spans" not in line:
                    continue
                for span in line["spans"]:
                    text = (span.get("text") or "").strip()
                    if not text:
                        continue

                    size_pt  = float(span.get("size", 0) or 0)
                    size_mm  = _pt_to_mm(size_pt)
                    fontname = span.get("font", "") or ""
                    flags    = int(span.get("flags", 0) or 0)
                    bbox     = span.get("bbox", None)

                    raw_spans.append({
                        "text": text,
                        "bold": (flags & 2) != 0 or ("bold" in fontname.lower()),
                        "italic": (flags & 1) != 0,
                        "underline": ((flags & 8) != 0) or ("underline" in fontname.lower()),
                        "size_pt": size_pt,
                        "size_mm": size_mm,
                        "size_unit": "pt",
                        "font": fontname,
                        "bbox": bbox,
                    })

        # เติม underline จาก “เส้นกราฟิก” (เส้น/สี่เหลี่ยมเตี้ย ๆ)
        segs = _collect_underline_segments(page)
        if segs:
            for it in raw_spans:
                if it.get("underline"):
                    continue  # มีอยู่แล้ว
                b = it.get("bbox")
                if not b:
                    continue
                x0, y0, x1, y1 = b
                width = max(1.0, x1 - x0)
                # baseline ≈ y1; ยอม ±2pt และทับแกน X ≥ 50%
                for sx0, sy, sx1 in segs:
                    if abs(sy - y1) <= 2.0 and _x_overlap(x0, x1, sx0, sx1) >= 0.5 * width:
                        it["underline"] = True
                        break

        # ส่งออกหน้า (ตัด bbox ทิ้งได้)
        page_items = [{k: v for k, v in it.items() if k != "bbox"} for it in raw_spans]
        all_pages.append(page_items)

    return all_pages

def extract_product_info_by_page(pages, size_threshold=1.6):
    product_infos = []
    for page_num, page_items in enumerate(pages, start=1):
        products = []
        part_no = ""
        rev = ""
        for item in page_items:
            size_mm = float(item.get("size_mm", item.get("size", 0)))
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