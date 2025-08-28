import re
import fitz  

# OCR text 
try:
    import pytesseract
    from PIL import Image
except Exception:
    pytesseract = None
    Image = None

# ตรวจเส้นใต้จากภาพ
try:
    import cv2
    import numpy as np
except Exception:
    cv2 = None
    np = None


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

def _render_page_to_pil(page, zoom=2.0):
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    mode = "RGBA" if pix.alpha else "RGB"
    img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
    return img, zoom

def _has_underline_in_roi(img_gray, x, y, w, h):
    if img_gray is None or cv2 is None:
        return None
    H, W = img_gray.shape[:2]
    if w <= 3 or h <= 3:
        return None

    # ROI กว้างขึ้นเล็กน้อย: ใต้ baseline ~0.85h ถึง 1.40h
    x1 = _safe_int(x + 0.04 * w, 0, W - 1)
    x2 = _safe_int(x + 0.96 * w, 0, W - 1)
    y1 = _safe_int(y + 0.85 * h, 0, H - 1)
    y2 = _safe_int(y + 1.40 * h, 0, H - 1)

    if x2 <= x1 or y2 <= y1:
        return None

    roi = img_gray[y1:y2, x1:x2]
    if roi.size == 0:
        return None

    _th, bw = cv2.threshold(roi, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)

    # ผ่อน threshold จาก 0.70 → 0.60
    row_coverage = (bw.sum(axis=1) / 255.0) / max(1, bw.shape[1])
    if row_coverage.size == 0:
        return None
    return bool(row_coverage.max() >= 0.60)

def _safe_int(v, lo, hi):
    return max(lo, min(int(v), hi))

def _dedup_extend_items(existing_items, new_items, iou_thresh=0.6):
    def _norm(s): return (s or "").strip().lower()
    def _iou(a, b):
        ax0, ay0, ax1, ay1 = a
        bx0, by0, bx1, by1 = b
        inter_x0, inter_y0 = max(ax0, bx0), max(ay0, by0)
        inter_x1, inter_y1 = min(ax1, bx1), min(ay1, by1)
        iw, ih = max(0, inter_x1 - inter_x0), max(0, inter_y1 - inter_y0)
        inter = iw * ih
        if inter <= 0: return 0.0
        aarea = (ax1 - ax0) * (ay1 - ay0)
        barea = (bx1 - bx0) * (by1 - by0)
        return inter / max(1e-6, (aarea + barea - inter))

    out = existing_items[:]
    for ni in new_items:
        keep = True
        nt = _norm(ni.get("text"))
        nb = ni.get("bbox")
        for ei in existing_items:
            et = _norm(ei.get("text"))
            eb = ei.get("bbox")
            if nt and et and nt == et and nb and eb and _iou(nb, eb) >= iou_thresh:
                keep = False
                break
        if keep:
            out.append(ni)
    return out

def _union_bbox_px(b):
    x0, y0, x1, y1 = b[0], b[1], b[2], b[3]
    return [x0, y0, x1, y1]

def _merge_bbox_px(b1, b2):
    return [
        min(b1[0], b2[0]),
        min(b1[1], b2[1]),
        max(b1[2], b2[2]),
        max(b1[3], b2[3]),
    ]

def _group_ocr_words_into_lines(ocr_words):
    """
    ocr_words: list of dict (ต้องมี bbox_px, height_px)
    กลุ่มด้วย y-center ใกล้กันและช่องว่าง x เล็ก
    """
    if not ocr_words:
        return []

    # จัดเรียงซ้าย→ขวา, บน→ล่าง
    ws = sorted(ocr_words, key=lambda w: ( (w["bbox_px"][1]+w["bbox_px"][3])/2.0, w["bbox_px"][0] ))

    lines = []
    for w in ws:
        x0,y0,x1,y1 = w["bbox_px"]
        cy = (y0+y1)/2.0
        h  = max(1.0, y1-y0)
        placed = False

        # y_tol = 0.45*h: ผ่อนสักหน่อยสำหรับฟอนต์เล็ก/พิมพ์ไทย
        for ln in lines:
            lcy, lh = ln["cy"], ln["h"]
            if abs(cy - lcy) <= 0.45 * max(lh, h):
                # ช่องไฟ x ไม่ห่างเกิน 1.2*h
                if not ln["words"] or (x0 - ln["words"][-1]["bbox_px"][2]) <= 1.2 * max(lh, h):
                    ln["words"].append(w)
                    ln["bbox_px"] = _merge_bbox_px(ln["bbox_px"], w["bbox_px"])
                    ln["cy"] = (ln["cy"]*len(ln["words"]) + cy)/(len(ln["words"])+1)
                    ln["h"]  = (ln["h"]*len(ln["words"]) + h )/(len(ln["words"])+1)
                    placed = True
                    break
        if not placed:
            lines.append({
                "words":[w],
                "bbox_px": list(w["bbox_px"]),
                "cy": cy,
                "h":  h,
            })
    return lines

def _ocr_extract_items(page, ocr_lang="eng+tha", zoom=2.0, conf_threshold=60):
    if pytesseract is None or Image is None:
        return []

    img, z = _render_page_to_pil(page, zoom=zoom)

    img_gray = None
    if cv2 is not None and np is not None:
        try:
            arr = np.array(img.convert("RGB"))
            img_gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
        except Exception:
            img_gray = None

    try:
        data = pytesseract.image_to_data(img, lang=ocr_lang, output_type=pytesseract.Output.DICT)
    except Exception:
        return []

    tmp = [] 
    n = len(data.get("text", []))
    for i in range(n):
        txt = (data["text"][i] or "").strip()
        if not txt:
            continue

        try:
            conf = float(data.get("conf", ["-1"]*n)[i])
        except Exception:
            conf = -1.0
        if conf_threshold is not None and conf < conf_threshold:
            continue

        # พิกเซลจากภาพเรนเดอร์ (สำหรับ grouping)
        x = float(data["left"][i]); y = float(data["top"][i])
        w = float(data["width"][i]); h = float(data["height"][i])

        size_pt = h / z
        size_mm = _pt_to_mm(size_pt)
        bbox_pt = (x / z, y / z, (x + w) / z, (y + h) / z)

        tmp.append({
            "text": txt,
            "bold": None,
            "italic": None,
            "underline": None,    
            "size_pt": size_pt,
            "size_mm": size_mm,
            "font": "",
            "bbox": bbox_pt,
            "bbox_px": (x, y, x+w, y+h),
            "height_px": h,
            "source": "ocr",
            "confidence": conf
        })

    if not tmp:
        return []

    # group เป็นบรรทัด แล้วตรวจเส้นใต้ระดับบรรทัด
    lines = _group_ocr_words_into_lines(tmp)
    for ln in lines:
        X0,Y0,X1,Y1 = ln["bbox_px"]
        ul_line = _has_underline_in_roi(img_gray, X0, Y0, X1-X0, Y1-Y0) if img_gray is not None else None
        if ul_line is True:
            for w in ln["words"]:
                w["underline"] = True
        elif ul_line is False:
            for w in ln["words"]:
                if w["underline"] is None:
                    w["underline"] = False

    # คืน items (ตัด bbox_px ช่วยลดน้ำหนัก)
    items = []
    for w in tmp:
        w.pop("bbox_px", None)
        w.pop("height_px", None)
        items.append(w)
    return items

def extract_text_by_page(pdf_path, enable_ocr=True, ocr_lang="eng+tha", ocr_only_suspect_pages=True):
    doc = fitz.open(pdf_path)
    try:
        all_pages = []

        for page_index in range(len(doc)):
            page = doc.load_page(page_index)
            blocks = page.get_text("dict")["blocks"]

            raw_spans = []
            line_groups = []

            for block in blocks:
                if "lines" not in block:
                    continue

                for line in block["lines"]:
                    if "spans" not in line:
                        continue

                    __line_indices = []

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
                            "source": "pdf",
                        })
                        __line_indices.append(len(raw_spans) - 1)

                    if __line_indices:
                        line_groups.append(__line_indices)

            # เติม underline จากเส้นกราฟิก
            segs = _collect_underline_segments(page)
            if segs:
                for it in raw_spans:
                    if it.get("underline"):
                        continue
                    b = it.get("bbox")
                    if not b:
                        continue
                    x0, y0, x1, y1 = b
                    width = max(1.0, x1 - x0)
                    for sx0, sy, sx1 in segs:
                        if abs(sy - y1) <= 2.0 and _x_overlap(x0, x1, sx0, sx1) >= 0.5 * width:
                            it["underline"] = True
                            break

            # ยังไม่ตัด bbox ต้องใช้ dedup ตอน merge OCR
            page_items = [dict(it) for it in raw_spans]

            # --- รวมเป็น line-items ต่อบรรทัด (ทำหลังเติม underline จากกราฟิกแล้ว) ---
            for __idxs in line_groups:
                if not __idxs:
                    continue
                __spans = [raw_spans[i] for i in __idxs if 0 <= i < len(raw_spans)]
                if not __spans:
                    continue
                __texts = [s.get("text","") for s in __spans if (s.get("text") or "").strip()]
                if not __texts:
                    continue

                __bold      = any(bool(s.get("bold")) for s in __spans)
                __italic    = any(bool(s.get("italic")) for s in __spans)
                __underline = any(bool(s.get("underline")) for s in __spans)
                __size_mm   = 0.0
                for s in __spans:
                    try:
                        __size_mm = max(__size_mm, float(s.get("size_mm") or 0.0))
                    except Exception:
                        pass

                raw_spans.append({
                    "text": " ".join(__texts),
                    "bold": __bold,
                    "italic": __italic,
                    "underline": __underline,
                    "size_mm": __size_mm,
                    "size_unit": "mm",
                    "font": "",
                    "level": "line", 
                    "source": "pdf",
                })

            # ---- OCR fallback ----
            if enable_ocr:
                do_ocr = True
                if ocr_only_suspect_pages:
                    enough_items = len(page_items) >= 5
                    has_readable_size = any((it.get("size_mm") or 0) >= 1.0 for it in page_items)
                    do_ocr = not (enough_items and has_readable_size)
                if do_ocr:
                    ocr_items = _ocr_extract_items(page, ocr_lang=ocr_lang, zoom=2.0, conf_threshold=60.0)
                    if ocr_items:
                        page_items = _dedup_extend_items(page_items, ocr_items)

            # ตัด bbox ออกก่อนส่งออก
            for it in page_items:
                it.pop("bbox", None)
            all_pages.append(page_items)

        return all_pages 
    except Exception as e:

        raise
    finally:
        try:
            doc.close()
        except Exception:
            pass

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