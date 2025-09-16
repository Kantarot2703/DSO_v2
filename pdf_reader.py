import re
import fitz  
import logging
from PIL import Image as _PIL_Image


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
                elif op == "re": 
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

# helpers to remove colored overlay lines and detect plus shape
def _remove_colored_lines(rgb_roi):
    if cv2 is None or np is None:
        return rgb_roi

    roi = rgb_roi
    H, W = roi.shape[:2]

    hsv = cv2.cvtColor(roi, cv2.COLOR_RGB2HSV)
    h, s, v = cv2.split(hsv)
    sat_mask = (s > 110).astype(np.uint8) * 255              
    val_mask = ((v > 40) & (v < 245)).astype(np.uint8) * 255 
    m1 = cv2.bitwise_and(sat_mask, val_mask)

    g   = cv2.cvtColor(roi, cv2.COLOR_RGB2GRAY)
    k3  = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    whitehat = cv2.morphologyEx(g, cv2.MORPH_TOPHAT,   k3)
    blackhat = cv2.morphologyEx(g, cv2.MORPH_BLACKHAT, k3)
    _, m2a = cv2.threshold(whitehat, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    _, m2b = cv2.threshold(blackhat, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    m2 = cv2.bitwise_or(m2a, m2b)

    edges = cv2.Canny(g, 60, 160)
    m3 = np.zeros_like(g)
    try:
        lines = cv2.HoughLinesP(
            edges, 1, np.pi/180, threshold=25,
            minLineLength=max(6, int(0.12 * min(H, W))),
            maxLineGap=3
        )
        if lines is not None:
            for x1, y1, x2, y2 in lines[:, 0, :]:
                cv2.line(m3, (x1, y1), (x2, y2), 255, 2)
    except Exception:
        pass

    mask = cv2.bitwise_or(m1, cv2.bitwise_or(m2, m3))

    k = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN,  k, iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, k, iterations=1)

    ratio = float(mask.sum()) / (255.0 * max(1, H * W))
    if ratio > 0.35:
        mask = cv2.erode(mask, k, iterations=1)

    if mask.sum() == 0:
        return roi

    cleaned = cv2.inpaint(roi, mask, 3, cv2.INPAINT_TELEA)
    return cleaned

def _detect_plus_by_hough(gray_roi):
    if cv2 is None or np is None:
        return False
    try:
        edges = cv2.Canny(gray_roi, 60, 160)
        lines = cv2.HoughLinesP(edges, 1, np.pi/180, threshold=25, minLineLength=6, maxLineGap=3)
        
        if lines is None:
            return False

        H, W = gray_roi.shape[:2]
        cx, cy = W/2.0, H/2.0

        horizontals, verticals = [], []
        for x1, y1, x2, y2 in lines[:,0,:]:
            dx, dy = x2-x1, y2-y1
            length = (dx*dx + dy*dy)**0.5
            if length < 6:  
                continue
            ang = abs(np.degrees(np.arctan2(dy, dx)))
            if ang <= 15:  
                horizontals.append((x1, y1, x2, y2))
            elif ang >= 75:  
                verticals.append((x1, y1, x2, y2))

        if not horizontals or not verticals:
            return False

        for hx1, hy1, hx2, hy2 in horizontals:
            hx0, hx1b = min(hx1, hx2), max(hx1, hx2)
            hy = (hy1 + hy2)/2.0

            for vx1, vy1, vx2, vy2 in verticals:
                vx = (vx1 + vx2)/2.0
                vy0, vy1b = min(vy1, vy2), max(vy1, vy2)

                if (hx0 <= vx <= hx1b) and (vy0 <= hy <= vy1b):
                    if abs(vx - cx) <= 0.35*W and abs(hy - cy) <= 0.35*H:
                        return True
        return False
    except Exception:
        return False

def _group_ocr_words_into_lines(ocr_words):
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

        for ln in lines:
            lcy, lh = ln["cy"], ln["h"]
            if abs(cy - lcy) <= 0.45 * max(lh, h):
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

def _ocr_extract_items(page, ocr_lang="eng+tha", zooms=None, conf_threshold=30, configs=None):
    if pytesseract is None or Image is None:
        return []

    # ใช้ซูม/คอนฟิกที่ส่งมา ถ้าไม่ส่งให้ใช้ดีฟอลต์แบบเดิม
    if zooms is None:
        zooms = [3.0, 3.6, 4.0]
    if configs is None:
        configs = [
            "--oem 3 --psm 6 -c preserve_interword_spaces=1",
            "--oem 3 --psm 7 -c preserve_interword_spaces=1",
            "--oem 3 --psm 11 -c preserve_interword_spaces=1",
            "--oem 3 --psm 13 -c preserve_interword_spaces=1",
        ]

    all_words = []

    for z in zooms:
        img, zf = _render_page_to_pil(page, zoom=z)

        img_gray = None
        if cv2 is not None and np is not None:
            try:
                arr = np.array(img.convert("RGB"))
                g_r = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
                g_y = cv2.cvtColor(arr, cv2.COLOR_RGB2YCrCb)[:, :, 0]
                img_gray = g_r if g_r.std() >= g_y.std() else g_y
                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
                img_gray = clahe.apply(img_gray)
            except Exception:
                img_gray = None

        candidates = [img.convert("L")]
        if img_gray is not None and cv2 is not None:
            try:
                den = cv2.fastNlMeansDenoising(img_gray, None, 10, 7, 21)
                thr = cv2.adaptiveThreshold(den, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                            cv2.THRESH_BINARY, 31, 15)
                inv = 255 - thr
                k3 = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
                closed = cv2.morphologyEx(thr, cv2.MORPH_CLOSE, k3, iterations=1)
                dil    = cv2.dilate(closed, k3, iterations=1)
                for arr in (den, thr, inv, closed, dil):
                    candidates.append(_PIL_Image.fromarray(arr))
            except Exception:
                pass

        def _try_ocr(img_pil, lang, cfg):
            try:
                return pytesseract.image_to_data(
                    img_pil, lang=lang, config=cfg, output_type=pytesseract.Output.DICT
                )
            except Exception:
                return None

        # จัดลำดับภาษาที่จะลอง
        BIG  = "eng+spa+fra+por+ita+deu+nld+swe+fin+dan+nor+pol+ces+slk+hun+rus+ell+tur+ara+tha"
        LITE = "eng+spa+fra+por+ita+deu+nld+tha"
        TINY = "eng+tha"
        FALL = "eng"
        langs = (ocr_lang or BIG, BIG, LITE, TINY, FALL)

        got = None
        for im in candidates:
            for cfg in configs:
                for lg in langs:
                    if not lg:
                        continue
                    data = _try_ocr(im, lg, cfg)
                    if data and len(data.get('text', []) or []) > 0:
                        got = (data, zf)
                        break
                if got: break
            if got: break

        if not got:
            continue

        data, used_zoom = got
        n = len(data.get("text", []))
        confs = data.get("conf", ["-1"] * n)
        for i in range(n):
            txt = (data["text"][i] or "").strip()
            if not txt:
                continue
            try:
                conf = float(confs[i])
            except Exception:
                conf = -1.0

            # เก็บ '+'/‘＋’ แม้คอนฟิเดนซ์ต่ำ
            low_punct_keep = txt in {"+", "＋"}
            if conf_threshold is not None and (conf < conf_threshold) and not low_punct_keep:
                continue

            x = float(data["left"][i]); y = float(data["top"][i])
            w = float(data["width"][i]); h = float(data["height"][i])

            size_pt = h / used_zoom
            size_mm = _pt_to_mm(size_pt)
            bbox_pt = (x / used_zoom, y / used_zoom, (x + w) / used_zoom, (y + h) / used_zoom)

            all_words.append({
                "text": txt,
                "bold": None,
                "italic": None,
                "underline": None,
                "size_pt": size_pt,
                "size_mm": size_mm,
                "size_unit": "pt",
                "font": "",
                "bbox": bbox_pt,
                "bbox_px": (x, y, x+w, y+h),
                "height_px": h,
                "source": "ocr",
                "confidence": conf
            })

    if not all_words:
        return []

    # จัดกลุ่มเป็นบรรทัด + ตรวจ underline จากภาพ (เหมือนเดิม)
    lines = _group_ocr_words_into_lines(all_words)
    img_gray = None
    try:
        img_hi, z_hi = _render_page_to_pil(page, zoom=4.0)
        if cv2 is not None and np is not None:
            arr = np.array(img_hi.convert("RGB"))
            img_gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    except Exception:
        pass

    if img_gray is not None:
        for ln in lines:
            X0, Y0, X1, Y1 = ln["bbox_px"]
            ul_line = _has_underline_in_roi(img_gray, X0, Y0, X1 - X0, Y1 - Y0)
            if ul_line is True:
                for w in ln["words"]:
                    w["underline"] = True
            elif ul_line is False:
                for w in ln["words"]:
                    if w["underline"] is None:
                        w["underline"] = False

    # รวม word+line items
    line_items = []
    for ln in lines:
        texts = [w["text"] for w in ln["words"] if (w.get("text") or "").strip()]
        if not texts:
            continue
        size_mm = 0.0
        for w in ln["words"]:
            try: size_mm = max(size_mm, float(w.get("size_mm") or 0.0))
            except Exception: pass
        X0, Y0, X1, Y1 = ln["bbox_px"]
        bbox_pt = (X0 / 3.0, Y0 / 3.0, X1 / 3.0, Y1 / 3.0) 
        line_items.append({
            "text": " ".join(texts),
            "bold": None,
            "italic": None,
            "underline": any(bool(w.get("underline")) for w in ln["words"]),
            "size_pt": None,
            "size_mm": size_mm,
            "size_unit": "pt",
            "font": "",
            "bbox": bbox_pt,
            "source": "ocr",
            "level": "line",
            "confidence": min((w.get("confidence", 0) for w in ln["words"]), default=0),
        })

    items = []
    for w in all_words:
        w.pop("bbox_px", None)
        w.pop("height_px", None)
        items.append(w)
    items.extend(line_items)
    return items

def _detect_vector_plus_signs(page, min_len=2.5, max_len=None,
                              center_tol=None, length_ratio_tol=0.55):
    if page is None:
        return []

    try:
        pw = float(page.rect.width)
        ph = float(page.rect.height)
        diag = (pw * pw + ph * ph) ** 0.5
    except Exception:
        pw, ph, diag = 595.0, 842.0, 1024.0 

    if max_len is None:
        max_len = max(40.0, 0.25 * pw)

    if center_tol is None:
        center_tol = max(0.8, 0.01 * diag)

    Hs, Vs = [], []

    try:
        for d in page.get_drawings():
            for it in d.get("items", []):
                op = it[0]

                # เคสเส้นตรง 
                if op == "l":
                    p0, p1 = it[1], it[2]
                    dx, dy = p1.x - p0.x, p1.y - p0.y
                    length = (dx * dx + dy * dy) ** 0.5
                    if length < min_len or length > max_len:
                        continue

                    # ใช้มุมเพื่อจัดแนว
                    if abs(dy) <= 0.8:  
                        x0, x1 = sorted((p0.x, p1.x))
                        y = (p0.y + p1.y) / 2.0
                        Hs.append((x0, y, x1))
                    elif abs(dx) <= 0.8:  
                        y0, y1 = sorted((p0.y, p1.y))
                        x = (p0.x + p1.x) / 2.0
                        Vs.append((x, y0, y1))

                elif op == "re":
                    rect = it[1]
                    w = float(rect.width)
                    h = float(rect.height)
                    x0, y0, x1, y1 = rect.x0, rect.y0, rect.x1, rect.y1

                    thin = max(6.0, 0.006 * diag)

                    if h <= thin and w >= min_len and w <= max_len:
                        yc = (y0 + y1) / 2.0
                        Hs.append((x0, yc, x1))

                    elif w <= thin and h >= min_len and h <= max_len:
                        xc = (x0 + x1) / 2.0
                        Vs.append((xc, y0, y1))
    except Exception:
        return []

    plus_boxes = []
    for (hx0, hy, hx1) in Hs:
        hcx = (hx0 + hx1) / 2.0
        hlen = (hx1 - hx0)

        for (vx, vy0, vy1) in Vs:
            vcy = (vy0 + vy1) / 2.0
            vlen = (vy1 - vy0)

            if abs(vx - hcx) <= center_tol and abs(vcy - hy) <= center_tol:
                if hlen > 0 and vlen > 0:
                    ratio = abs(vlen - hlen) / max(hlen, vlen)
                    if ratio <= length_ratio_tol:
                        x0, x1 = min(hx0, vx), max(hx1, vx)
                        y0, y1 = min(vy0, hy), max(vy1, hy)
                        pad = max(1.2, 0.003 * diag)
                        plus_boxes.append((x0 - pad, y0 - pad, x1 + pad, y1 + pad))

    return plus_boxes

def _synthesize_3plus_items_from_vectors(raw_spans, plus_boxes, proximity_pt=14.0):
    items = []

    def _center(b):
        x0, y0, x1, y1 = b
        return ( (x0+x1)/2.0, (y0+y1)/2.0 )

    threes = []
    for it in raw_spans:
        if (it.get("source") or "pdf") != "pdf":
            continue
        t = (it.get("text") or "").strip()
        if t == "3":
            b = it.get("bbox")
            if b: threes.append((it, _center(b)))

    for pb in plus_boxes:
        pc = _center(pb)
        best = None; best_d = 1e9
        for it, c in threes:
            dx = pc[0] - c[0]; dy = pc[1] - c[1]
            d = (dx*dx + dy*dy)**0.5
            if d < best_d:
                best, best_d = it, d
        if best is not None and best_d <= proximity_pt:
            size_mm = float(best.get("size_mm") or 0.0)
            items.append({
                "text": "3+",
                "bold": best.get("bold"),
                "italic": best.get("italic"),
                "underline": best.get("underline"),
                "size_mm": size_mm,
                "size_unit": "mm",
                "font": best.get("font",""),
                "source": "pdf", 
                "level": "line",
            })
    return items

def _find_token_plus_boxes_from_spans(raw_spans):
    boxes = []
    for it in raw_spans:
        if (it.get("source") or "pdf") != "pdf":
            continue
        t = (it.get("text") or "").strip()
        if t in {"+", "＋"}:
            b = it.get("bbox")
            if b:
                boxes.append(tuple(b))
    return boxes

def _synthesize_3plus_items_from_tokens(raw_spans, proximity_pt=14.0):
    """
    สังเคราะห์ '3+' จากโทเคน '3' และ '+' ที่อยู่ใกล้กันใน PDF spans
    (สำหรับเคสที่ไม่มีเส้นเวกเตอร์เป็น '+')
    """
    def _center(b): return ((b[0]+b[2])/2.0, (b[1]+b[3])/2.0)

    threes = []
    pluses = []
    for it in raw_spans:
        if (it.get("source") or "pdf") != "pdf":
            continue
        t = (it.get("text") or "").strip()
        if t == "3" and it.get("bbox"):
            threes.append(it)
        elif t in {"+", "＋"} and it.get("bbox"):
            pluses.append(it)

    items = []
    for p in pluses:
        pc = _center(p["bbox"])
        best, best_d = None, 1e9
        for t in threes:
            tc = _center(t["bbox"])
            d = ((pc[0]-tc[0])**2 + (pc[1]-tc[1])**2)**0.5
            if d < best_d:
                best, best_d = t, d
        if best and best_d <= proximity_pt:
            size_mm = float(best.get("size_mm") or p.get("size_mm") or 0.0)
            items.append({
                "text": "3+",
                "bold": bool(best.get("bold")) or bool(p.get("bold")),
                "italic": bool(best.get("italic")) or bool(p.get("italic")),
                "underline": bool(best.get("underline")) or bool(p.get("underline")),
                "size_mm": size_mm,
                "size_unit": "mm",
                "font": best.get("font","") or p.get("font",""),
                "source": "pdf",
                "level": "line",
                "bbox": [
                    min(best["bbox"][0], p["bbox"][0]),
                    min(best["bbox"][1], p["bbox"][1]),
                    max(best["bbox"][2], p["bbox"][2]),
                    max(best["bbox"][3], p["bbox"][3]),
                ],
            })
    return items

def _join_adjacent_3_plus(items, max_gap_factor=1.2, same_line_tol=0.6):
    pluses, threes = [], []
    for it in items or []:
        b = it.get("bbox")
        t = (it.get("text") or "").strip()
        if not b or not t:
            continue
        if t in {"+", "＋"}:
            pluses.append(it)
        elif t == "3":
            threes.append(it)

    synth = []
    for p in pluses:
        px0, py0, px1, py1 = p["bbox"]
        ph = max(1.0, py1 - py0)

        best, best_gap = None, 1e9
        for t in threes:
            tx0, ty0, tx1, ty1 = t["bbox"]
            th = max(1.0, ty1 - ty0)

            if abs((ty0+ty1)/2.0 - (py0+py1)/2.0) <= same_line_tol * max(ph, th):
                if tx1 <= px0 + 0.3 * max(ph, th):
                    gap = px0 - tx1
                    if 0 <= gap <= max_gap_factor * max(ph, th) and gap < best_gap:
                        best, best_gap = t, gap

        if best:
            size_mm = float(best.get("size_mm") or p.get("size_mm") or 0.0)
            synth.append({
                "text": "3+",
                "bold": bool(best.get("bold")) or bool(p.get("bold")),
                "italic": bool(best.get("italic")) or bool(p.get("italic")),
                "underline": bool(best.get("underline")) or bool(p.get("underline")),
                "size_mm": size_mm,
                "size_unit": "mm",
                "font": best.get("font","") or p.get("font",""),
                "source": "pdf",
                "level": "line",
                "bbox": [
                    min(best["bbox"][0], p["bbox"][0]),
                    min(best["bbox"][1], p["bbox"][1]),
                    max(best["bbox"][2], p["bbox"][2]),
                    max(best["bbox"][3], p["bbox"][3]),
                ],
            })
    return synth

def _page_has_3plus_text(items):
    P = re.compile(r"(?<!\w)3\s*[\+\＋](?!\w)")
    for it in items or []:
        t = (it.get("text") or "").strip()
        if t and P.search(t):
            return True
    return False

def _ocr_3plus_via_roi(page, plus_boxes, zoom=4.0):
    out = []
    if not plus_boxes or pytesseract is None or cv2 is None or np is None:
        return out

    img_pil, z = _render_page_to_pil(page, zoom=zoom)
    rgb = np.array(img_pil.convert("RGB"))
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    H, W = gray.shape[:2]

    def _clip(v, lo, hi): 
        return max(lo, min(int(v), hi))

    for (x0, y0, x1, y1) in plus_boxes:
        X0, Y0, X1, Y1 = x0 * z, y0 * z, x1 * z, y1 * z
        w = max(1.0, X1 - X0); h = max(1.0, Y1 - Y0)

        left_pad   = 2.2 * w
        right_pad  = 0.9 * w
        top_pad    = 0.9 * h
        bottom_pad = 0.9 * h

        rx0 = _clip(X0 - left_pad, 0, W - 1)
        rx1 = _clip(X1 + right_pad, 0, W - 1)
        ry0 = _clip(Y0 - top_pad, 0, H - 1)
        ry1 = _clip(Y1 + bottom_pad, 0, H - 1)
        if rx1 <= rx0 or ry1 <= ry0:
            continue

        roi_rgb = rgb[ry0:ry1, rx0:rx1].copy()
        roi_rgb = _remove_colored_lines(roi_rgb)
        roi_g   = cv2.cvtColor(roi_rgb, cv2.COLOR_RGB2GRAY)

        candidates = [roi_g]
        try:
            den = cv2.fastNlMeansDenoising(roi_g, None, 10, 7, 21)
            thr = cv2.adaptiveThreshold(den, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                        cv2.THRESH_BINARY, 31, 15)
            inv = 255 - thr
            k3  = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
            close = cv2.morphologyEx(thr, cv2.MORPH_CLOSE, k3, iterations=1)
            candidates += [thr, inv, close]
        except Exception:
            pass

        hit = False
        for img in candidates:
            try:
                data = pytesseract.image_to_data(
                    _PIL_Image.fromarray(img),
                    lang="eng",
                    config="--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789+＋",
                    output_type=pytesseract.Output.DICT,
                )
            except Exception:
                data = None
            if not data:
                continue
            joined = " ".join([(data["text"][i] or "").strip()
                               for i in range(len(data.get("text", []))) if (data["text"][i] or "").strip()])
            if re.search(r"(?<!\w)3\s*[\+\＋](?!\w)", joined) or ("+" in joined or "＋" in joined):
                size_pt = (ry1 - ry0) / max(1.0, z)
                size_mm = _pt_to_mm(size_pt)
                out.append({
                    "text": "3+",
                    "bold": None, "italic": None, "underline": None,
                    "size_mm": size_mm, "size_unit": "mm", "font": "",
                    "source": "ocr", "level": "line",
                })
                hit = True
                break

        if not hit and _detect_plus_by_hough(roi_g):
            size_pt = (ry1 - ry0) / max(1.0, z)
            size_mm = _pt_to_mm(size_pt)
            out.append({
                "text": "3+",
                "bold": None, "italic": None, "underline": None,
                "size_mm": size_mm, "size_unit": "mm", "font": "",
                "source": "ocr", "level": "line",
            })

    return out

def _ocr_plus_next_to_three(page, three_boxes, zoom=4.0, max_targets=4):
    out = []
    if not three_boxes or pytesseract is None or cv2 is None or np is None:
        return out

    img_pil, z = _render_page_to_pil(page, zoom=zoom)
    rgb = np.array(img_pil.convert("RGB"))
    H, W = rgb.shape[:2]

    tb = sorted(three_boxes, key=lambda b: (b[3]-b[1])*(b[2]-b[0]), reverse=True)[:max_targets]

    def _clip(v, lo, hi): 
        return max(lo, min(int(v), hi))

    for (x0, y0, x1, y1) in tb:
        X0, Y0, X1, Y1 = x0 * z, y0 * z, x1 * z, y1 * z
        w = max(1.0, X1 - X0); h = max(1.0, Y1 - Y0)

        rx0 = _clip(X1 - 0.2*w, 0, W-1)
        rx1 = _clip(X1 + 1.4*w, 0, W-1)
        ry0 = _clip(Y0 - 0.3*h, 0, H-1)
        ry1 = _clip(Y1 + 0.3*h, 0, H-1)
        if rx1 <= rx0 or ry1 <= ry0:
            continue

        roi_rgb = rgb[ry0:ry1, rx0:rx1].copy()
        roi_rgb = _remove_colored_lines(roi_rgb)
        roi_g   = cv2.cvtColor(roi_rgb, cv2.COLOR_RGB2GRAY)

        candidates = [roi_g]
        try:
            den = cv2.fastNlMeansDenoising(roi_g, None, 10, 7, 21)
            thr = cv2.adaptiveThreshold(den, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                        cv2.THRESH_BINARY, 31, 15)
            inv = 255 - thr
            k3  = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
            close = cv2.morphologyEx(thr, cv2.MORPH_CLOSE, k3, iterations=1)
            candidates += [thr, inv, close]
        except Exception:
            pass

        hit = False
        for img in candidates:
            try:
                data = pytesseract.image_to_data(
                    _PIL_Image.fromarray(img),
                    lang="eng",
                    config="--oem 3 --psm 7 -c tessedit_char_whitelist=+＋",
                    output_type=pytesseract.Output.DICT,
                )
            except Exception:
                data = None
            if not data:
                continue
            joined = " ".join([(data["text"][i] or "").strip()
                               for i in range(len(data.get("text", []))) if (data["text"][i] or "").strip()])
            if ("+" in joined) or ("＋" in joined):
                size_pt = (ry1 - ry0) / max(1.0, z)
                size_mm = _pt_to_mm(size_pt)
                out.append({
                    "text": "3+",
                    "bold": None, "italic": None, "underline": None,
                    "size_mm": size_mm, "size_unit": "mm", "font": "",
                    "source": "ocr", "level": "line",
                })
                hit = True
                break

        if not hit and _detect_plus_by_hough(roi_g):
            size_pt = (ry1 - ry0) / max(1.0, z)
            size_mm = _pt_to_mm(size_pt)
            out.append({
                "text": "3+",
                "bold": None, "italic": None, "underline": None,
                "size_mm": size_mm, "size_unit": "mm", "font": "",
                "source": "ocr", "level": "line",
            })

    return out

def extract_text_by_page(pdf_path, enable_ocr=True, ocr_lang="eng+tha", ocr_only_suspect_pages=True,
                         ocr_lang_fast=None, ocr_lang_full=None):
    
    if (ocr_lang_fast is None) and (ocr_lang_full is None):
        ocr_lang_fast = ocr_lang or "eng"
        ocr_lang_full = ocr_lang_fast
    elif (ocr_lang_fast is None) and (ocr_lang_full is not None):   
        ocr_lang_fast = ocr_lang_full     
    elif (ocr_lang_full is None) and (ocr_lang_fast is not None):
        ocr_lang_full = ocr_lang_fast

    doc = fitz.open(pdf_path)

    # ใช้ normalize สำหรับตรวจ SPW/SPG บนชั้นข้อความ PDF (ภายในฟังก์ชันนี้)
    def _norm_sp(s: str) -> str:
        s = "" if s is None else str(s)
        s = s.replace("\u00A0", " ")            
        s = s.replace("‐", "-").replace("–", "-").replace("—", "-")
        s = re.sub(r"\s+", " ", s).strip().lower()
        return s

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
                            "bold": (flags & 2) != 0 or (
                                "bold" in fontname.lower()
                                or re.search(
                                    r"(?i)(?:-|_)?("
                                    r"black|heavy|ultra\s*bold|extra\s*bold|semi\s*bold|semibold|demi\s*bold|demibold|"
                                    r"medium|med|md|boldmt|blk|bd|sb"
                                    r")\b",
                                    fontname
                                ) is not None
                            ),
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

            # รวมเป็น line-items ต่อบรรทัด 
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

            page_items = [dict(it) for it in raw_spans]

            try:
                if not _page_has_3plus_text(page_items):
                    plus_boxes_vec = _detect_vector_plus_signs(page)
                    if plus_boxes_vec:
                        synth_vec = _synthesize_3plus_items_from_vectors(raw_spans, plus_boxes_vec, proximity_pt=14.0)
                        if synth_vec:
                            page_items = _dedup_extend_items(page_items, synth_vec)

                    plus_boxes_tok = _find_token_plus_boxes_from_spans(raw_spans)
                    if plus_boxes_tok:
                        synth_tok = _synthesize_3plus_items_from_tokens(raw_spans, proximity_pt=14.0)
                        if synth_tok:
                            page_items = _dedup_extend_items(page_items, synth_tok)

                    three_boxes = []
                    for it in raw_spans:
                        if (it.get("source") or "pdf") == "pdf" and (it.get("text") or "").strip() == "3" and it.get("bbox"):
                            three_boxes.append(tuple(it["bbox"]))
                    for it in page_items:
                        if (it.get("text") or "").strip() == "3" and it.get("bbox"):
                            three_boxes.append(tuple(it["bbox"]))

                    anchors = (plus_boxes_vec or []) + (plus_boxes_tok or [])
                    if anchors:
                        def _center(b): return ((b[0]+b[2])/2.0, (b[1]+b[3])/2.0)
                        def _score(box):
                            x0, y0, x1, y1 = box
                            area = max(1e-6, (x1 - x0) * (y1 - y0))
                            if three_boxes:
                                cx, cy = _center(box)
                                d = min((((cx - _center(tb)[0]) ** 2) + ((cy - _center(tb)[1]) ** 2)) ** 0.5 for tb in three_boxes)
                            else:
                                d = 1e3
                            return (d, -area) 

                        anchors_sorted = sorted(anchors, key=_score)
                        anchors_top = anchors_sorted[:4] 

                        roi_items = _ocr_3plus_via_roi(page, anchors_top, zoom=4.0)
                        if roi_items:
                            page_items = _dedup_extend_items(page_items, roi_items)

                    if not _page_has_3plus_text(page_items) and three_boxes:
                        roi_from_three = _ocr_plus_next_to_three(page, three_boxes, zoom=4.0, max_targets=4)
                        if roi_from_three:
                            page_items = _dedup_extend_items(page_items, roi_from_three)

            except Exception:
                pass

            # OCR fallback 
            if enable_ocr:
                has_images = False
                try:
                    has_images = bool(page.get_images(full=True))
                except Exception:
                    has_images = False

                do_ocr = True

                base_skip = False
                force_sp_ocr = False

                if ocr_only_suspect_pages and not has_images:
                    enough_items = len(page_items) >= 5
                    has_readable_size = any((it.get("size_mm") or 0) >= 1.0 for it in page_items)
                    base_skip = (enough_items and has_readable_size)

                    # บังคับ OCR เฉพาะหน้า "เสี่ยง SP":
                    # มี small parts บน text-layer แต่ยังไม่เห็น may be generat...
                    # หรือพบหัวข้อ International warning statement
                    texts_join = " ".join(_norm_sp(it.get("text", "")) for it in page_items if it.get("text"))
                    has_small_parts  = ("small parts" in texts_join)
                    has_mbg_keyword  = ("small parts may be generat" in texts_join)
                    has_iws_heading  = ("international warning statement" in texts_join)

                    force_sp_ocr = (has_small_parts and not has_mbg_keyword) or has_iws_heading

                # สรุปว่าจะ OCR ไหม (ครอบคลุมทุกกรณี)
                do_ocr = (not base_skip) or force_sp_ocr

                if do_ocr:
                    fast_zooms   = [2.6, 3.0]
                    fast_cfgs    = [
                        "--oem 3 --psm 6 -c preserve_interword_spaces=1",
                        "--oem 3 --psm 11 -c preserve_interword_spaces=1",
                    ]
                    ocr_items = _ocr_extract_items(
                        page,
                        ocr_lang=ocr_lang_fast,
                        zooms=fast_zooms,
                        conf_threshold=35,
                        configs=fast_cfgs
                    )

                    need_full = False
                    if not ocr_items:
                        need_full = True
                    else:
                        text_join = " ".join([(it.get("text") or "") for it in ocr_items])[:600]
                        few_words = sum(1 for it in ocr_items if (it.get("text") or "").strip()) < 8
                        miss_plus = ("+" not in text_join) and ("＋" not in text_join)
                        need_full = (few_words and miss_plus)

                    if need_full and (ocr_lang_full and (ocr_lang_full != ocr_lang_fast)):
                        full_zooms = [3.6, 4.0]
                        full_cfgs  = [
                            "--oem 3 --psm 6 -c preserve_interword_spaces=1",
                            "--oem 3 --psm 7 -c preserve_interword_spaces=1",
                            "--oem 3 --psm 11 -c preserve_interword_spaces=1",
                        ]
                        ocr_items = _ocr_extract_items(
                            page,
                            ocr_lang=ocr_lang_full,
                            zooms=full_zooms,
                            conf_threshold=30,  
                            configs=full_cfgs
                        )

                    if ocr_items:
                        page_items = _dedup_extend_items(page_items, ocr_items)

                    # หลังรวม OCR แล้ว ลอง join '3' และ '+' ที่อยู่ชิดกันเป็น '3+'
                    try:
                        if not _page_has_3plus_text(page_items):
                            has3 = any((it.get("text") or "").strip() == "3" for it in page_items)
                            hasPlus = any((it.get("text") or "").strip() in {"+", "＋"} for it in page_items)
                            if has3 and hasPlus:
                                synth_join = _join_adjacent_3_plus(page_items)
                                if synth_join:
                                    page_items = _dedup_extend_items(page_items, synth_join)
                    except Exception:
                        pass

                    try:
                        if not _page_has_3plus_text(page_items):
                            three_boxes_ocr = [
                            tuple(it["bbox"])
                            for it in page_items
                            if (it.get("source") or "").lower() == "ocr"
                                and (it.get("text") or "").strip() == "3"
                                and it.get("bbox") is not None
                            ]
                            if three_boxes_ocr:
                                roi_from_three = _ocr_plus_next_to_three(page, three_boxes_ocr, zoom=4.0, max_targets=6)
                                if roi_from_three:
                                    page_items = _dedup_extend_items(page_items, roi_from_three)
                    except Exception:
                        pass

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