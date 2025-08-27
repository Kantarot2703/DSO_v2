import pandas as pd
import unicodedata as _ud
import difflib

def is_all_caps(text):
    return text == text.upper() and any(c.isalpha() for c in text)

def _is_uppercase_text(s: str) -> bool:
    s = _ud.normalize("NFKC", str(s or ""))
    letters = [ch for ch in s if ch.isalpha()]
    if not letters:
        return False
    return all(ch.isupper() for ch in letters)

def _norm_text(s: str) -> str:
    s = _ud.normalize("NFKC", str(s or ""))
    return s.strip().lower()

def _fuzzy_match(a: str, b: str, threshold=0.85) -> bool:
    """ True ถ้า similarity >= threshold """
    return difflib.SequenceMatcher(None, _norm_text(a), _norm_text(b)).ratio() >= threshold

def check_term_in_page(term, page_items, rule):
    results = []
    tnorm = _norm_text(term)

    for item in page_items:
        text = item.get("text", "")
        src  = (item.get("source") or "pdf").lower()

        # ----- การพบคำ -----
        found = False
        if src == "ocr":
            found = (tnorm in _norm_text(text)) or _fuzzy_match(term, text, threshold=0.88)
        else:
            found = (tnorm in _norm_text(text))

        if not found:
            continue

        matched = True
        reasons = []

        # ----- ตรวจ Style เฉพาะเมื่อไม่ใช่ OCR -----
        if src != "ocr":
            if rule.get('Bold', False) and not item.get('bold', False):
                matched = False
                reasons.append("Not bold")

            if rule.get('Underline', False) and not item.get('underline', False):
                matched = False
                reasons.append("Not underlined")
        else:
            if rule.get('Bold', False) or rule.get('Underline', False):
                reasons.append("Style not verifiable (OCR)")

        # ----- ตรวจขนาด (mm/pt) ได้ทั้ง PDF และ OCR -----
        req_size = rule.get('MinSizeMM', None) or rule.get('SizeMM', None)
        if req_size is not None:
            try:
                size_mm = float(item.get("size_mm") or 0)
                if size_mm + 1e-6 < float(req_size):
                    matched = False
                    reasons.append(f"Font size too small ({size_mm:.2f}mm < {float(req_size):.2f}mm)")
            except Exception:
                pass

        results.append({
            "found": True,
            "matched": matched,
            "text": text,
            "source": src,
            "reasons": reasons
        })

    if not results:
        return {
            "found": False,
            "matched": False,
            "text": "",
            "reasons": ["Term not found"]
        }

    # เลือกผลลัพธ์ที่ 'matched=True' ก่อน ถ้าไม่มีให้คืนตัวแรก
    for r in results:
        if r.get("matched"):
            return r
    return results[0]