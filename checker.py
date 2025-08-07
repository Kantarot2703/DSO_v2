import pandas as pd

def is_all_caps(text):
    return text == text.upper() and any(c.isalpha() for c in text)

def check_term_in_page(term, page_items, rule):
    results = []

    for item in page_items:
        if term in item["text"]:
            matched = True
            reasons = []

            # ตรวจเงื่อนไขรูปแบบ
            if rule.get('Bold', False) and not item['bold']:
                matched = False
                reasons.append("Not bold")

            if rule.get('Underline', False) and not item['underline']:
                matched = False
                reasons.append("Not underlined")

            if rule.get('Uppercase', False) and not item['text'].istitle():
                matched = False
                reasons.append("Not uppercase")

            if rule.get('All Caps (ตัวหนา+ตัวใหญทั้งหมด)', False) and not is_all_caps(item['text']):
                matched = False
                reasons.append("Not all caps")

            # ตรวจขนาด font
            min_size = rule.get('Min Size Value')
            operator = rule.get('Operator')

            if pd.notna(min_size) and operator == ">=":
                if item['size'] < min_size:
                    matched = False
                    reasons.append(f"Font size < {min_size}")

            results.append({
                "found": True,
                "matched": matched,
                "text": item["text"],
                "reasons": reasons
            })

    if not results:
        return {
            "found": False,
            "matched": False,
            "text": "",
            "reasons": ["Term not found"]
        }

    # Return first match
    return results[0]
