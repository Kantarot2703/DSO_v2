import os
import re
import pandas as pd
import unicodedata
from PyQt5.QtCore import Qt
from openpyxl import load_workbook
from openpyxl.styles.colors import Color
from collections import defaultdict
import logging

logging.basicConfig(level=logging.INFO, format="🔍 [%(levelname)s] %(message)s")


# Allowed part codes from PDF filenames
ALLOWED_PART_CODES = ['UU1_DOM', 'DOM', 'UU1', '2LB', '2XV', '4LB', '19L', '19A', '21A', 'DC1']

def get_strikeout_or_red_text_rows(excel_path, sheet_name, header_row_index):
    wb = load_workbook(excel_path)
    ws = wb[sheet_name]
    bad_rows = set()

    for row in ws.iter_rows(min_row=header_row_index + 2):  # ข้าม header
        row_index = row[0].row
        for cell in row:
            font = cell.font
            color = font.color

            if not font:
                continue

            is_strike = font.strike
            is_red = False

            if color:
                if color.type == "rgb" and color.rgb:
                    is_red = color.rgb.upper().startswith("FF0000")
                elif color.type == "theme" and hasattr(color, "theme"):
                    if color.theme == 10:
                        is_red = True

            if is_strike and is_red:
                bad_rows.add(row_index)
                break 
    return bad_rows

def extract_part_code_from_pdf(pdf_filename):
    basename = os.path.basename(pdf_filename).upper().replace(" ", "").replace(",", "")
    detected = [code for code in ALLOWED_PART_CODES if code in basename]
    return list(set(detected))

def fuzzy_find_columns(df):
    term_col = None
    lang_col = None
    spec_col = None  
    logging.info("🧾 Columns found in sheet:", list(df.columns))

    for col in df.columns:
        if pd.isna(col): continue
        col_str = str(col).strip().lower().replace("\xa0", "").replace(" ", "")
        if any(key in col_str for key in ['term', 'ข้อความ', 'exactwording', 'symbol', 'wording']):
            term_col = col
        if any(key in col_str for key in ['language', 'lang', 'ภาษา']):
            lang_col = col
        if any(key in col_str for key in ['specification', 'requirement', 'ข้อกำหนด']):
            spec_col = col

    return term_col, lang_col, spec_col

def normalize_text(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize('NFKC', text)
    text = re.sub(r"\([^)]*\)", "", text) 
    text = text.replace("’", "'")  
    text = re.sub(r"\s+", " ", text)    
    return text.strip().lower()

def load_checklist(excel_path, pdf_filename=None):
    all_sheets = pd.read_excel(excel_path, sheet_name=None)
    sheet_names = list(all_sheets.keys())

    if not pdf_filename:
        raise ValueError("📄 กรุณาอัปโหลดไฟล์ PDF ก่อน เพื่อจับคู่กับ Sheet ของ Checklist")

    part_codes = extract_part_code_from_pdf(pdf_filename)
    if not part_codes:
        raise ValueError("❌ ไม่พบ Part code ที่ระบุไว้ในชื่อไฟล์ PDF")

    logging.info(f"📂 PDF filename: {pdf_filename}")
    logging.info(f"🧠 Part codes detected: {part_codes}")
    logging.info(f"📄 Sheet names: {sheet_names}")

    for sheet_name in sheet_names:
        sheet_name_normalized = sheet_name.upper().replace(" ", "")
        for code in part_codes:
            if sheet_name_normalized.startswith(code):
                print(f"✅ Found matching sheet: {sheet_name}")
                df = all_sheets[sheet_name]

                # Header
                header_row_index = None
                for i in range(min(10, len(df))):
                    if df.iloc[i].notna().sum() >= 2:
                        header_row_index = i
                        break
                if header_row_index is None:
                    raise ValueError(f"❌ ไม่พบแถว header ที่เหมาะสมใน sheet: {sheet_name}")

                df.columns = df.iloc[header_row_index]
                df = df[header_row_index + 1:].reset_index(drop=True)

                # Column Mapping
                term_col, lang_col, spec_col = fuzzy_find_columns(df)
                logging.info(f"🔎 ใช้คอลัมน์ Term: {term_col}, Language: {lang_col}, Spec: {spec_col}")

                # Standardize
                if spec_col and spec_col in df.columns:
                    df[spec_col] = df[spec_col].apply(
                        lambda x: "-" if pd.isna(x) or str(x).strip().upper() in ["N/A", "NONE", "-"] else str(x)
                    )

                columns_to_ffill = [col for col in df.columns if str(col).strip().lower() in ["requirement", "language"]]
                df[columns_to_ffill] = df[columns_to_ffill].ffill()

                df = df.rename(columns={term_col: "Term (Text)"})

                # Drop Red+Strike Rows
                bad_row_numbers = get_strikeout_or_red_text_rows(excel_path, sheet_name, header_row_index)
                logging.info(f"❌ Red+Strike rows from Excel: {bad_row_numbers}")

                df["ExcelRow"] = df.index + header_row_index + 2 
                df.drop(columns=["ExcelRow"], inplace=True)

                # Filter empty Term 
                manual_keywords_for_load = ["lion", "logo", "symbol", "graphic", "trademark", "warning", "pictogram", "space", "copyright"]
                def keep_row_even_if_term_missing(row):
                    term = str(row.get("Term (Text)", "")).strip()
                    requirement = str(row.get("Requirement", "")).strip().lower()
                    if term:
                        return True
                    return any(key in requirement for key in manual_keywords_for_load)

                df = df[df.apply(keep_row_even_if_term_missing, axis=1)]

                # Drop columns
                columns_to_exclude = ["Instruction of Play function feature", "Warning statement"]
                df = df[[col for col in df.columns if col and str(col).strip().lower() not in columns_to_exclude]]

                # Explode multi-term 
                df["Term (Text)"] = df["Term (Text)"].astype(str)
                df = df.explode("Term (Text)", ignore_index=True) if df["Term (Text)"].apply(lambda x: "\n" in x or "," in x).any() else df

                # Language List 
                if lang_col:
                    df = df.rename(columns={lang_col: "Language"})
                    df["Language List"] = df["Language"].apply(lambda x: str(x).split(",") if pd.notna(x) else [])
                else:
                    df["Language List"] = [[] for _ in range(len(df))]

                # Extract from Remark 
                if "Remark" in df.columns:
                    def extract_languages_from_remark(remark, term):
                        langs = []
                        if pd.isna(remark): return langs
                        for line in str(remark).splitlines():
                            if "=" in line:
                                left, right = line.split("=", 1)
                                if term.strip().lower() in left.strip().lower():
                                    langs.append(right.strip())
                        return langs
                    df["Language List"] = [
                        extract_languages_from_remark(remark, term) or ["Unspecified"]
                        for remark, term in zip(df.get("Remark", []), df["Term (Text)"])
                    ]

                return df 

    raise ValueError("❌ ไม่พบ Sheet ที่ตรงกับ Part code จากชื่อไฟล์ PDF")

def start_check(df_checklist, extracted_text_list):
    results = []
    all_texts = []
    grouped = defaultdict(list)

    skip_keywords = [
        "instruction", "play function feature", "function feature",
        "do not print", "do not forget", "see template", "don't forget",
        "for reference", "note", "click", "reminder", "refer template", "remove from template", "remove from mb legal template"
    ]

    manual_keywords = [
        "brand logo", "copyright for t&f", "space for date code", "lion mark", "lionmark", "lion-mark", "ce mark", "en 71",
        "ukca", "mc mark", "cib", "argentina logo", "brazilian logo", "italy requirement", "france requirement",
        "sorting & donation label", "spain sorting symbols", "usa warning statement", "international warning statement",
        "upc requirement", "list of content : text", "list of content : pictorial", "product’s common", "product's common", "generic name"
    ]

    def is_artwork_page(page_items):
        if len(page_items) < 10:
            return False
        for item in page_items:
            if isinstance(item, dict) and float(item.get("size", 0)) >= 1.6:
                return True
        return False

    artwork_pages = []
    page_mapping = {}

    for real_page_index, page_items in enumerate(extracted_text_list):
        if is_artwork_page(page_items):
            artwork_pages.append(page_items)
            page_mapping[len(artwork_pages)] = real_page_index + 1

    for artwork_index, page_items in enumerate(artwork_pages):
        for item in page_items:
            text_norm = normalize_text(item.get("text", ""))
            page_number = page_mapping[artwork_index + 1]
            all_texts.append((text_norm, page_number, item))

    for idx, row in df_checklist.iterrows():
        requirement = str(row.get("Requirement", "")).strip()
        spec = str(row.get("Specification", "")).strip()

        # Normalize
        req_norm = normalize_text(requirement)
        spec = "-" if spec.lower() in ["", "n/a", "none", "unspecified", "nan"] else spec
        spec_norm = normalize_text(spec)

        # Term (ไม่ดึงค่าจากแถวอื่น)
        term_raw = row.get("Term (Text)", None)
        term_cell_raw = str(term_raw) if pd.notna(term_raw) else ""
        term_cell_clean = term_cell_raw.strip()
        term_cell_clean = "-" if term_cell_clean.lower() in ["", "n/a", "none", "unspecified", "nan"] else term_cell_clean
        term_lines = [term_cell_clean] if term_cell_clean != "-" else []


        # ตรวจ Manual
        is_manual = any(kw in req_norm for kw in manual_keywords)
        if "underline" in spec_norm and not is_manual:
            is_manual = True

        # ข้าม row ไม่จำเป็น
        if any(kw in req_norm for kw in skip_keywords) or any(kw in spec_norm for kw in skip_keywords):
            continue

        # Verified แต่ไม่มี Term → ข้าม
        if not is_manual and not term_lines:
            continue

        # --- MANUAL SECTION ---
        if is_manual:
            if not term_lines:
                grouped[(requirement, spec, "Manual")].append({
                    "Term": term_cell_raw,
                    "Found": "-",
                    "Match": "-",
                    "Pages": "-",
                    "Font Size": "-",
                    "Note": "Manual check required",
                    "Verification": "Manual"
                })
            else:
                for term in term_lines:
                    grouped[(requirement, spec, "Manual")].append({
                        "Term": term_cell_raw,
                        "Found": "-",
                        "Match": "-",
                        "Pages": "-",
                        "Font Size": "-",
                        "Note": "Manual check required",
                        "Verification": "Manual"
                    })
            continue

        # --- VERIFIED SECTION ---
        for term in term_lines:
            term_norm = normalize_text(term)
            if not term_norm or any(kw in term_norm for kw in skip_keywords):
                continue

            term_words = term_norm.split()
            found_pages = []
            matched_items = []

            for text_norm, page_number, item in all_texts:
                if any(word in text_norm for word in term_words):
                    found_pages.append(page_number)
                    matched_items.append(item)

            found_flag = "✅ Found" if found_pages else "❌ Not Found"
            match_result = "✔"
            notes = []

            sizes = [float(i.get("size", 0)) for i in matched_items]
            bolds = [i.get("bold", False) for i in matched_items]
            underlines = [i.get("underline", False) for i in matched_items]
            texts = [i.get("text", "") for i in matched_items]

            font_size_str = "-"

            if found_pages and matched_items and spec != "-":
                if "bold" in spec.lower() and not any(bolds):
                    match_result = "❌"
                    notes.append("Not Bold")
                if "underline" in spec.lower():
                    if any(underlines):
                        pass  # มีอย่างน้อย 1 คำขีดเส้นใต้ → ถือว่าผ่าน
                    else:
                        match_result = "❌"
                        notes.append("Underline missing")
                if "all caps" in spec.lower() and not any(t.isupper() for t in texts):
                    match_result = "❌"
                    notes.append("Not All Caps")

                if "≥" in spec:
                    try:
                        threshold = float(re.findall(r"≥\s*(\d+(?:\.\d+)?)", spec)[0])
                        if not any(size >= threshold for size in sizes):
                            match_result = "❌"
                            notes.append(f"Font < {threshold} mm")
                        font_size_str = "✔" if match_result == "✔" else f"{round(sizes[0], 2)} mm"
                    except:
                        font_size_str = "-"
                else:
                    font_size_str = "✔" if match_result == "✔" else "-"
            else:
                font_size_str = "-"

            pages = sorted(set(found_pages))
            all_pages = sorted(set(p for _, p, _ in all_texts))
            page_str = "All pages" if set(pages) == set(all_pages) else ", ".join(str(p) for p in pages)

            grouped[(requirement, spec, "Verified")].append({
                "Term": term,
                "Found": found_flag,
                "Match": match_result if found_flag == "✅ Found" else "❌",
                "Pages": page_str if found_flag == "✅ Found" else "-",
                "Font Size": font_size_str if found_flag == "✅ Found" else "-",
                "Note": ", ".join(notes) if notes else "-",
                "Verification": "Verified"
            })

    final_results = []
    for (requirement, spec, verification), items in grouped.items():
        for item in items:
            term_display = item.get("Term", "")
            if pd.isna(term_display) or str(term_display).strip().lower() in ["", "nan"]:
                term_display = "-"
            final_results.append({
                "Requirement": requirement,
                "Term": term_display,
                "Specification": spec,
                "Found": item["Found"],
                "Match": item["Match"],
                "Pages": item["Pages"],
                "Font Size": item["Font Size"],
                "Note": item["Note"],
                "Verification": verification
            })

    return pd.DataFrame(final_results)