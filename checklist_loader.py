import os
import re
import pandas as pd
import unicodedata
from PyQt5.QtCore import Qt
from openpyxl import load_workbook
from openpyxl.styles.colors import Color
from collections import defaultdict


# Allowed part codes from PDF filenames
ALLOWED_PART_CODES = ['UU1_DOM', '2LB', '2XV', '4LB', '19L', '19A', '21A', 'DC1']

def get_strikeout_or_red_text_rows(excel_path, sheet_name, header_row_index):
    wb = load_workbook(excel_path)
    ws = wb[sheet_name]
    strike_rows = set()

    for row in ws.iter_rows(min_row=header_row_index + 2):
        for cell in row:
            font = cell.font
            color = font.color
            if font and font.strike:
                if color:
                    if color.type == 'rgb':
                        hex_color = color.rgb.upper()
                        if hex_color.startswith("FF0000") or hex_color.startswith("FFFF0000"):
                            strike_rows.add(cell.row)
                            break
                    elif color.type == 'theme':
                        # ถ้ามาจากธีม ให้ถือว่าเข้าข่าย (อาจต้องระวัง False positive)
                        strike_rows.add(cell.row)
                        break
    return strike_rows


def extract_part_code_from_pdf(pdf_filename):
    basename = os.path.basename(pdf_filename).upper().replace(" ", "").replace(",", "")
    detected = [code for code in ALLOWED_PART_CODES if code in basename]
    return list(set(detected))

def fuzzy_find_columns(df):
    term_col = None
    lang_col = None
    spec_col = None  
    print("🧾 Columns found in sheet:", list(df.columns))

    for col in df.columns:
        if pd.isna(col): continue
        col_str = str(col).strip().lower().replace("\xa0", "").replace(" ", "")
        if any(key in col_str for key in ['term', 'ข้อความ', 'exactwording', 'symbol', 'wording']):
            term_col = col
        if any(key in col_str for key in ['language', 'lang', 'ภาษา']):
            lang_col = col
        if any(key in col_str for key in ['specification', 'requirement', 'ข้อกำหนด']):
            spec_col = col

    if not term_col:
        raise ValueError("❌ ไม่พบคอลัมน์ข้อความ เช่น 'Symbol / Exact wording', 'Term'\n\nคอลัมน์ที่เจอคือ: " + ", ".join([str(c) for c in df.columns]))

    return term_col, lang_col, spec_col

def load_checklist(excel_path, pdf_filename=None):
    all_sheets = pd.read_excel(excel_path, sheet_name=None)
    sheet_names = list(all_sheets.keys())

    if not pdf_filename:
        raise ValueError("📄 กรุณาอัปโหลดไฟล์ PDF ก่อน เพื่อจับคู่กับ Sheet ของ Checklist")

    part_codes = extract_part_code_from_pdf(pdf_filename)
    if not part_codes:
        raise ValueError("❌ ไม่พบ Part code ที่ระบุไว้ในชื่อไฟล์ PDF")

    print(f"📂 PDF filename: {pdf_filename}")
    print(f"🧠 Part codes detected: {part_codes}")
    print(f"📄 Sheet names: {sheet_names}")

    for sheet_name in sheet_names:
        sheet_name_normalized = sheet_name.upper().replace(" ", "")
        for code in part_codes:
            if sheet_name_normalized.startswith(code):
                print(f"✅ Found matching sheet: {sheet_name}")
                df = all_sheets[sheet_name]

                # หา header แถวแรกที่มีข้อมูลจริง
                header_row_index = None
                for i in range(min(10, len(df))):
                    if df.iloc[i].notna().sum() >= 2:
                        header_row_index = i
                        break

                if header_row_index is None:
                    raise ValueError(f"❌ ไม่พบแถว header ที่เหมาะสมใน sheet: {sheet_name}")

                df.columns = df.iloc[header_row_index]
                df = df[header_row_index + 1:]
                df.ffill(inplace=True)

                def get_strikethrough_red_rows(excel_path, sheet_name, header_row_index):
                    wb = load_workbook(excel_path)
                    ws = wb[sheet_name]
                    bad_rows = set()

                    for row_idx, row in enumerate(ws.iter_rows(min_row=header_row_index + 2), start=header_row_index + 2):
                        for cell in row:
                            font = cell.font
                            color = font.color
                            if font and font.strike and color and color.type == "rgb" and color.rgb.upper().startswith("FF0000"):
                                bad_rows.add(row_idx)
                                break  
                    return bad_rows

                # กรองแถวที่มีข้อความขีดฆ่าสีแดง
                bad_row_numbers = get_strikeout_or_red_text_rows(excel_path, sheet_name, header_row_index)
                print(f"❌ Red+Strike rows from Excel: {bad_row_numbers}")

                df = df.reset_index(drop=True)
                df['ExcelRow'] = df.index + header_row_index + 2
                df = df[~df['ExcelRow'].isin(bad_row_numbers)]
                df = df.drop(columns=['ExcelRow'])

                # หาคอลัมน์ที่เกี่ยวข้อง
                term_col, lang_col, spec_col = fuzzy_find_columns(df)
                print(f"🔎 ใช้คอลัมน์ Term: {term_col}, Language: {lang_col}, Spec: {spec_col}")

                df = df.rename(columns={term_col: "Term (Text)"})
                if lang_col:
                    df = df.rename(columns={lang_col: "Language"})
                    df["Language List"] = df["Language"].apply(lambda x: str(x).split(",") if pd.notna(x) else [])
                else:
                    df["Language List"] = [[] for _ in range(len(df))]

                # ข้ามแถวที่ไม่ต้องตรวจ (Spec เป็น None, N/A, -)
                if spec_col:
                    df = df[~df[spec_col].astype(str).str.strip().str.upper().isin(['N/A', 'NONE', '-'])]

                # แยกหลายคำใน 1 เซลล์ออกเป็นหลายบรรทัด (ถ้ามี)
                df = df.dropna(subset=["Term (Text)"])
                df = df.copy()
                df["Term (Text)"] = df["Term (Text)"].astype(str)
                df = df.explode("Term (Text)", ignore_index=True) if df["Term (Text)"].apply(lambda x: "\n" in x or "," in x).any() else df

                # ตรวจว่า cell มีสีแดงและขีดเส้นทับหรือไม่
                def is_strikethrough_red(cell):
                    font = cell.font
                    color = font.color
                    return font.strike and color and color.type == "rgb" and color.rgb.startswith("FFFF0000")

                # ตรวจสอบ Remark และแยกภาษาออกมา
                if "Remark" in df.columns:
                    def extract_languages_from_remark(remark, term):
                        langs = []
                        if pd.isna(remark): return langs
                        lines = str(remark).splitlines()
                        for line in lines:
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

def normalize_text(text):
    """ Normalize text for comparison """
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize('NFKC', text)
    text = text.replace("\u00A0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()


from collections import defaultdict
import pandas as pd
import re
import unicodedata


def normalize_text(text):
    """ Normalize text for comparison """
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize('NFKC', text)
    text = text.replace("\u00A0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()

def start_check(df_checklist, extracted_text_list):
    from collections import defaultdict
    import pandas as pd

    results = []
    all_texts = []

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

    grouped = defaultdict(list)
    last_valid_requirement = None

    for _, row in df_checklist.iterrows():
        requirement = str(row.get("Requirement", "")).strip()
        term = str(row.get("Term (Text)", "")).strip()
        spec = str(row.get("Specification", "")).strip()

        # ใช้ Requirement ล่าสุดถ้า cell ปัจจุบันว่าง
        if requirement:
            last_valid_requirement = requirement
        else:
            requirement = last_valid_requirement

        if not term or term.upper() in ['N/A', 'NONE', '-', 'UNSPECIFIED']:
            continue

        # Manual check detection
        manual_keywords = [
            "logo", "symbol", "icon", "mark", "graphic", "trademark", "logotype", "emblem",
            "artwork", "brandmark", "warning", "statement", "block", "list of content", "pictorial", "pictogram",
        ]

        if any(kw in requirement.lower() for kw in manual_keywords):
            grouped[(requirement, spec, "Manual")].append({
                "Term": term,
                "Found": "-",
                "Match": "-",
                "Page(s)": "-",
                "Font Size": "-",
                "Note": "Manual check required",
                "Verification": "Manual"
            })
            continue

        if any(x in term.upper() for x in ["PLEASE REFER", "REMARK ONLY", "SEE TEMPLATE"]):
            continue

        spec = spec.replace("Unspecified", "-") or "-"
        target_term = normalize_text(term)

        found_pages = []
        matched_items = []

        for text_norm, page_number, item in all_texts:
            if text_norm == target_term:
                found_pages.append(page_number)
                matched_items.append(item)

        found_flag = "✅ Found" if found_pages else "❌ Not Found"

        if found_pages:
            item = matched_items[0]
            size_mm = float(item.get("size", 0))
            bold = item.get("bold", False)
            underline = item.get("underline", False)
            text = item.get("text", "")

            match_result = "✔"
            notes = []

            if "bold" in spec.lower() and not bold:
                match_result = "❌"
                notes.append("Not Bold")
            if "underline" in spec.lower() and not underline:
                match_result = "❌"
                notes.append("Underline missing")
            if "all caps" in spec.lower() and not text.isupper():
                match_result = "❌"
                notes.append("Not All Caps")
            if "≥" in spec:
                try:
                    threshold = float(spec.replace("≥", "").replace("mm", "").strip())
                    if size_mm < threshold:
                        match_result = "❌"
                        notes.append(f"Font < {threshold} mm")
                except:
                    pass

            pages = sorted(set(found_pages))
            all_artwork_pages = sorted(page_mapping.values())
            page_str = "All pages" if set(pages) == set(all_artwork_pages) else ", ".join(str(p) for p in pages)

            grouped[(requirement, spec, "Auto")].append({
                "Term": term,
                "Found": found_flag,
                "Match": match_result,
                "Page(s)": page_str,
                "Font Size": "✔" if match_result == "✔" else f"{round(size_mm, 2)} mm",
                "Note": ", ".join(notes) if notes else "-",
                "Verification": "Auto"
            })
        else:
            grouped[(requirement, spec, "Auto")].append({
                "Term": term,
                "Found": found_flag,
                "Match": "❌",
                "Page(s)": "-",
                "Font Size": "-",
                "Note": "-",
                "Verification": "Auto"
            })

    # รวมกลุ่ม Term ที่มี Requirement เดียวกันให้อยู่ใน row เดียว
    final_results = []
    for (requirement, spec, verification), items in grouped.items():
        combined_term = "\n".join(item["Term"] for item in items)
        combined_found = "\n".join(item["Found"] for item in items)
        combined_match = "\n".join(item["Match"] for item in items)
        combined_pages = "\n".join(item["Page(s)"] for item in items)
        combined_font = "\n".join(item["Font Size"] for item in items)
        combined_note = "\n".join(item["Note"] for item in items)

        final_results.append({
            "Requirement": requirement,
            "Term": combined_term,
            "Specification": spec,
            "Found": combined_found,
            "Match": combined_match,
            "Page(s)": combined_pages,
            "Font Size": combined_font,
            "Note": combined_note,
            "Verification": verification
        })

    return pd.DataFrame(final_results)
