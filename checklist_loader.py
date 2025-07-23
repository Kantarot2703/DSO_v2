import os
import re
import pandas as pd
import unicodedata
from openpyxl import load_workbook
from openpyxl.styles.colors import Color


# Allowed part codes from PDF filenames
ALLOWED_PART_CODES = ['UU1_DOM', '2LB', '2XV', '4LB', '19L', '19A', '21A', 'DC1']

def get_strikeout_or_red_text_rows(excel_path, sheet_name, header_row_index):
    wb = load_workbook(excel_path)
    sheet = wb[sheet_name]
    strike_rows = set()

    for row in sheet.iter_rows(min_row=header_row_index+2):
        for cell in row:
            if cell.font is not None:
                is_red = cell.font.color and cell.font.color.type == "rgb" and cell.font.color.rgb.startswith("FFFF0000")
                is_strike = cell.font.strike
                if is_red and is_strike:
                    strike_rows.add(cell.row - (header_row_index+2))  
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

                # ตรวจว่า cell มีสีแดงและขีดเส้นทับหรือไม่ โดยใช้ openpyxl
                def is_strikethrough_red(cell):
                    font = cell.font
                    color = font.color
                    if font.strike and color and (color.type == "rgb" and color.rgb in ["FFFF0000", "FF0000"]):
                        return True
                    return False

                # โหลดไฟล์ Excel แบบ openpyxl เพื่อตรวจ font
                wb = load_workbook(excel_path, data_only=True)
                ws = wb[sheet_name]
                rows_to_exclude = set()
                for row in ws.iter_rows(min_row=header_row_index+2):
                    for cell in row:
                        if is_strikethrough_red(cell):
                            rows_to_exclude.add(cell.row)
                            break 

                # แปลง Excel row number -> pandas index
                df = df.reset_index(drop=True)
                df['ExcelRow'] = df.index + header_row_index + 2
                df = df[~df['ExcelRow'].isin(rows_to_exclude)]
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

                # ข้ามแถวที่มีข้อความสีแดง + ขีดเส้นทับ
                strike_rows = get_strikeout_or_red_text_rows(excel_path, sheet_name, header_row_index)
                if strike_rows:
                    df = df[~df.index.isin(strike_rows)]
                    df.reset_index(drop=True, inplace=True)

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


def start_check(df_checklist, extracted_text_list):
    results = []
    all_texts = []
    for page_index, page_items in enumerate(extracted_text_list):
        for item in page_items:
            text_norm = normalize_text(item.get("text", ""))
            all_texts.append((text_norm, page_index + 1, item))

    for _, row in df_checklist.iterrows():
        requirement = str(row.get("Requirement", "")).strip()
        term = str(row.get("Term (Text)", "")).strip()
        spec = str(row.get("Specification", "")).strip()

        # ข้าม term ที่ว่าง หรือคำอธิบายที่ไม่ต้องตรวจ
        if not term or term.upper() in ['N/A', 'NONE', '-', 'UNSPECIFIED']:
            continue
        if any(x in term.upper() for x in ["PLEASE REFER", "REMARK ONLY", "SEE TEMPLATE", "LEGAL ADDRESS"]):
            continue

        target_term = normalize_text(term)

        found_pages = []
        matched_items = []

        for text_norm, page_number, item in all_texts:
            if text_norm == target_term:
                found_pages.append(str(page_number))
                matched_items.append(item)

        if found_pages:
            font_size = round(float(matched_items[0].get("size", 0)), 2)
            results.append({
                "Requirement": requirement,
                "Term": term,
                "Specification": spec.replace("...", "").strip(),
                "Match Result": "✅ พบแล้ว",
                "Page(s)": ", ".join(sorted(set(found_pages))),
                "Font Size": f"{font_size} pt"
            })
        else:
            results.append({
                "Requirement": requirement,
                "Term": term,
                "Specification": spec.replace("...", "").strip(),
                "Match Result": "❌ ไม่พบ",
                "Page(s)": "-",
                "Font Size": "-"
            })

    return pd.DataFrame(results)