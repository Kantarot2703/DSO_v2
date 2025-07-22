import os
import re
import pandas as pd

# Allowed part codes from PDF filenames
ALLOWED_PART_CODES = ['UU1_DOM', '2LB', '2XV', '4LB', '19L', '19A', '21A', 'DC1']

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

                return df

    raise ValueError("❌ ไม่พบ Sheet ที่ตรงกับ Part code จากชื่อไฟล์ PDF")
