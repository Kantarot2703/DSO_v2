import os, io, re
import pandas as pd
import logging
import unicodedata
from PyQt5.QtCore import Qt
from openpyxl import load_workbook
from openpyxl.styles.colors import Color
from collections import defaultdict


# Allowed part codes from PDF filenames
ALLOWED_PART_CODES = ['UU1_DOM', 'DOM', 'UU1', '2LB', '2XV', '4LB', '19L', '19A', '21A', 'DC1']

def normalize_headers(df):
    rename = {}
    for col in df.columns:
        name = str(col).strip().lower()
        # Requirement
        if ("require" in name) or ("ข้อกำหนด" in name) or ("หัวข้อ" in name):
            rename[col] = "Requirement"
            continue
        # Term
        if ("symbol" in name) or ("exact" in name) or ("term" in name):
            rename[col] = "Symbol/ Exact wording"
            continue
        # Spec
        if ("spec" in name) or ("specification" in name):
            rename[col] = "Specification"
            continue

        # Package Panel
        if ("package panel" in name) or ("package" in name and "panel" in name):
            rename[col] = "Package Panel"
            continue

        # Procedure
        if ("procedure" in name) or ("process" in name):
            rename[col] = "Procedure"
            continue

        # Remark
        if ("remark" in name) or ("หมายเหตุ" in name):
            rename[col] = "Remark"
            continue

    df = df.rename(columns=rename)

    if "Requirement" not in df.columns:
        raise ValueError("Checklist Excel must contain a column recognizable as 'Requirement'.")
    return df

def get_strikeout_or_red_text_rows(excel_path, sheet_name, header_row_index):
    wb = load_workbook(excel_path)
    ws = wb[sheet_name]
    bad_rows = set()

    for row in ws.iter_rows(min_row=header_row_index + 2):  
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

    logging.info(f"🧾 Columns found in sheet: {list(df.columns)}")

    # add Fallback mapping ถ้า fuzzy หาไม่เจอ
    def _norm(s: object) -> str:
         return str(s).replace("\xa0", " ").strip().lower()

    # ตรง Keyword ชัดเจน 
    for col in df.columns:
        if pd.isna(col): 
            continue
        n = _norm(col)
        n_compact = n.replace(" ", "")

        if term_col is None and any(k in n_compact for k in ["term", "ข้อความ", "exactwording", "symbol", "wording"]):
            term_col = col

        if lang_col is None and any(k in n_compact for k in ["languagecode", "langcode", "language", "lang", "ภาษา"]):
            lang_col = col

        if spec_col is None and any(k in n_compact for k in ["specification", "spec", "requirement", "ข้อกำหนด"]):
            spec_col = col

    # Fallback ผ่อนเงื่อไข
    if term_col is None:
        for col in df.columns:
            n = _norm(col)
            n_compact = n.replace(" ", "")
            if any(k in n_compact for k in ["term", "symbol", "exact", "wording"]):
                term_col = col
                break

    if lang_col is None:
        for col in df.columns:
            n = _norm(col)
            if ("language" in n) or (n == "lang") or ("language code" in n) or ("lang code" in n):
                lang_col = col
                break

    if spec_col is None:
        for col in df.columns:
            n = _norm(col)
            if (n == "spec") or ("specification" in n) or ("requirement" in n) or ("ข้อกำหนด" in n):
                spec_col = col
                break

    # OPTIONAL บังคับตั้งชื่อคอลัมน์ Requirement ถ้าตรวจพบ
    for c in df.columns:
        if str(c).strip().lower() == "requirement":
            if c != "Requirement":
                df = df.rename(columns={c: "Requirement"})
            break

    logging.info(f"🛟 Fallback columns → Term: {term_col}, Language: {lang_col}, Spec: {spec_col}")
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
                logging.info(f"✅ Found matching sheet: {sheet_name}")
                df = all_sheets[sheet_name]

                HEADER_HINTS = [
                    r"\brequirement\b",
                    r"\blanguage\b|\blang\b|\blanguage\s*code\b",
                    r"\bsymbol\b|\bexact\s*wording\b|\bterm\b",
                    r"\bspec(ification)?\b",
                ]

                def _row_score(cells):
                    score = 0
                    for c in cells:
                        s = "" if c is None else str(c).strip()
                        s_norm = s.lower()
                        if not s:
                            continue

                        # ให้แต้มถ้าตรงคีย์เวิร์ดหัวตาราง
                        if any(re.search(p, s_norm) for p in HEADER_HINTS):
                            score += 5
                        # ไม่ใช่ประโยคยาว/มีสัญลักษณ์ = เยอะ
                        if len(s) <= 24:
                            score += 1
                        if "=" in s or "“" in s or "”" in s:
                            score -= 2
                    return score
                
                # ค้นหา header ภายใน 15 แถวแรก เลือกแถวที่ได้คะแนนมากสุด
                header_row_index = None
                best_score = -10
                scan_upto = min(15, len(df))
                for i in range(scan_upto):
                    row_vals = list(df.iloc[i].values)
                    if pd.Series(row_vals).notna().sum() < 2:
                        continue
                    sc = _row_score(row_vals)
                    if sc > best_score:
                        best_score = sc
                        header_row_index = i

                if header_row_index is None:
                    raise ValueError(f"❌ ไม่พบแถว header ที่เหมาะสมใน sheet: {sheet_name}")
                
                df.columns = df.iloc[header_row_index]
                df = df[header_row_index + 1:].reset_index(drop=True)
                logging.info(f"🧾 Header chosen at row {header_row_index+1} | columns: {list(df.columns)[:6]}...")

                # Column Mapping
                term_col, lang_col, spec_col = fuzzy_find_columns(df)
                logging.info(f"🔎 ใช้คอลัมน์ Term: {term_col}, Language: {lang_col}, Spec: {spec_col}")

                # Standardize
                if spec_col and spec_col in df.columns:
                    df[spec_col] = df[spec_col].apply(
                        lambda x: "-" if pd.isna(x) or str(x).strip().upper() in ["N/A", "NONE", "-"] else str(x)
                    )

                # ffill เฉพาะคอลัมน์ที่มีอยู่จริง
                columns_to_ffill = [c for c in df.columns if str(c).strip().lower() in ["requirement", "language"]]
                if columns_to_ffill:
                    df[columns_to_ffill] = df[columns_to_ffill].ffill()

                # Rename ถ้า term_col ไม่อยู่ ให้สร้างคอลัมน์ว่างกัน KeyError
                if term_col in df.columns:
                    df = df.rename(columns={term_col: "Symbol/Exact wording"})
                elif "Symbol/Exact wording" not in df.columns:
                    df["Symbol/Exact wording"] = "-"

                GROUP_RE = re.compile(r"^\s*\[GROUP:\s*(?P<name>.+?)\s*\]\s*\[(?P<mode>ANY|ALL)\]\s*$", re.IGNORECASE)

                def _split_simple_list(cell: str):
                    """รองรับหลายพาธคั่นด้วย ; | หรือขึ้นบรรทัดใหม่"""
                    if not isinstance(cell, str):
                        return []
                    s = cell.strip()
                    if not s or s in ["-", "N/A", "None"]:
                        return []
                    parts = re.split(r"[;\n|]", s.replace("\r", ""))
                    return [p.strip().replace("\\", "/") for p in parts if p.strip()]
                
                def _parse_image_groups(cell: str):
                    """
                    รูปแบบที่รองรับ:
                    - แบบมี group/tag:
                        [GROUP: Old logo][ALL]
                        //server/share/old1.png
                        //server/share/old2.png
                        [GROUP: New logo][ANY]
                        assets/new1.png
                        assets/new2.png
                    - แบบธรรมดา: หลายพาธในเซลล์เดียว -> กลุ่มเดียว mode=ANY
                    - สมมติว่ามีคอลัมน์ Image Match แยก: จะไป normalize ต่อ
                    """
                    if not isinstance(cell, str) or not cell.strip():
                        return []
                    lines = [ln.strip() for ln in cell.replace("\r", "").split("\n")]
                    groups, cur = [], None
                    for ln in lines:
                        if not ln:
                            continue
                        m = GROUP_RE.match(ln)
                        if m:
                            if cur and cur["paths"]:
                                groups.append(cur)
                            cur = {"name": m.group("name"), "mode": m.group("mode").lower(), "paths": []}
                        else:
                            for p in re.split(r"[;|]", ln):
                                p = p.strip()
                                if p and p not in ["-", "N/A", "None"]:
                                    if cur is None:
                                        cur = {"name": "", "mode": "any", "paths": []}
                                    cur["paths"].append(p.replace("\\", "/"))
                    if cur and cur["paths"]:
                        groups.append(cur)

                    # ถ้าไม่มี [GROUP:...] เลย และมีพาธเดียว/หลายพาธ -> กลุ่มเดียว ANY 
                    if not groups:
                        paths = _split_simple_list(cell)
                        return [{"name": "", "mode": "any", "paths": paths}] if paths else []
                    return groups
                
                def _flatten_paths(groups):
                    out = []
                    for g in (groups or []):
                        out.extend(g.get("paths", []))
                    return out 
                
                # สร้างคอลัมน์ Image_Groups + Image_Paths_Flat
                if "Image Path" in df.columns:
                    df["Image_Groups"] = df["Image Path"].fillna("").astype(str).apply(_parse_image_groups)
                    df["Image_Paths_Flat"] = df["Image_Groups"].apply(_flatten_paths)
                else:
                    df["Image_Groups"] = [[] for _ in range(len(df))]
                    df["Image_Paths_Flat"] = [[] for _ in range(len(df))]

                # หากมีคอลัมน์ Image Match แยก (ANY/ALL) ให้บังคับโหมดกลุ่มเดี่ยวให้ตรงค่านี้
                if "Image Match" in df.columns:
                    def _apply_global_mode(groups, mode_cell):
                        mode = (str(mode_cell).strip().lower() if isinstance(mode_cell, str) else "")
                        if mode in ["all", "any"]:
                            # ถ้ามีหลายกลุ่ม จะ apply ให้ทุกกลุ่ม
                            for g in groups or []:
                                g["mode"] = mode
                        return groups
                    df["Image_Groups"] = [
                        _apply_global_mode(g, m) for g, m in zip(df["Image_Groups"], df["Image Match"])
                    ] 

                # Add extract hyperlink targets from Remark cells
                try:
                    wb = load_workbook(excel_path, data_only=True)
                    ws = wb[sheet_name]
                    # หา index ของคอลัมน์ Remark จาก df.columns
                    if "Remark" in df.columns:
                        # ตำแหน่งคอลัมน์ในชีต: +1 เพราะ openpyxl เริ่มที่ 1
                        remark_col_idx = list(df.columns).index("Remark") + 1
                        start_row = header_row_index + 2  # แถวแรกของข้อมูลจริงในชีต

                        remark_links = []
                        for i in range(len(df)):
                            cell = ws.cell(row=start_row + i, column=remark_col_idx)
                            url = ""
                            if cell and cell.hyperlink:
                                # openpyxl ให้ .hyperlink.target เป็น URL จริง
                                url = cell.hyperlink.target or ""
                            remark_links.append(url)
                        df["Remark Link"] = remark_links
                    else:
                        df["Remark Link"] = ""
                except Exception as _e:
                    logging.debug(f"Remark hyperlink extraction skipped: {_e}")
                    df["Remark Link"] = ""

                # --- RESOLVE absolute paths for Image_Groups and add _HasImage BEFORE filtering ---

                excel_dir = os.path.dirname(excel_path)

                def _resolve_group_paths(groups):
                    out = []
                    for g in (groups or []):
                        paths = []
                        for p in g.get("paths", []):
                            if not isinstance(p, str) or not p.strip():
                                continue
                            p2 = p.strip().replace("\\", os.sep).replace("/", os.sep)
                            if not os.path.isabs(p2):
                                p2 = os.path.abspath(os.path.join(excel_dir, p2))
                            paths.append(p2)
                        out.append({"name": g.get("name",""), "mode": (g.get("mode") or "any").lower(), "paths": paths})
                    return out

                df["Image_Groups_Resolved"] = df["Image_Groups"].apply(_resolve_group_paths)

                def _has_any_image(groups):
                    if not groups:
                        return False
                    for g in groups:
                        for p in g.get("paths", []):
                            if isinstance(p, str) and p.strip():
                                return True
                    return False

                df["_HasImage"] = df["Image_Groups_Resolved"].apply(_has_any_image)

                # (ถ้ามีคอลัมน์ Image Path แบบเดี่ยว ก็ resolve ให้ด้วย)
                if "Image Path" in df.columns:
                    df["Image Path"] = df["Image Path"].fillna("").astype(str)
                else:
                    df["Image Path"] = ""

                def _resolve_path(p):
                    if not isinstance(p, str) or not p.strip():
                        return ""
                    p = p.strip().replace("\\", os.sep).replace("/", os.sep)
                    if os.path.isabs(p):
                        return p
                    return os.path.abspath(os.path.join(excel_dir, p))

                df["Image Path Resolved"] = df["Image Path"].apply(_resolve_path)
                
                # ตัดแถวว่าง/แถวผีหลัง
                def _clean(s):
                    return str(s).strip().lower()

                # ชื่อคอลัมน์สำคัญ (บางไฟล์อาจไม่มี spec_col)
                term_col_safe = "Symbol/Exact wording"
                spec_col_safe = spec_col if spec_col in df.columns else None

                # เงื่อนไขว่าง
                term_empty  = df[term_col_safe].astype(str).str.strip().isin(["", "-", "nan", "none", "n/a"])
                if spec_col_safe:
                    spec_empty  = df[spec_col_safe].astype(str).str.strip().isin(["", "-", "nan", "none", "n/a"])
                else:
                    spec_empty  = pd.Series([True] * len(df), index=df.index) # Series ของ True ยาวเท่า df (ให้ผ่านเงื่อนไขนี้ไป)

                # สร้าง Series ว่างสำหรับ Remark ถ้าไม่มีคอลัมน์
                remark_series = df["Remark"] if "Remark" in df.columns else pd.Series([""] * len(df), index=df.index)
                remark_empty = remark_series.astype(str).str.strip().isin(["", "-", "nan"])

                # Add ลิงก์ Remark ก็ถือว่า "มีข้อมูล"
                remark_link_series = df.get("Remark Link", pd.Series([""] * len(df), index=df.index))
                remark_link_empty = remark_link_series.astype(str).str.strip().isin(["", "-", "nan", ""])

                # แถวที่ควรเก็บ = อย่างน้อยต้องมี Term หรือ Spec หรือ Remark ไม่ว่าง
                keep_mask = ~(term_empty & spec_empty & remark_empty & remark_link_empty) | df["_HasImage"]
                df = df[keep_mask].reset_index(drop=True)

                # กันช่องว่างล้วน (ยกเว้น Requirement/Language) เป็น NaN ล้วน
                non_struct_cols = [c for c in df.columns if str(c).strip().lower() not in ["requirement", "language"]]
                df = df[~df[non_struct_cols].isna().all(axis=1)].reset_index(drop=True)

                # Drop Red+Strike Rows (คง log ไว้ แต่ให้เป็น debug)
                bad_row_numbers = get_strikeout_or_red_text_rows(excel_path, sheet_name, header_row_index)
                logging.debug(f"❌ Red+Strike rows from Excel: {bad_row_numbers}")

                df["ExcelRow"] = df.index + header_row_index + 2 
                df.drop(columns=["ExcelRow"], inplace=True)

                # Drop columns ที่ไม่ใช้
                columns_to_exclude = ["Instruction of Play function feature", "Warning statement"]
                df = df[[col for col in df.columns if col and str(col).strip().lower() not in columns_to_exclude]]

                # Explode multi-term 
                df["Symbol/Exact wording"] = df["Symbol/Exact wording"].astype(str).str.replace("\r", "")
                df["Symbol/Exact wording"] = df["Symbol/Exact wording"].apply(
                    lambda s: [p.strip() for p in s.split("\n") if p.strip()] if ("\n" in s) else [s.strip()]
                )
                df = df.explode("Symbol/Exact wording", ignore_index=True)

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
                        for remark, term in zip(df.get("Remark", []), df["Symbol/Exact wording"])
                    ]

                return df 

    raise ValueError("❌ ไม่พบ Sheet ที่ตรงกับ Part code จากชื่อไฟล์ PDF")

def start_check(df_checklist, extracted_text_list):

    logger = logging.getLogger(__name__)

    results = []
    all_texts = []
    grouped = defaultdict(list)

    skip_keywords = [
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

    logger.info(
    "📄 Artwork-like pages considered: %d (real pages: %s)",
    len(artwork_pages),
    list(page_mapping.values())
    )

    for artwork_index, page_items in enumerate(artwork_pages):
        for item in page_items:
            text_norm = normalize_text(item.get("text", ""))
            page_number = page_mapping[artwork_index + 1]
            all_texts.append((text_norm, page_number, item))

    for idx, row in df_checklist.iterrows():
        requirement = str(row.get("Requirement", "")).strip()
        spec = str(row.get("Specification", "")).strip()
        package_panel = (str(row.get("Package Panel", "")) or "").strip() or "-"
        procedure = (str(row.get("Procedure", "")) or "").strip() or "-"
        remark_text = (str(row.get("Remark", "")) or "").strip()
        remark_link = (str(row.get("Remark Link", "")) or "").strip()

        # Normalize
        req_norm = normalize_text(requirement)
        spec = "-" if spec.lower() in ["", "n/a", "none", "unspecified", "nan"] else spec
        spec_norm = normalize_text(spec)

        # Term (ไม่ดึงค่าจากแถวอื่น)
        term_raw = row.get("Symbol/Exact wording", None)
        term_cell_raw = str(term_raw) if pd.notna(term_raw) else ""
        term_cell_clean = term_cell_raw.strip()
        term_cell_clean = "-" if term_cell_clean.lower() in ["", "n/a", "none", "unspecified", "nan"] else term_cell_clean
        term_lines = [term_cell_clean] if term_cell_clean != "-" else []

        # ตรวจ Manual
        is_manual = any(kw in req_norm for kw in manual_keywords)

        # ข้าม row ไม่จำเป็น
        if any(kw in req_norm for kw in skip_keywords) or any(kw in spec_norm for kw in skip_keywords):
            continue

        # VERIFIED SECTION
        if not is_manual:
            if not term_lines:
                has_image = bool(row.get("_HasImage", False))
                if has_image:
                    grouped[(requirement, spec, "Manual")].append({
                        "Term": "-",
                        "Remark": remark_text or "-",
                        "Remark URL": remark_link or "-",
                        "Found": "❌ Not Found",
                        "Match": "❌",
                        "Pages": "-",
                        "Font Size": "-",
                        "Note": "Image mapping only (no text term)",
                        "Package Panel": package_panel,
                        "Procedure": procedure,
                    })
                    continue 
                continue

        # --- MANUAL SECTION ---
        if is_manual:
            remark_str = str(row.get("Remark", "")).strip().lower()
            is_term_empty   = (len(term_lines) == 0)
            is_spec_empty   = (str(spec)).strip().lower() in ["", "-", "nan", "none", "null"]
            is_remark_empty = (remark_str in ["", "-", "nan"])

            if is_term_empty and is_spec_empty and is_remark_empty:
                if bool(row.get("_HasImage", False)):
                    grouped[(requirement, spec, "Manual")].append({
                        "Term": "-",
                        "Remark": remark_text or "-",
                        "Remark URL": remark_link or "-",
                        "Found": "❌ Not Found",
                        "Match": "❌",
                        "Pages": "-",
                        "Font Size": "-",
                        "Note": "Image mapping only (manual; empty text fields)",
                        "Package Panel": package_panel,
                        "Procedure": procedure,
                    })
                    continue
                else:
                    continue

            logger.info(
                "🟨 [MANUAL] Req: '%s' | Spec: '%s' → Manual verification",
            requirement, (spec or "-")
            )

            if not term_lines:
                grouped[(requirement, spec, "Manual")].append({
                    "Term": term_cell_raw,
                    "Found": "-",
                    "Match": "-",
                    "Pages": "-",
                    "Font Size": "-",
                    "Note": "Manual check required",
                    "Verification": "Manual",
                    "Remark": remark_text,
                    "Remark URL": remark_link,
                    "Package Panel": package_panel,
                    "Procedure": procedure,
                })
            else:
                for term in term_lines:
                    grouped[(requirement, spec, "Manual")].append({
                    "Term": term,
                    "Found": "-",
                    "Match": "-",
                    "Pages": "-",
                    "Font Size": "-",
                    "Note": "Manual check required",
                    "Verification": "Manual",
                    "Remark": remark_text,
                    "Remark URL": remark_link,
                    "Package Panel": package_panel,
                    "Procedure": procedure,
                    })
            continue

        # VERIFIED SECTION
        for term in term_lines:
            term_norm = normalize_text(term)
            if not term_norm or any(kw in term_norm for kw in skip_keywords):
                continue

            # รองรับ "or" ใน term เช่น "DOM or UU1"
            sep = " or "
            term_variants = [t.strip() for t in term_norm.split(sep)] if sep in term_norm else [term_norm]

            found_pages = []
            matched_items = []
            seen = set () # กันซ้ำด้วย (page_number, id(item))
            
            # OR loop
            for variant in term_variants:
                if not variant:
                    continue
                for text_norm, page_number, item in all_texts:
                    if variant in text_norm:
                        key = (page_number, id(item))
                        if key not in seen:
                            seen.add(key)
                            found_pages.append(page_number)
                            matched_items.append(item)

            if not found_pages:
                term_words = term_norm.split()
                for text_norm, page_number, item in all_texts:
                    if any(w in text_norm for w in term_words):
                        key = (page_number, id(item))
                        if key not in seen:
                            seen.add(key)
                            found_pages.append(page_number)
                            matched_items.append(item)

            found_flag = "✅ Found" if found_pages else "❌ Not Found"
            match_result = "✔"
            notes = []

            # กัน list ว่าง/แปลงชนิด
            spec_lower = spec.lower() if isinstance(spec, str) else "-"
            sizes = [float(i.get("size", 0) or 0) for i in matched_items]
            max_size = max(sizes) if sizes else 0.0
            bolds = [bool(i.get("bold", False)) for i in matched_items]
            underlines = [bool(i.get("underline", False)) for i in matched_items]
            texts = [i.get("text", "") for i in matched_items]

            font_size_str = "-"

            if found_pages and matched_items and spec != "-":
                if "bold" in spec_lower and not any(bolds):
                    match_result = "❌"
                    notes.append("Not Bold")

                if "no underline" in spec_lower:
                    if any(underlines):
                        match_result = "❌"
                        notes.append("Underline must be absent")
                elif "underline" in spec_lower:
                    if not any(underlines):
                        match_result = "❌"
                        notes.append("Underline Missing")

                if "all caps" in spec_lower and not any(t.isupper() for t in texts):
                    match_result = "❌"
                    notes.append("Not All Caps")

                if "≥" in spec:
                    try:
                        m = re.search(r"≥\s*(\d+(?:\.\d+)?)", spec)
                        if m:
                            threshold = float(m.group(1))
                            if max_size < threshold:
                                match_result = "❌"
                                notes.append(f"Font < {threshold} mm")
                            # โชว์ค่าจริงที่ตรวจเจอ (max) กัน crash index 0
                            font_size_str = "✔" if match_result == "✔" else f"{round(max_size, 2)} mm"
                        else:
                            font_size_str = "-"
                    except Exception:
                        font_size_str = "-"
                else:
                    font_size_str = "✔" if match_result == "✔" else "-"
            else:
                font_size_str = "-"

            pages = sorted(set(found_pages))
            all_pages = sorted(set(p for _, p, _ in all_texts))
            page_str = "All pages" if set(pages) == set(all_pages) else ", ".join(str(p) for p in pages)

            if found_pages:
                logger.info("✅ [FOUND] Req: '%s' | Term: '%s' | Pages: %s | Match: %s | Font: %s | Notes: %s",
                            requirement, term, (page_str or "-"),
                            match_result, font_size_str, (", ".join(notes) or "-"))
            else:
                logger.warning("❌ [NOT FOUND] Req: '%s' | Term tried: '%s'", requirement, term)
            
            grouped[(requirement, spec, "Verified")].append({
                "Term": term,
                "Found": found_flag,
                "Match": match_result if found_flag == "✅ Found" else "❌",
                "Pages": page_str if found_flag == "✅ Found" else "-",
                "Font Size": font_size_str if found_flag == "✅ Found" else "-",
                "Note": ", ".join(notes) if notes else "-",
                "Verification": "Verified",
                "Remark": remark_text,
                "Remark URL": remark_link,
                "Package Panel": package_panel,   
                "Procedure": procedure,           
            })

    final_results = []
    for (requirement, spec, verification), items in grouped.items():
        for item in items:
            term_display = item.get("Term", "")
            if pd.isna(term_display) or str(term_display).strip().lower() in ["", "nan"]:
                term_display = "-"
            final_results.append({
                "Requirement": requirement,
                "Symbol/ Exact wording": term_display,
                "Specification": spec,
                "Package Panel": item.get("Package Panel", "-"),
                "Procedure": item.get("Procedure", "-"),
                "Remark": item.get("Remark", "-"),
                "Remark URL": item.get("Remark URL", "-"), 
                "Found": item["Found"],
                "Match": item["Match"],
                "Pages": item["Pages"],
                "Font Size": item["Font Size"],
                "Note": item["Note"],
                "Verification": verification
            })

    return pd.DataFrame(final_results)