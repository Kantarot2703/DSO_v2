import os, re
import pandas as pd
import logging
from PyQt5.QtWidgets import QWidget, QLabel, QHBoxLayout, QVBoxLayout, QGridLayout, QHeaderView
from PyQt5.QtCore import Qt
from PyQt5 import QtWidgets, QtGui, QtCore
from ui.pdf_viewer import PDFViewer
from checklist_loader import load_checklist, start_check
from pdf_reader import extract_text_by_page
from checker import check_term_in_page
from result_exporter import export_result_to_excel
from PyQt5.QtGui import QColor, QIcon, QPixmap
from pdf_reader import extract_product_info_by_page
from collections import defaultdict


APP_ICON_PATH = os.path.join("assets", "app", "dso_icon.ico")

class DSOApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        try:
            self.setWindowIcon(QIcon(APP_ICON_PATH))
        except Exception:
            pass

        self.setWindowTitle("DSO - Digital Sign Off")
        self.setGeometry(100, 100, 1200, 800)

        self.excel_path = ""
        self.pdf_path = ""
        self.checklist_df = None
        self.pages = None
        self.result_df = None
        self.product_infos = []
        self._image_cache = {}

        self.init_ui()

    def init_ui(self):
        shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+F"), self)
        shortcut.activated.connect(self.search_text)
        layout = QtWidgets.QVBoxLayout()

        # Upload buttons
        file_layout = QtWidgets.QHBoxLayout()
        self.pdf_btn = QtWidgets.QPushButton("📄 Upload PDF (Artwork)")
        self.excel_btn = QtWidgets.QPushButton("📋 Upload Excel (Checklist)")
        self.pdf_label = QtWidgets.QLabel("PDF: Not selected")
        self.excel_label = QtWidgets.QLabel("Checklist: Not selected")
        self.pdf_btn.clicked.connect(self.load_pdf)
        self.excel_btn.clicked.connect(self.load_excel)
        file_layout.addWidget(self.pdf_btn)
        file_layout.addWidget(self.excel_btn)
        file_layout.addStretch()

        # Search bar
        search_layout = QtWidgets.QHBoxLayout()
        self.search_input = QtWidgets.QLineEdit()
        self.search_input.setPlaceholderText("Search term ...")
        self.search_btn = QtWidgets.QPushButton("Search")
        self.search_btn.clicked.connect(self.search_text)
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.search_btn)

        # Result Table
        self.result_table = QtWidgets.QTableWidget()
        self.result_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.result_table.setColumnCount(0)
        self.result_table.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.result_table.horizontalHeader().setStretchLastSection(False)
        self.result_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        
        self.hovered_row = -1
        self.result_table.setMouseTracking(True)
        self.result_table.viewport().installEventFilter(self)
        self.result_table.itemSelectionChanged.connect(self.update_row_highlight)
        self.result_table.horizontalHeader().sectionResized.connect(self._on_column_resized)

        # Action Buttons
        action_layout = QtWidgets.QHBoxLayout()
        self.check_btn = QtWidgets.QPushButton("Start Checking")
        self.export_btn = QtWidgets.QPushButton("Export to Excel")
        self.preview_btn = QtWidgets.QPushButton("Preview PDF")
        self.check_btn.clicked.connect(self.start_checking)
        self.export_btn.clicked.connect(self.export_results)
        self.preview_btn.clicked.connect(self.preview_pdf)
        action_layout.addWidget(self.check_btn)
        action_layout.addWidget(self.export_btn)
        action_layout.addWidget(self.preview_btn)
        action_layout.addStretch()

        # Add to layout
        layout.addLayout(file_layout)
        layout.addWidget(self.pdf_label)
        layout.addWidget(self.excel_label)
        layout.addLayout(search_layout)
        layout.addWidget(self.result_table)
        layout.addLayout(action_layout)
        self.setLayout(layout)

    def eventFilter(self, source, event):
        if source == self.result_table.viewport() and event.type() == QtCore.QEvent.MouseMove:
            index = self.result_table.indexAt(event.pos())
            row = index.row()
            if row != self.hovered_row:
                self.hovered_row = row
                self.update_row_highlight()
        return super().eventFilter(source, event)

    def update_row_highlight(self):
        verif_col = self.get_column_index("Verification")
        for row in range(self.result_table.rowCount()):
            is_selected = self.result_table.selectionModel().isRowSelected(row, QtCore.QModelIndex())
            is_hovered = (row == self.hovered_row)
            for col in range(self.result_table.columnCount()):
                item = self.result_table.item(row, col)
                if not item:
                    continue
                is_manual = False
                if verif_col != -1:
                    verif_item = self.result_table.item(row, verif_col)
                    if verif_item and verif_item.text().lower() == "manual":
                        is_manual = True
                if is_manual:
                    if is_selected:
                        item.setBackground(QColor("#fff1b0"))  
                    elif is_hovered:
                        item.setBackground(QColor("#FCF4AF"))  
                    else:
                        item.setBackground(QColor("#FFFACD"))  
                else:
                    if is_selected:
                        item.setBackground(QColor("#dddddd"))  
                    elif is_hovered:
                        item.setBackground(QColor("#eeeeee"))  
                    else:
                        item.setBackground(QColor("white"))

    def get_column_index(self, column_name):
        for col in range(self.result_table.columnCount()):
            header_item = self.result_table.horizontalHeaderItem(col)
            if header_item and header_item.text().strip().lower() == column_name.strip().lower():
                return col
        return -1

    def load_pdf(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select PDF Artwork", "", "PDF Files (*.pdf)")
        if path:
            self.pdf_path = path
            self.pdf_label.setText(f"PDF: {os.path.basename(path)}")
            self.pages = extract_text_by_page(self.pdf_path)
            self.product_infos = extract_product_info_by_page(self.pages)

    def load_excel(self):
        if not self.pdf_path:
            QtWidgets.QMessageBox.warning(self, "PDF Required", "Please upload a PDF file first.")
            return

        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select Checklist Excel", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.excel_path = path
            self.excel_label.setText(f"Checklist: {os.path.basename(path)}")
            try:
                self.checklist_df = load_checklist(path, os.path.basename(self.pdf_path))
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Checklist Load Error", str(e))

    def start_checking(self):
        if self.checklist_df is None or not self.pages:
            QtWidgets.QMessageBox.warning(self, "Missing File", "Please upload both Checklist and PDF before checking.")
            return

        extracted_text_list = []
        for page in self.pages:
            page_items = []
            for item in page:
                if isinstance(item, dict) and "text" in item:
                    page_items.append(item)
            extracted_text_list.append(page_items)

            # Check each term in the page
            self.product_infos = extract_product_info_by_page(extracted_text_list)

        self.result_df = start_check(self.checklist_df, extracted_text_list)

        if self.result_df.empty:
            QtWidgets.QMessageBox.information(self, "No Result", "No matching terms found.")
            return

        self.display_results(self.result_df)

    def display_results(self, df: pd.DataFrame):
        df_src = df.copy()
        symbol_cols_protect = {"Symbol/ Exact wording", "Symbol/Exact wording", "__Term_HTML__"}
        cols_to_fill = [c for c in df.columns if c not in symbol_cols_protect]
        df[cols_to_fill] = df[cols_to_fill].fillna("-")

        preferred = ["Requirement", "Symbol/ Exact wording", "Specification", 
                    "Package Panel", "Procedure", "Remark", "Found", "Match", 
                    "Font Size", "Pages", "Note", "Verification"]
        
        # helper column ที่ไม่แสดงใน UI แต่ยังแสดงใน df_scr
        helper_names = {"remark url", "remark link"}
        helper_cols = [ c for c in df.columns if c.strip().lower() in helper_names]

        # คอลัมน์ภายในที่ต้องซ่อนจาก UI
        internal_hide = {
            "__Term_HTML__", "Image_Groups_Resolved", "Image_Groups",
            "Image Path Resolved", "Image Path", "_HasImage", "Language List"
        }

        # จัดลำดับเฉพาะคอลัมน์ที่จะแสดง (ไม่รวม helper)
        ordered = [ c for c in preferred if c in df.columns]
        tail = [c for c in df.columns if c not in ordered + helper_cols]

        # ตัด internal ออกจากรายการแสดงผล
        tail = [c for c in tail if c not in internal_hide]
        df_ui = df.loc[:, ordered + tail]

        # ตั้งค่าตาราง
        self.result_table.setRowCount(len(df_ui))
        self.result_table.setColumnCount(len(df_ui.columns))
        self.result_table.setHorizontalHeaderLabels(df_ui.columns.tolist())

        symbol_names = {"Symbol/  Exact wording", "Symbol/ Exact wording", "Symbol/Exact wording"}
        sym_col = None
        for c in range(self.result_table.columnCount()):
            hi = self.result_table.horizontalHeaderItem(c)
            if hi and hi.text().strip() in symbol_names:
                sym_col = c
                break

        if sym_col is not None:
            hh = self.result_table.horizontalHeader()
            hh.setSectionResizeMode(sym_col, QHeaderView.Fixed) 
            self.result_table.setColumnWidth(sym_col, 320) 

        # Bold header font
        header_font = self.result_table.horizontalHeader().font()
        header_font.setBold(True)
        self.result_table.horizontalHeader().setFont(header_font)

        if hasattr(self, "update_row_highlight"):
            self.update_row_highlight()

        # ตั้งความกว้างหลัก
        equal_width = 240
        if "Requirement" in df_ui.columns:
            self.result_table.setColumnWidth(df_ui.columns.get_loc("Requirement"), equal_width)
        if "Specification" in df_ui.columns:
            self.result_table.setColumnWidth(df_ui.columns.get_loc("Specification"), equal_width)
        if "Symbol/ Exact wording" in df_ui.columns:
            self.result_table.setColumnWidth(df_ui.columns.get_loc("Symbol/ Exact wording"), 380)
        if "Package Panel" in df_ui.columns:
            self.result_table.setColumnWidth(df_ui.columns.get_loc("Package Panel"), equal_width)
        if "Procedure" in df_ui.columns:
            self.result_table.setColumnWidth(df_ui.columns.get_loc("Procedure"), equal_width)
        if "Remark" in df_ui.columns:
            self.result_table.setColumnWidth(df_ui.columns.get_loc("Remark"), 280)

        header = self.result_table.horizontalHeader()
        for i in range(df_ui.shape[1]):
            header.setSectionResizeMode(i, QtWidgets.QHeaderView.Fixed)

        try:
            if "Pages" in df_ui.columns:
                page_index = df_ui.columns.get_loc("Pages")
                ref_width = self.result_table.columnWidth(page_index)
            else:
                ref_width = 120
        except ValueError:
            ref_width = 120

        for col in ["Found", "Match", "Font Size", "Note", "Verification"]:
            if col in df_ui.columns:
                        self.result_table.setColumnWidth(df_ui.columns.get_loc(col), ref_width)

            if "Note" in df_ui.columns:
                self.result_table.setColumnWidth(df_ui.columns.get_loc("Note"), 250)

        self.result_table.resizeRowsToContents()

        # helpers
        def linkify(text: str) -> str:
            if not isinstance(text, str) or not text:
                return "-"
            return re.sub(r'(https?://[^\s]+)', r'<a href="\1">\1</a>', text)

        # หา image path จาก self.checklist_df ด้วย (Requirement, Symbol/Exact wording)
        def _lookup_image_groups(requirement_text: str, term_text: str):
            if getattr(self, "checklist_df", None) is None:
                return []
            req_series  = self.checklist_df.get("Requirement")
            term_series = self.checklist_df.get("Symbol/Exact wording")
            if req_series is None:
                return []

            req_left = req_series.astype(str).str.strip().str.lower()
            groups = []

            if term_series is not None:
                term_left = term_series.astype(str).str.strip().str.lower()
                if term_text and term_text.strip() and term_text.strip() != "-":
                    mask_rt = (req_left == str(requirement_text).strip().lower()) & \
                            (term_left == str(term_text).strip().lower())
                    sub = self.checklist_df[mask_rt]
                    if not sub.empty:
                        g = sub.iloc[0].get("Image_Groups_Resolved") or sub.iloc[0].get("Image_Groups", [])
                        groups = g or []

            # ถ้าไม่เจอ และ term_text เป็น "-" ให้ fallback จับคู่ด้วย Requirement อย่างเดียว
            if not groups:
                mask_r = (req_left == str(requirement_text).strip().lower())
                sub_r = self.checklist_df[mask_r]
                if not sub_r.empty:
                    for _, r in sub_r.iterrows():
                        g = r.get("Image_Groups_Resolved") or r.get("Image_Groups", [])
                        if g:
                            groups = g
                            break
            return groups or []

        # แคชภาพ
        if not hasattr(self, "_image_cache"):
            self._image_cache = {}

        for row_idx in range(len(df_ui)):
            row_ui  = df_ui.iloc[row_idx]   
            row_src = df_src.iloc[row_idx]     

            found        = str(row_ui.get("Found", ""))
            match        = str(row_ui.get("Match", ""))
            font_size    = str(row_ui.get("Font Size", ""))
            note         = str(row_ui.get("Note", ""))
            verification = str(row_ui.get("Verification", "")).strip().lower()

            for col_idx, header in enumerate(df_ui.columns):
                value = row_ui.get(header, "-")

                if header == "Symbol/ Exact wording":
                    req_text  = str(row_src.get("Requirement", "")).strip()

                    def _clean_plain(s: str) -> str:
                        if s is None:
                            return ""
                        s = str(s).strip()
                        return "" if s.lower() in ("nan", "none", "-") else s

                    # กลุ่มรูปของแถวนี้
                    groups = row_src.get("Image_Groups_Resolved") or row_src.get("Image_Groups") or []
                    has_images = bool(groups and any(g.get("paths") for g in groups))

                    # plain text (จาก df_ui)
                    term_raw  = row_ui.get(header, "")
                    term_text = _clean_plain(term_raw)

                    # html (underline/bold) จาก excel
                    def _clean_html(s: str) -> str:
                        if not isinstance(s, str):
                            return ""
                        s2 = s.strip()
                        if s2.lower() in ("nan", "none", "-"):
                            return ""
                        plain = re.sub(r"<[^>]+>", "", s2).strip()
                        return "" if plain == "" else s2

                    html_val = ""
                    for k in ("__Term_HTML__", "Term_Underline_HTML"):
                        v = row_src.get(k, "")
                        if isinstance(v, str) and v.strip():
                            html_val = _clean_html(v)
                            if html_val:
                                break

                    # ตัดสินใจข้อความที่จะแสดง
                    if html_val:
                        display_text = html_val
                        plain_for_measure = re.sub(r"<[^>]+>", "", html_val)
                    elif term_text:
                        display_text = term_text
                        plain_for_measure = term_text
                    else:
                        display_text = "" if has_images else "-"
                        plain_for_measure = ""

                    # สร้าง UI ของเซลล์
                    container = QtWidgets.QWidget()
                    outer = QtWidgets.QVBoxLayout(container)
                    outer.setContentsMargins(4, 2, 4, 2)
                    outer.setSpacing(4)
                    outer.setAlignment(QtCore.Qt.AlignCenter)

                    # วางข้อความเฉพาะเมื่อมีข้อความจริงเท่านั้น
                    term_label = None
                    if display_text.strip():
                        term_label = QtWidgets.QLabel()
                        term_label.setTextFormat(QtCore.Qt.RichText)
                        term_label.setTextInteractionFlags(Qt.TextBrowserInteraction)
                        term_label.setAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                        # ตัดบรรทัดแบบยืดหยุ่น + คุมความกว้างตามคอลัมน์
                        col_width = self.result_table.columnWidth(col_idx)
                        fm = term_label.fontMetrics()
                        inner_w = max(40, col_width - 12)
                        term_label.setMinimumWidth(inner_w)
                        term_label.setMaximumWidth(inner_w)

                        if plain_for_measure:
                            need_wrap = fm.horizontalAdvance(plain_for_measure) > inner_w
                            term_label.setWordWrap(need_wrap)
                        else:
                            term_label.setWordWrap(True)

                        term_label.setText(display_text)

                        # สีแดงกรณี Not Found (คงพฤติกรรมเดิม)
                        if str(row_ui.get("Found", "")).startswith("❌"):
                            term_label.setStyleSheet("color:#b91c1c;")
                        else:
                            term_label.setStyleSheet("")

                        outer.addWidget(term_label, 0, QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                    # ส่วนรูปภาพ
                    if has_images:
                        all_paths = []
                        for g in groups:
                            all_paths.extend(g.get("paths", []))

                        if not hasattr(self, "_image_cache"):
                            self._image_cache = {}

                        if all_paths:
                            img_wrap = QtWidgets.QWidget()
                            img_vbox = QtWidgets.QVBoxLayout(img_wrap)
                            padding_lr = 8
                            img_vbox.setContentsMargins(padding_lr, 0, padding_lr, 0)  # กันภาพไม่ชิดขอบ
                            img_vbox.setSpacing(8)
                            img_vbox.setAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                            # ความกว้างรูปสูงสุด = ความกว้างคอลัมน์ - padding ซ้าย/ขวา
                            col_width = self.result_table.columnWidth(col_idx)
                            max_img_w = max(40, col_width - 2 * padding_lr)

                            def _is_logo(path: str, req: str) -> bool:
                                p_low = os.path.basename(path or "").lower()
                                r_low = (req or "").lower()
                                keys = ("logo", "mark", "lion", "ce ", "ukca", "mc ")
                                return any(k in p_low for k in keys) or any(k in r_low for k in keys)

                            for p in all_paths:
                                if not p:
                                    continue

                                pm = self._image_cache.get(p)
                                if pm is None:
                                    qpm = QtGui.QPixmap(p)
                                    pm = qpm if not qpm.isNull() else None
                                    self._image_cache[p] = pm if pm else QtGui.QPixmap()

                                lbl = QtWidgets.QLabel()
                                lbl.setAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                                if not pm:
                                    lbl.setText(f"[!] Missing image: {p}")
                                    lbl.setStyleSheet("color:#b91c1c;")
                                    img_vbox.addWidget(lbl, 0, QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                                    continue

                                # โลโก้ = เล็ก / non-logo = กว้าง (เผื่อ margin)
                                if _is_logo(p, req_text):
                                    target_w = min(140, max_img_w)
                                else:
                                    target_w = min(int(max_img_w * 0.96), pm.width())

                                # สเกลเฉพาะจำเป็น (ไม่ขยายเกินไฟล์จริง)
                                if pm.width() > target_w:
                                    lbl.setPixmap(pm.scaledToWidth(target_w, QtCore.Qt.SmoothTransformation))
                                else:
                                    lbl.setPixmap(pm)

                                img_vbox.addWidget(lbl, 0, QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                            # ถ้า "ไม่มีข้อความ" → ใส่ stretch ก่อน/หลังภาพให้ลอยกึ่งกลางแนวตั้งจริง
                            if not display_text.strip():
                                outer.addStretch(1)
                                outer.addWidget(img_wrap, 0, QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                                outer.addStretch(1)
                            else:
                                outer.addWidget(img_wrap, 0, QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                    container.setLayout(outer)
                    self.result_table.takeItem(row_idx, col_idx)
                    self.result_table.setCellWidget(row_idx, col_idx, container)

                    # ปรับความสูงแถวให้พอดีเนื้อหา
                    self.result_table.resizeRowToContents(row_idx)
                    if self.result_table.rowHeight(row_idx) < 28:
                        self.result_table.setRowHeight(row_idx, 28)

                    continue

                # Remark คอลัมน์เดียว แต่คลิกได้ถ้ามีลิงก์จาก Excel; ถ้าไม่มีให้ linkify/หรือ "-" ; จัดกึ่งกลาง 
                if header == "Remark":
                    url = str(row_src.get("Remark URL", "") or row_src.get("Remark Link", "") or "").strip()
                    txt = (str(value) if value is not None else "").strip()

                    lbl = QtWidgets.QLabel()
                    lbl.setTextFormat(QtCore.Qt.RichText)
                    lbl.setWordWrap(True)
                    lbl.setOpenExternalLinks(True)
                    lbl.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)

                    if url and txt and txt not in ["-", "nan", "None"]:
                        lbl.setText(f'<a href="{url}">{QtCore.QCoreApplication.translate("", txt)}</a>')
                        lbl.setToolTip(url)
                    elif url and not txt:
                        lbl.setText(f'<a href="{url}">{url}</a>')
                        lbl.setToolTip(url)
                    else:
                        def linkify(s: str) -> str:
                            if not s:
                                return "-"
                            return re.sub(r'(https?://[^\s]+)', r'<a href="\1">\1</a>', s)
                        lbl.setText(linkify(txt) if txt else "-")

                    lbl.setAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    self.result_table.setCellWidget(row_idx, col_idx, lbl)
                    continue

                # คอลัมน์อื่นๆ 
                item = QtWidgets.QTableWidgetItem(str(value))
                item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                item.setToolTip(str(value))
                item.setText(str(value))

                # Alignment: Requirement = ซ้าย/หนา, อื่นๆ = กึ่งกลาง
                if header == "Requirement":
                    item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                    f = item.font(); f.setBold(True); item.setFont(f)
                else:
                    item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                
                # สีตามสถานะ
                if verification == "manual":
                    item.setBackground(QColor("#fff9cc"))
                    if header in ["Found", "Match", "Font Size", "Note"]:
                        item.setForeground(QColor("gray"))
                else:
                    # ❌ ถ้า Found เป็น Not Found → คอลัมน์ทั้งหมดแดง ยกเว้น Requirement
                    if found.startswith("❌") and header != "Requirement":
                        item.setForeground(QColor("red"))

                    # Logic เดิมอื่น ๆ (ยังคงไว้)
                    elif header == "Match" and match.startswith("❌"):
                        item.setForeground(QColor("red"))
                    elif header == "Font Size" and not font_size.startswith("✔"):
                        item.setForeground(QColor("red"))
                    elif header == "Note" and note.strip() not in ["-", ""]:
                        item.setForeground(QColor("red"))

                self.result_table.setItem(row_idx, col_idx, item)

            self.result_table.resizeRowsToContents()

    def _on_column_resized(self, logicalIndex: int, oldSize: int, newSize: int):
        header_item = self.result_table.horizontalHeaderItem(logicalIndex)
        if not header_item:
            return
        if header_item.text().strip().lower() != "symbol/ exact wording":
            return

        col_idx = logicalIndex
        col_w = self.result_table.columnWidth(col_idx)
        inner_w = max(40, col_w - 12)

        for row in range(self.result_table.rowCount()):
            cell = self.result_table.cellWidget(row, col_idx)
            if not cell:
                continue
            labels = cell.findChildren(QtWidgets.QLabel)
            for lbl in labels:
                # ข้าม QLabel ที่เป็นภาพ (ไม่มี RichText)
                txt_html = lbl.text() if hasattr(lbl, "text") else ""
                if not isinstance(txt_html, str) or "<img" in txt_html.lower():
                    continue
                txt_plain = re.sub(r"<[^>]+>", "", txt_html)
                fm = lbl.fontMetrics()
                lbl.setMinimumWidth(inner_w)
                lbl.setMaximumWidth(inner_w)
                need_wrap = fm.horizontalAdvance(txt_plain) > inner_w if txt_plain else True
                lbl.setWordWrap(need_wrap)

        self.result_table.resizeRowsToContents()

    def export_results(self):
        if self.result_df is None or self.result_df.empty:
            QtWidgets.QMessageBox.warning(self, "No Results", "Please run checking before exporting.")
            return
        export_result_to_excel(self.result_df)

    def preview_pdf(self):
        if self.pdf_path:
            term = self.search_input.text().strip()
            if not term:
                QtWidgets.QMessageBox.information(self, "Enter Term", "Please enter a term to highlight in PDF.")
                return
            viewer = PDFViewer(self.pdf_path, search_term=term, product_infos=getattr(self, "product_infos", []))
            viewer.show()
        else:
            QtWidgets.QMessageBox.warning(self, "No PDF", "Please upload a PDF first.")

    def search_text(self):
        query = self.search_input.text().strip()
        if not query:
            QtWidgets.QMessageBox.information(self, "Search", "Please enter a term to search.")
            return

        if not self.result_table or self.result_table.rowCount() == 0:
            QtWidgets.QMessageBox.information(self, "Search", "No results to search.")
            return

        q = query.lower()

        # helper: ตัดแท็ก HTML ออกจาก QLabel (RichText)
        def _plain_text_from_html(s: str) -> str:
            if not isinstance(s, str):
                return ""
            return re.sub(r"<[^>]+>", "", s).strip()

        # ค้นหาทั้ง QTableWidgetItem และ cellWidget(QLabel)
        for row in range(self.result_table.rowCount()):
            for col in range(self.result_table.columnCount()):
                hit = False

                item = self.result_table.item(row, col)
                if item:
                    txt = item.text() or ""
                    if q in txt.lower():
                        hit = True

                if not hit:
                    w = self.result_table.cellWidget(row, col)
                    if w:

                        labels = w.findChildren(QtWidgets.QLabel)
                        for lbl in labels:
                            txt = _plain_text_from_html(lbl.text()).lower()
                            if q in txt:
                                hit = True
                                break
                            
                if hit:
                    self.result_table.setCurrentCell(row, col)
                    index = self.result_table.model().index(row, col)
                    self.result_table.scrollTo(index, QtWidgets.QAbstractItemView.PositionAtCenter)
                    return

        QtWidgets.QMessageBox.information(self, "Search", f"'{query}' not found in results.")
