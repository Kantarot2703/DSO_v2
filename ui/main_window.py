import os, re
import pandas as pd
import logging
from PyQt5.QtWidgets import QWidget, QLabel, QHBoxLayout, QVBoxLayout, QGridLayout
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
        df.fillna("-", inplace=True)
        # บังคับคอลัมน์แกนหลักให้มีเสมอ
        base_cols = ["Requirement","Symbol/ Exact wording","Specification",
                    "Found","Match","Font Size","Pages","Verification",
                    "Note","Remark","Remark URL"]
        tail_cols = [c for c in df.columns if c not in base_cols]
        df = df.reindex(columns=base_cols + tail_cols, fill_value="-")
        self.result_table.setRowCount(len(df))
        self.result_table.setColumnCount(len(df.columns))
        self.result_table.setHorizontalHeaderLabels(df.columns.tolist())

        # Bold header font
        header_font = self.result_table.horizontalHeader().font()
        header_font.setBold(True)
        self.result_table.horizontalHeader().setFont(header_font)

        self.update_row_highlight()

        # ซ่อนคอลัมน์ Remark URL (ใช้เป็นแหล่ง URL ให้ Remark เท่านั้น)
        if "Remark URL" in df.columns:
            url_idx = df.columns.get_loc("Remark URL")
            self.result_table.setColumnHidden(url_idx, True)

        equal_width = 280 

        if "Requirement" in df.columns:
            req_index = df.columns.get_loc("Requirement")
            self.result_table.setColumnWidth(req_index, equal_width)

        if "Specification" in df.columns:
            spec_index = df.columns.get_loc("Specification")
            self.result_table.setColumnWidth(spec_index, equal_width)

        if "Symbol/ Exact wording" in df.columns:
            term_index = df.columns.get_loc("Symbol/ Exact wording")
            self.result_table.setColumnWidth(term_index, 700)

        # ให้ Remark กว้างพอ ๆ กับ Spec
        if "Remark" in df.columns:
            remark_index = df.columns.get_loc("Remark")
            self.result_table.setColumnWidth(remark_index, equal_width)

        header = self.result_table.horizontalHeader()
        for i in range(df.shape[1]):
            header.setSectionResizeMode(i, QtWidgets.QHeaderView.Fixed)

        try:
            if "Pages" in df.columns:
                page_index = df.columns.get_loc("Pages")
                ref_width = self.result_table.columnWidth(page_index)
        except ValueError:
            ref_width = 120

        fixed_columns = ["Found", "Match", "Font Size", "Note", "Verification"]
        for col in fixed_columns:
            try:
                idx = df.columns.get_loc(col)
                self.result_table.setColumnWidth(idx, ref_width)
            except ValueError:
                pass

        self.result_table.resizeRowsToContents()

        def linkify(text: str) -> str:
            if not isinstance(text, str):
                return "-"
            url_pattern = r'(https?://[^\s]+)'
            return re.sub(url_pattern, r'<a href="\1">\1</a>', text)

        # helper: หา image path จาก self.checklist_df ด้วย (Requirement, Symbol/Exact wording)
        def _lookup_image_groups(requirement_text: str, term_text: str):
            if self.checklist_df is None:
                return []
            req_series  = self.checklist_df.get("Requirement")
            term_series = self.checklist_df.get("Symbol/Exact wording")
            if req_series is None:
                return []

            req_left = req_series.astype(str).str.strip().str.lower()

            # จับคู่ปกติ Req + Term (ถ้า term_text ไม่ใช่ "-" และไม่ว่าง)
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
                    # เลือกแถวแรกที่มีภาพจริง
                    for _, r in sub_r.iterrows():
                        g = r.get("Image_Groups_Resolved") or r.get("Image_Groups", [])
                        if g:
                            groups = g
                            break
            return groups or []

        # แคชภาพ
        if not hasattr(self, "_image_cache"):
            self._image_cache = {}

        for row_idx, (_, row) in enumerate(df.iterrows()):
            found = str(row.get("Found", ""))
            match = str(row.get("Match", ""))
            font_size = str(row.get("Font Size", ""))
            note = str(row.get("Note", ""))
            verification = str(row.get("Verification", "")).strip().lower()

            for col_idx, (header, value) in enumerate(row.items()):
                # --- BEGIN: render grouped images in Term (Manual only) ---
                if header == "Symbol/ Exact wording":
                    req_text  = str(row.get("Requirement", "")).strip()
                    term_text = str(row.get("Symbol/ Exact wording", "")).strip()
                    groups    = _lookup_image_groups(req_text, term_text)

                    if groups:
                        container = QtWidgets.QWidget()
                        outer = QtWidgets.QVBoxLayout(container)
                        outer.setContentsMargins(6, 2, 6, 2); outer.setSpacing(6)

                        # แสดงกฎของกลุ่มแรกแบบสั้น (ANY/ALL)
                        mode = (groups[0].get("mode") or "any").lower()
                        rule = "(Rule: Need BOTH)" if mode == "all" else "(Rule: Choose ONE)"
                        head = QtWidgets.QLabel(f"{term_text}  {rule}")
                        head.setStyleSheet("font-weight:600;color:#333;")
                        outer.addWidget(head)

                        # วาดเป็นกริด 2 รูป/แถว
                        grid = QtWidgets.QGridLayout(); grid.setContentsMargins(0,0,0,0)
                        grid.setHorizontalSpacing(10); grid.setVerticalSpacing(10)

                        # รวมพาธทั้งหมดจากทุกกลุ่ม (ถ้าต้องการแยกกลุ่ม ให้ลูปทีละ g)
                        paths = []
                        for g in groups:
                            paths.extend(g.get("paths", []))

                        for i, p in enumerate(paths):
                            if not p: 
                                continue
                            # cache
                            pm = self._image_cache.get(p)
                            if pm is None:
                                qpm = QtGui.QPixmap(p)
                                pm = qpm if not qpm.isNull() else None
                                self._image_cache[p] = pm if pm else QtGui.QPixmap()
                            if not pm:
                                miss = QtWidgets.QLabel(f"[!] Missing image: {p}")
                                miss.setStyleSheet("color:#b91c1c;")
                                grid.addWidget(miss, i//2, i%2)
                            else:
                                lbl = QtWidgets.QLabel()
                                lbl.setPixmap(pm.scaledToWidth(220, QtCore.Qt.SmoothTransformation))
                                lbl.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)
                                grid.addWidget(lbl, i//2, i%2)

                        outer.addLayout(grid)
                        self.result_table.setCellWidget(row_idx, col_idx, container)
                        continue
                
                # Add clickable Remark links
                if header == "Remark":
                    url = str(row.get("Remark URL", "")).strip()
                    txt = (str(value) if value is not None else "").strip()
                    lbl = QtWidgets.QLabel()
                    lbl.setTextFormat(QtCore.Qt.RichText)
                    lbl.setWordWrap(True)
                    lbl.setOpenExternalLinks(True)
                    lbl.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)

                    if url and txt and txt not in ["-", "nan", "None"]:
                        # ถ้ามี URL จาก Excel + ข้อความ → ใช้ URL เป็น anchor ของทั้งข้อความ
                        lbl.setText(f'<a href="{url}">{QtCore.QCoreApplication.translate("", txt)}</a>')
                        lbl.setToolTip(url)
                    elif url and not txt:
                        lbl.setText(f'<a href="{url}">{url}</a>')
                        lbl.setToolTip(url)
                    else:
                        # ไม่มี hyperlink จาก Excel → linkify ข้อความเอง
                        lbl.setText(linkify(txt) if txt else "-")

                    self.result_table.setCellWidget(row_idx, col_idx, lbl)
                    continue

                item = QtWidgets.QTableWidgetItem(str(value))
                item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                item.setToolTip(str(value))
                item.setText(str(value))

                # Alignment and styling based on header
                if header == "Requirement":
                    item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                elif header in ["Symbol/ Exact wording", "Specification", "Remark"]:
                    item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                else:
                    item.setTextAlignment(QtCore.Qt.AlignCenter | QtCore.Qt.AlignVCenter)

                if verification == "manual":
                    item.setBackground(QColor("#fff9cc"))
                    if header in ["Found", "Match", "Font Size", "Note"]:
                        item.setForeground(QColor("gray"))

                elif header == "Symbol/ Exact wording" and found.startswith("❌"):
                    item.setForeground(QColor("red"))
                elif header == "Match" and match.startswith("❌"):
                    item.setForeground(QColor("red"))
                elif header == "Font Size" and not font_size.startswith("✔"):
                    item.setForeground(QColor("red"))
                elif header == "Note" and note.strip() not in ["-", ""]:
                    item.setForeground(QColor("red"))

                self.result_table.setItem(row_idx, col_idx, item)

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

        found = False
        for row in range(self.result_table.rowCount()):
            for col in range(self.result_table.columnCount()):
                item = self.result_table.item(row, col)
                if item and query.lower() in item.text().lower():
                    self.result_table.setCurrentCell(row, col)
                    self.result_table.scrollToItem(item)
                    found = True
                    return

        if not found:
            QtWidgets.QMessageBox.information(self, "Search", f"'{query}' not found in results.")