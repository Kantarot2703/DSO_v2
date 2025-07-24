import os
import re
import pandas as pd
from PyQt5 import QtWidgets
from PyQt5 import QtGui
from PyQt5 import QtCore
from ui.pdf_viewer import PDFViewer
from checklist_loader import load_checklist, start_check
from pdf_reader import extract_text_by_page
from checker import check_term_in_page
from result_exporter import export_result_to_excel
from PyQt5.QtGui import QColor
from pdf_reader import extract_product_info_by_page


class DSOApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DSO - Digital Sign Out Checker")
        self.setGeometry(100, 100, 1000, 700)
        self.excel_path = ""
        self.pdf_path = ""
        self.checklist_df = None
        self.pages = None
        self.result_df = None
        self.init_ui()
        self.product_infos = []

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
        self.result_table.setColumnCount(0)

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

            # ดึงข้อมูล Product ต่อหน้า
            self.product_infos = extract_product_info_by_page(extracted_text_list)

        self.result_df = start_check(self.checklist_df, extracted_text_list)

        if self.result_df.empty:
            QtWidgets.QMessageBox.information(self, "No Result", "No matching terms found.")
            return

        self.display_results(self.result_df)

    def display_results(self, df: pd.DataFrame):
        from PyQt5.QtGui import QColor
        df.fillna("-", inplace=True)
        self.result_table.setRowCount(len(df))
        self.result_table.setColumnCount(len(df.columns))
        self.result_table.setHorizontalHeaderLabels(df.columns.tolist())
        self.result_table.horizontalHeader().setStretchLastSection(True)

        self.result_table.resizeRowsToContents()
        self.result_table.resizeColumnsToContents()  # เรียกก่อน setColumnWidth

        try:
            found_col_index = df.columns.get_loc("Found")
            self.result_table.setColumnWidth(found_col_index, 110)
        except ValueError:
            pass

        for row_idx, (_, row) in enumerate(df.iterrows()):
            found = str(row.get("Found", ""))
            match = str(row.get("Match", ""))
            font_size = str(row.get("Font Size", ""))
            note = str(row.get("Note", ""))
            term = str(row.get("Term", ""))
            verification = str(row.get("Verification", "")).strip().lower()

            for col_idx, (header, value) in enumerate(row.items()):
                item = QtWidgets.QTableWidgetItem(str(value))
                item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                item.setToolTip(str(value))
                item.setText(str(value))

                if header in ["Requirement", "Term", "Specification"]:
                    item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                else:
                    item.setTextAlignment(QtCore.Qt.AlignCenter | QtCore.Qt.AlignVCenter)

                if verification == "manual":
                    if header in ["Found", "Match", "Font Size", "Note"]:
                        item.setForeground(QColor("gray"))
                elif header == "Term" and found.startswith("❌"):
                    item.setForeground(QColor("red"))
                elif header == "Match" and match.startswith("❌"):
                    item.setForeground(QColor("red"))
                elif header == "Font Size" and not font_size.startswith("✔"):
                    item.setForeground(QColor("red"))
                elif header == "Note" and note.strip() not in ["-", ""]:
                    item.setForeground(QColor("red"))

                self.result_table.setItem(row_idx, col_idx, item)

        # ปรับขนาดคอลัมน์: 3 คอลัมน์แรกตามเนื้อหา, ที่เหลือกว้างเท่ากัน
        fixed_width = 110 
        for col in range(self.result_table.columnCount()):
            header = df.columns[col]
            if header in ["Requirement", "Term", "Specification"]:
                self.result_table.resizeColumnToContents(col)
            else:
                self.result_table.setColumnWidth(col, fixed_width)

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


    def extract_product_info_by_page(pages):
        product_infos = []
        for page_number, page_items in enumerate(pages, start=1):
            product_name = "-"
            part_no = "-"
            rev = "-"

            for item in page_items:
                text = item.get("text", "")
                size = float(item.get("size", 0))

                # หา Product Name: ใช้ขนาด ≥ 1.6 mm
                if size >= 1.6 and product_name == "-":
                    product_name = text

                # หา Part No (เช่น 4LB45-MF4A)
                if re.match(r"[A-Z0-9]{2,}-[A-Z0-9]{2,}", text) and part_no == "-":
                    part_no = text

                # หา Revision เช่น Rev: A1, Revision B
                if re.search(r"rev(ision)?\s*[:\-]?\s*[A-Z0-9]{1,2}", text, re.IGNORECASE):
                    rev_match = re.findall(r"[A-Z0-9]{1,2}$", text)
                    if rev_match:
                        rev = rev_match[0]

            product_infos.append({
                "page": page_number,
                "product_name": product_name,
                "part_no": part_no,
                "rev": rev
            })

        return product_infos